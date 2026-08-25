#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Smart Unity Addressables bundle downloader.

功能：
1. 启动时自动扫描映射 CSV、catalog.json/catalog.txt、catalog.bundle。
2. 若扫描到 catalog，询问是否在线更新。
3. 若没有 CSV，也没有可用 catalog，则自动下载：
   https://d3mya90gbacu0m.cloudfront.net/prod/StreamingAssets/aa/catalog.bundle
4. catalog.bundle 支持：
   - 明文 JSON（即使扩展名是 .bundle）
   - gzip / zip
   - UnityFS AssetBundle（内置 LZ4 解压并从 TextAsset 数据中抽取 catalog JSON）
5. catalog 分析完成后可生成 bundle_id_mapping.csv；下次直接用 CSV 筛选会更快。
6. 查询支持：
   - 精确 ID，例如 ch_1041005
   - 路径/分类关键词，例如 still、sister、character
   - 多个条件用逗号、分号或空格分隔
   - 强制模式：id:ch_1041005 / key:still
7. 同一查询命中多个 bundle 时全部下载。
8. 多线程下载、总体进度、断点续传、失败重试、已存在跳过。
9. 下载后的命名：
   ID_Category_Version.bundle
   若重名则自动追加 __02、__03 ...
10. 所有 bundle 统一保存到一个目录。
11. 下载完成后可选择 Export，并选择三种目录布局：原 Assets 路径 / ID+类型 / ID 平铺。
12. Sprite 不再按 SpriteRect 裁切；改为导出其底层完整 Texture2D，Texture2D 也始终导出完整 PNG。
13. TextAsset 特殊处理：.atlas.txt -> .atlas，.skel.bytes -> .skel。
14. AudioClip 优先导出真实 OGG/WAV/M4A/MP3；Font/Mesh/Shader/Texture2DArray 等增加专用导出。
15. 同时扫描主 bundle 中没有 container 路径的常见内容对象，减少声音和其他资源漏导。
16. .asset / .prefab / .mat 等其他 container 对象继续保存原始序列化数据，不自动追加额外后缀。
17. Export 前检查 UnityPy / Pillow 依赖，并检查已知 bundle 依赖是否下载齐全。
18. Windows 状态文件不再使用高频 os.replace，避免读写竞争导致 WinError 5。
19. env.container 中的 PPtr 会先 deref 成 ObjectReader，修复 .asset/.mat raw 导出失败。
20. Export 完成/失败后自动清理 .export_work、状态 JSON、临时 JSON 和无用 .part 文件。
"""
from __future__ import annotations

import argparse
import atexit
import base64
import csv
import gzip
import io
import importlib
import importlib.util
import json
import lzma
import mmap
import os
import posixpath
import re
import shutil
import struct
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from array import array
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path, PurePosixPath
from typing import Dict, Iterable, List, Sequence, Set, Tuple

CATALOG_BUNDLE_URL = "https://d3mya90gbacu0m.cloudfront.net/prod/StreamingAssets/aa/catalog.bundle"
DEFAULT_BUNDLE_BASE_URL = "https://d3mya90gbacu0m.cloudfront.net/prod/StreamingAssets/aa/WebGL/"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AddressablesSmartDownloader/6.0"

ENTRY_SIZE = 28
ENTRY_STRUCT = struct.Struct("<7i")

CSV_FIELDS = [
    "bundle", "id", "version", "category", "asset_root", "asset_count",
    "confidence", "max_dependency_bundle_count", "dependency_bundles",
    "bundle_internal_id", "bundle_location_index",
]


class CatalogFormatError(RuntimeError):
    pass


# ---------------------------------------------------------------------------
# 通用辅助
# ---------------------------------------------------------------------------

def human_size(n: int | None) -> str:
    if n is None:
        return "?"
    f = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if f < 1024 or unit == "TB":
            return f"{int(f)}B" if unit == "B" else f"{f:.1f}{unit}"
        f /= 1024
    return str(n)


def ask_yes_no(prompt: str, default: bool = True) -> bool:
    suffix = " [Y/n]: " if default else " [y/N]: "
    while True:
        try:
            s = input(prompt + suffix).strip().lower()
        except EOFError:
            return default
        if not s:
            return default
        if s in {"y", "yes", "1", "是", "好", "需要"}:
            return True
        if s in {"n", "no", "0", "否", "不", "不要"}:
            return False
        print("请输入 Y 或 N。")


def safe_piece(value: str, fallback: str) -> str:
    s = (value or "").strip().replace(" ", "_")
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    s = re.sub(r"_+", "_", s).strip(" ._")
    return s or fallback


def split_queries(text: str) -> list[str]:
    # 保留 id:xxx / key:xxx；普通输入允许逗号、分号、空格分隔
    return [x.strip() for x in re.split(r"[,;\s]+", text) if x.strip()]


def unique_paths(paths: Iterable[Path]) -> list[Path]:
    out = []
    seen = set()
    for p in paths:
        try:
            rp = p.resolve()
        except Exception:
            rp = p
        key = str(rp).casefold()
        if key not in seen:
            seen.add(key)
            out.append(rp)
    return out


# ---------------------------------------------------------------------------
# 自动扫描 CSV / catalog
# ---------------------------------------------------------------------------

def _scan_roots() -> list[Path]:
    roots = [Path.cwd(), Path(__file__).resolve().parent]
    return unique_paths([p for p in roots if p.exists()])


def looks_like_mapping_csv(path: Path) -> bool:
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            header = next(csv.reader(f), [])
        cols = {x.strip().casefold() for x in header}
        return {"bundle", "id", "version", "category"}.issubset(cols)
    except Exception:
        return False


def find_csv_candidates() -> list[Path]:
    preferred = ["bundle_id_mapping.csv", "bundle_id_mapping_compact.csv"]
    candidates: list[Path] = []
    for root in _scan_roots():
        for name in preferred:
            p = root / name
            if p.is_file() and looks_like_mapping_csv(p):
                candidates.append(p)
        for p in root.glob("*.csv"):
            if p.is_file() and looks_like_mapping_csv(p):
                candidates.append(p)
    candidates = unique_paths(candidates)
    candidates.sort(
        key=lambda p: (
            2 if p.name.lower() == "bundle_id_mapping.csv" else
            1 if p.name.lower() == "bundle_id_mapping_compact.csv" else 0,
            p.stat().st_mtime if p.exists() else 0,
        ),
        reverse=True,
    )
    return candidates


def find_catalog_text_candidates() -> list[Path]:
    candidates: list[Path] = []
    for root in _scan_roots():
        for pattern in ("catalog*.json", "catalog*.txt"):
            for p in root.glob(pattern):
                if p.is_file():
                    candidates.append(p)
    candidates = unique_paths(candidates)
    candidates.sort(
        key=lambda p: (
            2 if p.name.lower() == "catalog.json" else
            1 if p.name.lower() == "catalog.txt" else 0,
            p.stat().st_mtime if p.exists() else 0,
        ),
        reverse=True,
    )
    return candidates


def find_catalog_bundle_candidates() -> list[Path]:
    candidates: list[Path] = []
    for root in _scan_roots():
        for pattern in ("catalog.bundle", "catalog_*.bundle", "catalog-*.bundle"):
            for p in root.glob(pattern):
                if p.is_file():
                    candidates.append(p)
    candidates = unique_paths(candidates)
    candidates.sort(key=lambda p: p.stat().st_mtime if p.exists() else 0, reverse=True)
    return candidates


# ---------------------------------------------------------------------------
# HTTP 下载
# ---------------------------------------------------------------------------

def download_single_with_progress(
    url: str,
    dest: Path,
    *,
    timeout: int = 60,
    retries: int = 3,
    force: bool = True,
    label: str = "下载",
) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and not force:
        return dest

    part = dest.with_name(dest.name + ".part")
    last_error: Exception | None = None

    for attempt in range(1, retries + 1):
        existing = part.stat().st_size if part.exists() else 0
        headers = {"User-Agent": USER_AGENT, "Accept-Encoding": "identity"}
        if existing:
            headers["Range"] = f"bytes={existing}-"

        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                status = getattr(resp, "status", resp.getcode())
                clen = resp.headers.get("Content-Length")
                remain = int(clen) if clen and clen.isdigit() else None

                if existing and status == 206:
                    mode = "ab"
                    done = existing
                    total = existing + remain if remain is not None else None
                else:
                    mode = "wb"
                    done = 0
                    total = remain

                last_print = 0.0
                with part.open(mode) as f:
                    while True:
                        chunk = resp.read(1024 * 1024)
                        if not chunk:
                            break
                        f.write(chunk)
                        done += len(chunk)
                        now = time.monotonic()
                        if now - last_print >= 0.25:
                            if total:
                                pct = done * 100.0 / total
                                print(
                                    f"\r[{label}] {dest.name}: "
                                    f"{human_size(done)}/{human_size(total)} {pct:6.2f}%",
                                    end="",
                                    flush=True,
                                )
                            else:
                                print(
                                    f"\r[{label}] {dest.name}: {human_size(done)}",
                                    end="",
                                    flush=True,
                                )
                            last_print = now
                print()
                os.replace(part, dest)
                return dest

        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
            last_error = e
            print(f"\n[{label}失败] 第 {attempt}/{retries} 次: {e}")
            if attempt < retries:
                time.sleep(min(2 ** attempt, 8))

    raise RuntimeError(f"{label}失败: {url}\n最后错误: {last_error}")


# ---------------------------------------------------------------------------
# catalog.bundle 解包
# ---------------------------------------------------------------------------

def _read_cstring(data: bytes, off: int) -> tuple[str, int]:
    end = data.find(b"\x00", off)
    if end < 0:
        raise CatalogFormatError("UnityFS 字符串没有终止符")
    return data[off:end].decode("utf-8", errors="replace"), end + 1


def _align(off: int, alignment: int) -> int:
    return (off + alignment - 1) // alignment * alignment


def _lz4_block_decompress(src: bytes, expected_size: int | None = None) -> bytes:
    out = bytearray()
    i = 0
    n = len(src)

    while i < n:
        token = src[i]
        i += 1

        literal_len = token >> 4
        if literal_len == 15:
            while True:
                if i >= n:
                    raise CatalogFormatError("LZ4 literal 长度损坏")
                b = src[i]
                i += 1
                literal_len += b
                if b != 255:
                    break

        if i + literal_len > n:
            raise CatalogFormatError("LZ4 literal 越界")
        out.extend(src[i:i + literal_len])
        i += literal_len

        if i >= n:
            break

        if i + 2 > n:
            raise CatalogFormatError("LZ4 match offset 损坏")
        match_offset = src[i] | (src[i + 1] << 8)
        i += 2
        if match_offset <= 0 or match_offset > len(out):
            raise CatalogFormatError("LZ4 match offset 非法")

        match_len = token & 0x0F
        if match_len == 15:
            while True:
                if i >= n:
                    raise CatalogFormatError("LZ4 match 长度损坏")
                b = src[i]
                i += 1
                match_len += b
                if b != 255:
                    break
        match_len += 4

        start = len(out) - match_offset
        for _ in range(match_len):
            out.append(out[start])
            start += 1

    if expected_size is not None and len(out) != expected_size:
        raise CatalogFormatError(
            f"LZ4 解压长度不符: expected={expected_size}, actual={len(out)}"
        )
    return bytes(out)


def _decompress_unity_block(src: bytes, compression: int, expected_size: int) -> bytes:
    if compression == 0:
        out = src
    elif compression in (2, 3):  # LZ4 / LZ4HC
        out = _lz4_block_decompress(src, expected_size)
    elif compression == 1:  # LZMA
        # Unity 不同版本的 LZMA 包装略有差异；先尝试 Python 原生格式。
        errors = []
        for fmt in (lzma.FORMAT_AUTO, lzma.FORMAT_ALONE):
            try:
                out = lzma.decompress(src, format=fmt)
                break
            except Exception as e:
                errors.append(str(e))
        else:
            raise CatalogFormatError(
                "UnityFS 使用 LZMA，Python 原生解压失败。"
                "可将 catalog.bundle 用 UnityPy/AssetRipper 解出 TextAsset 后再运行本脚本。"
                f" 详细错误: {' | '.join(errors)}"
            )
    else:
        raise CatalogFormatError(f"不支持的 UnityFS 压缩类型: {compression}")

    if len(out) != expected_size:
        raise CatalogFormatError(
            f"UnityFS 数据块长度不符: expected={expected_size}, actual={len(out)}"
        )
    return out


def unpack_unityfs_nodes(data: bytes) -> list[tuple[str, bytes]]:
    if not data.startswith(b"UnityFS\x00"):
        raise CatalogFormatError("不是 UnityFS 文件")

    off = len(b"UnityFS\x00")
    if off + 4 > len(data):
        raise CatalogFormatError("UnityFS 头损坏")

    version = struct.unpack_from(">I", data, off)[0]
    off += 4
    unity_version, off = _read_cstring(data, off)
    unity_revision, off = _read_cstring(data, off)

    if version < 6:
        raise CatalogFormatError(f"暂不支持 UnityFS version {version}")

    if off + 20 > len(data):
        raise CatalogFormatError("UnityFS 头长度不足")

    file_size = struct.unpack_from(">Q", data, off)[0]
    off += 8
    compressed_info_size = struct.unpack_from(">I", data, off)[0]
    off += 4
    uncompressed_info_size = struct.unpack_from(">I", data, off)[0]
    off += 4
    flags = struct.unpack_from(">I", data, off)[0]
    off += 4

    if version >= 7:
        off = _align(off, 16)

    blocks_info_at_end = bool(flags & 0x80)
    need_padding_at_start = bool(flags & 0x200)
    info_compression = flags & 0x3F

    if blocks_info_at_end:
        info_pos = len(data) - compressed_info_size
        data_pos = off
    else:
        info_pos = off
        data_pos = off + compressed_info_size
        if need_padding_at_start:
            data_pos = _align(data_pos, 16)

    if info_pos < 0 or info_pos + compressed_info_size > len(data):
        raise CatalogFormatError("UnityFS blocks info 越界")

    compressed_info = data[info_pos:info_pos + compressed_info_size]
    info = _decompress_unity_block(
        compressed_info, info_compression, uncompressed_info_size
    )

    p = 0
    if len(info) < 20:
        raise CatalogFormatError("UnityFS blocks info 太短")
    p += 16  # uncompressed data hash

    block_count = struct.unpack_from(">I", info, p)[0]
    p += 4
    blocks = []
    for _ in range(block_count):
        if p + 10 > len(info):
            raise CatalogFormatError("UnityFS block 描述损坏")
        u_size = struct.unpack_from(">I", info, p)[0]
        c_size = struct.unpack_from(">I", info, p + 4)[0]
        b_flags = struct.unpack_from(">H", info, p + 8)[0]
        p += 10
        blocks.append((u_size, c_size, b_flags & 0x3F))

    if p + 4 > len(info):
        raise CatalogFormatError("UnityFS node count 缺失")
    node_count = struct.unpack_from(">I", info, p)[0]
    p += 4

    nodes_meta = []
    for _ in range(node_count):
        if p + 20 > len(info):
            raise CatalogFormatError("UnityFS node 描述损坏")
        node_off = struct.unpack_from(">Q", info, p)[0]
        node_size = struct.unpack_from(">Q", info, p + 8)[0]
        node_flags = struct.unpack_from(">I", info, p + 16)[0]
        p += 20
        name, p = _read_cstring(info, p)
        nodes_meta.append((name, node_off, node_size, node_flags))

    stream = bytearray()
    cur = data_pos
    for u_size, c_size, compression in blocks:
        if cur + c_size > len(data):
            raise CatalogFormatError("UnityFS 数据块越界")
        block_src = data[cur:cur + c_size]
        cur += c_size
        stream.extend(_decompress_unity_block(block_src, compression, u_size))

    nodes = []
    stream_bytes = bytes(stream)
    for name, node_off, node_size, _node_flags in nodes_meta:
        end = node_off + node_size
        if end > len(stream_bytes):
            continue
        nodes.append((name, stream_bytes[node_off:end]))

    if not nodes:
        # 极端情况下至少让后续 JSON 扫描整个解压流。
        nodes.append(("unityfs_stream", stream_bytes))

    return nodes


def _extract_balanced_json(data: bytes, start: int) -> bytes | None:
    depth = 0
    in_string = False
    escaped = False
    for i in range(start, len(data)):
        c = data[i]
        if in_string:
            if escaped:
                escaped = False
            elif c == 0x5C:  # \
                escaped = True
            elif c == 0x22:  # "
                in_string = False
        else:
            if c == 0x22:
                in_string = True
            elif c == 0x7B:  # {
                depth += 1
            elif c == 0x7D:  # }
                depth -= 1
                if depth == 0:
                    return data[start:i + 1]
    return None


def find_catalog_json_blob(data: bytes) -> bytes | None:
    # 先尝试整个文件就是 JSON。
    stripped = data.lstrip(b"\xef\xbb\xbf \t\r\n\x00")
    if stripped.startswith(b"{"):
        try:
            obj = json.loads(stripped.decode("utf-8"))
            if isinstance(obj, dict) and "m_InternalIds" in obj:
                return stripped
        except Exception:
            pass

    # TextAsset / 序列化数据中通常保留原始 JSON 字节。
    patterns = [
        re.compile(rb'\{\s*"m_LocatorId"\s*:'),
        re.compile(rb'\{\s*"m_BuildResultHash"\s*:'),
        re.compile(rb'\{\s*"m_InternalIdPrefixes"\s*:'),
    ]
    for pat in patterns:
        for m in pat.finditer(data):
            blob = _extract_balanced_json(data, m.start())
            if not blob:
                continue
            try:
                obj = json.loads(blob.decode("utf-8"))
                if isinstance(obj, dict) and "m_InternalIds" in obj:
                    return blob
            except Exception:
                continue
    return None


def extract_catalog_from_bundle(bundle_path: Path, output_path: Path) -> Path:
    data = bundle_path.read_bytes()

    # 1) 直接明文 JSON
    blob = find_catalog_json_blob(data)
    if blob:
        output_path.write_bytes(blob)
        return output_path

    # 2) gzip
    if data.startswith(b"\x1f\x8b"):
        try:
            unpacked = gzip.decompress(data)
            blob = find_catalog_json_blob(unpacked)
            if blob:
                output_path.write_bytes(blob)
                return output_path
        except Exception:
            pass

    # 3) zip
    if data.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                for name in zf.namelist():
                    raw = zf.read(name)
                    blob = find_catalog_json_blob(raw)
                    if blob:
                        output_path.write_bytes(blob)
                        return output_path
        except Exception:
            pass

    # 4) UnityFS AssetBundle
    if data.startswith(b"UnityFS\x00"):
        nodes = unpack_unityfs_nodes(data)
        for name, raw in nodes:
            blob = find_catalog_json_blob(raw)
            if blob:
                output_path.write_bytes(blob)
                return output_path

        # 某些 serialized file 里 JSON 可能跨我们关注的逻辑对象，
        # 再合并扫一次节点数据。
        combined = b"\n".join(raw for _, raw in nodes)
        blob = find_catalog_json_blob(combined)
        if blob:
            output_path.write_bytes(blob)
            return output_path

    raise CatalogFormatError(
        "未能从 catalog.bundle 中找到可用的 JSON catalog。\n"
        "该文件可能不是 JSON catalog，或使用了当前脚本未覆盖的 Unity/LZMA 变体。"
    )


def download_and_extract_catalog(
    work_dir: Path,
    *,
    bundle_url: str,
    timeout: int,
    retries: int,
) -> Path:
    bundle_path = work_dir / "catalog.bundle"
    catalog_path = work_dir / "catalog.json"

    print(f"\n下载 catalog.bundle:\n  {bundle_url}")
    download_single_with_progress(
        bundle_url,
        bundle_path,
        timeout=timeout,
        retries=retries,
        force=True,
        label="catalog",
    )

    print("正在从 catalog.bundle 提取 catalog...")
    tmp = catalog_path.with_suffix(".json.tmp")
    tmp.unlink(missing_ok=True)
    extract_catalog_from_bundle(bundle_path, tmp)

    # 验证确实可作为 catalog JSON 解析。
    with tmp.open("r", encoding="utf-8-sig") as f:
        obj = json.load(f)
    if not isinstance(obj, dict) or "m_InternalIds" not in obj:
        tmp.unlink(missing_ok=True)
        raise CatalogFormatError("提取出的 JSON 不是有效 Addressables catalog")

    os.replace(tmp, catalog_path)
    print(f"catalog 已提取到: {catalog_path}")
    return catalog_path


# ---------------------------------------------------------------------------
# catalog 分析
# ---------------------------------------------------------------------------

def _find_json_value_span(mm: mmap.mmap, key: str) -> Tuple[int, int]:
    needle = (f'"{key}"').encode("utf-8")
    p = mm.find(needle)
    if p < 0:
        raise CatalogFormatError(f"catalog 缺少字段: {key}")

    p += len(needle)
    n = len(mm)
    while p < n and mm[p] in b" \t\r\n":
        p += 1
    if p >= n or mm[p] != ord(":"):
        raise CatalogFormatError(f"catalog 字段格式错误: {key}")
    p += 1
    while p < n and mm[p] in b" \t\r\n":
        p += 1
    if p >= n:
        raise CatalogFormatError(f"catalog 字段没有值: {key}")

    start = p
    first = mm[p]

    if first == ord('"'):
        p += 1
        escaped = False
        while p < n:
            c = mm[p]
            if escaped:
                escaped = False
            elif c == ord("\\"):
                escaped = True
            elif c == ord('"'):
                return start, p + 1
            p += 1
        raise CatalogFormatError(f"catalog 字符串未结束: {key}")

    if first in (ord("["), ord("{")):
        opener = first
        closer = ord("]") if opener == ord("[") else ord("}")
        depth = 0
        in_string = False
        escaped = False
        while p < n:
            c = mm[p]
            if in_string:
                if escaped:
                    escaped = False
                elif c == ord("\\"):
                    escaped = True
                elif c == ord('"'):
                    in_string = False
            else:
                if c == ord('"'):
                    in_string = True
                elif c == opener:
                    depth += 1
                elif c == closer:
                    depth -= 1
                    if depth == 0:
                        return start, p + 1
            p += 1
        raise CatalogFormatError(f"catalog JSON 值未结束: {key}")

    while p < n and mm[p] not in b",}\r\n":
        p += 1
    return start, p


def load_required_catalog_fields(path: Path) -> Tuple[List[str], List[str], str, str]:
    with path.open("rb") as f:
        with mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) as mm:
            values = {}
            for key in (
                "m_InternalIdPrefixes",
                "m_InternalIds",
                "m_BucketDataString",
                "m_EntryDataString",
            ):
                a, b = _find_json_value_span(mm, key)
                values[key] = json.loads(mm[a:b].decode("utf-8"))

    prefixes = values["m_InternalIdPrefixes"]
    internal_ids = values["m_InternalIds"]
    bucket_b64 = values["m_BucketDataString"]
    entry_b64 = values["m_EntryDataString"]

    if not isinstance(prefixes, list) or not isinstance(internal_ids, list):
        raise CatalogFormatError("m_InternalIdPrefixes / m_InternalIds 类型错误")
    if not isinstance(bucket_b64, str) or not isinstance(entry_b64, str):
        raise CatalogFormatError("m_BucketDataString / m_EntryDataString 类型错误")

    return prefixes, internal_ids, bucket_b64, entry_b64


def resolve_internal_id(raw: str, prefixes: Sequence[str]) -> str:
    m = re.match(r"^(\d+)#(.*)$", raw, flags=re.S)
    if not m:
        return raw.replace("\\", "/")
    prefix_index = int(m.group(1))
    suffix = m.group(2)
    if not (0 <= prefix_index < len(prefixes)):
        return raw.replace("\\", "/")
    return (prefixes[prefix_index] + suffix).replace("\\", "/")


def _bundle_filename(raw_or_resolved: str) -> str:
    s = raw_or_resolved.replace("\\", "/")
    if "#/" in s:
        s = s.split("#/", 1)[1]
    return s.rsplit("/", 1)[-1]


def _decode_entries(entry_b64: str) -> Tuple[bytes, int]:
    try:
        data = base64.b64decode(entry_b64, validate=True)
    except Exception as e:
        raise CatalogFormatError(f"m_EntryDataString Base64 无效: {e}") from e
    if len(data) < 4:
        raise CatalogFormatError("m_EntryDataString 太短")
    count = struct.unpack_from("<I", data, 0)[0]
    expected = 4 + count * ENTRY_SIZE
    if expected != len(data):
        raise CatalogFormatError(
            f"EntryData 大小异常: count={count}, expected={expected}, actual={len(data)}"
        )
    return data, count


def _decode_buckets(bucket_b64: str) -> Tuple[bytes, int]:
    try:
        data = base64.b64decode(bucket_b64, validate=True)
    except Exception as e:
        raise CatalogFormatError(f"m_BucketDataString Base64 无效: {e}") from e
    if len(data) < 4:
        raise CatalogFormatError("m_BucketDataString 太短")
    return data, struct.unpack_from("<I", data, 0)[0]


def infer_semantics(asset_paths: Sequence[str]) -> Tuple[str, str, str, str, str]:
    if not asset_paths:
        return "", "", "", "", "no-assets"

    paths = [p.replace("\\", "/") for p in asset_paths]
    dirs = [posixpath.dirname(p) for p in paths]

    try:
        common_dir = posixpath.commonpath(dirs)
    except ValueError:
        common_dir = ""

    base = None
    if common_dir == "Assets/AssetBundles" or common_dir.startswith("Assets/AssetBundles/"):
        base = "Assets/AssetBundles"
    elif common_dir == "Assets" or common_dir.startswith("Assets/"):
        base = "Assets"

    category = ""
    version = ""
    asset_id = ""
    confidence = "mixed"

    if base:
        rel = common_dir[len(base):].strip("/")
        parts = [x for x in rel.split("/") if x]
        if parts:
            category = parts[0]
        if len(parts) >= 2 and re.search(r"\d", parts[-1]):
            asset_id = parts[-1]
            version = " ".join(parts[1:-1])
            confidence = "folder-id"
        elif len(parts) >= 2:
            version = " ".join(parts[1:])
            confidence = "category-only"

    if not asset_id:
        token_re = re.compile(r"(?i)([a-z][a-z0-9]*[_-]\d{3,})")
        tokens = set()
        for p in paths:
            m = token_re.search(posixpath.basename(p))
            if m:
                tokens.add(m.group(1))
        if len(tokens) == 1:
            asset_id = next(iter(tokens))
            confidence = "filename-id"

    return asset_id, version, category, common_dir, confidence


def analyze_catalog(path: Path) -> List[dict]:
    prefixes, internal_ids, bucket_b64, entry_b64 = load_required_catalog_fields(path)
    entry_data, entry_count = _decode_entries(entry_b64)
    bucket_data, bucket_count = _decode_buckets(bucket_b64)

    bundle_loc_to_internal: Dict[int, int] = {}
    for loc_index in range(entry_count):
        internal_idx = ENTRY_STRUCT.unpack_from(
            entry_data, 4 + loc_index * ENTRY_SIZE
        )[0]
        if 0 <= internal_idx < len(internal_ids):
            raw = internal_ids[internal_idx]
            if isinstance(raw, str) and raw.lower().endswith(".bundle"):
                bundle_loc_to_internal[loc_index] = internal_idx

    if not bundle_loc_to_internal:
        raise CatalogFormatError("catalog 中没有找到 .bundle location")

    bundle_location_set = set(bundle_loc_to_internal)
    owner_by_dependency_key = array("i", [-1]) * bucket_count
    dependency_bundle_count = array("I", [0]) * bucket_count
    dependency_bundle_locs_by_key: Dict[int, Tuple[int, ...]] = {}

    off = 4
    for bucket_index in range(bucket_count):
        if off + 8 > len(bucket_data):
            raise CatalogFormatError(f"BucketData 在 bucket {bucket_index} 截断")

        _key_data_offset, location_count = struct.unpack_from("<II", bucket_data, off)
        off += 8
        byte_count = location_count * 4

        if off + byte_count > len(bucket_data):
            raise CatalogFormatError(f"Bucket location 在 bucket {bucket_index} 截断")

        owner = -1
        bundle_count = 0
        bundle_locs: list[int] = []
        for j in range(location_count):
            loc = struct.unpack_from("<I", bucket_data, off + j * 4)[0]
            if loc in bundle_location_set:
                bundle_count += 1
                bundle_locs.append(loc)
                # Addressables BundledAssetProvider：依赖列表中第一个 bundle 为 asset 所在 bundle。
                if owner < 0:
                    owner = loc

        owner_by_dependency_key[bucket_index] = owner
        dependency_bundle_count[bucket_index] = bundle_count
        if bundle_locs:
            dependency_bundle_locs_by_key[bucket_index] = tuple(bundle_locs)
        off += byte_count

    assets_by_bundle: Dict[int, Set[int]] = defaultdict(set)
    dependencies_by_bundle: Dict[int, Set[int]] = defaultdict(set)
    max_dep_bundles_by_bundle: Dict[int, int] = defaultdict(int)

    for loc_index in range(entry_count):
        (
            internal_idx,
            _provider_idx,
            dep_key,
            _dep_hash,
            _data_idx,
            _primary_key,
            _rtype,
        ) = ENTRY_STRUCT.unpack_from(entry_data, 4 + loc_index * ENTRY_SIZE)

        if loc_index in bundle_location_set:
            continue
        if not (0 <= internal_idx < len(internal_ids)):
            continue
        if not (0 <= dep_key < bucket_count):
            continue

        owner = owner_by_dependency_key[dep_key]
        if owner < 0:
            continue

        assets_by_bundle[owner].add(internal_idx)
        for dep_loc in dependency_bundle_locs_by_key.get(dep_key, ()):
            if dep_loc != owner:
                dependencies_by_bundle[owner].add(dep_loc)
        if dependency_bundle_count[dep_key] > max_dep_bundles_by_bundle[owner]:
            max_dep_bundles_by_bundle[owner] = dependency_bundle_count[dep_key]

    bundle_name_by_loc = {
        loc: _bundle_filename(internal_ids[internal_idx])
        for loc, internal_idx in bundle_loc_to_internal.items()
    }

    rows: List[dict] = []
    for bundle_loc in sorted(bundle_loc_to_internal):
        bundle_internal_idx = bundle_loc_to_internal[bundle_loc]
        bundle_raw = internal_ids[bundle_internal_idx]
        asset_indexes = sorted(assets_by_bundle.get(bundle_loc, ()))
        asset_paths = [
            resolve_internal_id(internal_ids[i], prefixes)
            for i in asset_indexes
        ]

        asset_id, version, category, asset_root, confidence = infer_semantics(
            asset_paths
        )

        rows.append({
            "bundle": _bundle_filename(bundle_raw),
            "id": asset_id,
            "version": version,
            "category": category,
            "asset_root": asset_root,
            "asset_count": len(asset_paths),
            "confidence": confidence,
            "max_dependency_bundle_count": max_dep_bundles_by_bundle.get(bundle_loc, 0),
            "dependency_bundles": ";".join(
                bundle_name_by_loc[x]
                for x in sorted(dependencies_by_bundle.get(bundle_loc, ()))
                if x in bundle_name_by_loc
            ),
            "bundle_internal_id": resolve_internal_id(bundle_raw, prefixes),
            "bundle_location_index": bundle_loc,
        })

    return rows


# ---------------------------------------------------------------------------
# CSV
# ---------------------------------------------------------------------------

def read_mapping_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        required = {"bundle", "id", "version", "category"}
        fields = set(reader.fieldnames or [])
        missing = required - fields
        if missing:
            raise ValueError(f"CSV 缺少字段: {', '.join(sorted(missing))}")

        rows = []
        for r in reader:
            item = dict(r)
            item.setdefault("asset_root", "")
            item.setdefault("asset_count", "")
            item.setdefault("confidence", "")
            item.setdefault("max_dependency_bundle_count", "")
            item.setdefault("dependency_bundles", "")
            item.setdefault("bundle_internal_id", "")
            item.setdefault("bundle_location_index", "")
            rows.append(item)
        return rows


def write_mapping_csv(path: Path, rows: Sequence[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)

    if path.exists():
        stamp = time.strftime("%Y%m%d_%H%M%S")
        backup = path.with_name(f"{path.stem}.backup_{stamp}{path.suffix}")
        try:
            shutil.copy2(path, backup)
            print(f"旧 CSV 已备份: {backup.name}")
        except Exception:
            pass

    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in CSV_FIELDS})

    os.replace(tmp, path)
    print(f"对照 CSV 已生成: {path}")
    print("下次启动可直接用 CSV 快速筛选，不必重新解析 catalog，筛选速度会明显更快。")


# ---------------------------------------------------------------------------
# 查询筛选
# ---------------------------------------------------------------------------

ID_LIKE_RE = re.compile(r"(?i)^[a-z][a-z0-9]*[_-]\d+[a-z0-9_-]*$")


def _row_search_text(row: dict) -> str:
    return " | ".join(
        str(row.get(k, ""))
        for k in ("id", "category", "version", "asset_root", "bundle_internal_id", "bundle")
    ).casefold()


def filter_rows(rows: Sequence[dict], queries: Sequence[str]) -> list[dict]:
    selected: list[dict] = []
    seen = set()

    for raw_q in queries:
        q = raw_q.strip()
        if not q:
            continue

        mode = "auto"
        value = q

        if q.lower().startswith("id:"):
            mode = "id"
            value = q[3:].strip()
        elif q.lower().startswith(("key:", "path:", "type:")):
            mode = "key"
            value = q.split(":", 1)[1].strip()
        elif ID_LIKE_RE.fullmatch(q):
            mode = "id"
        else:
            mode = "key"

        needle = value.casefold()

        for row in rows:
            if mode == "id":
                matched = str(row.get("id", "")).casefold() == needle
            else:
                matched = needle in _row_search_text(row)

            if not matched:
                continue

            bundle = str(row.get("bundle", "")).strip()
            if not bundle:
                continue

            # 以 bundle hash 为唯一标识，避免多个关键词重复选中同一 bundle。
            key = bundle.casefold()
            if key not in seen:
                seen.add(key)
                selected.append(dict(row))

    selected.sort(
        key=lambda r: (
            str(r.get("category", "")).casefold(),
            str(r.get("id", "")).casefold(),
            str(r.get("version", "")).casefold(),
            str(r.get("bundle", "")).casefold(),
        )
    )
    return selected


def _parse_dependency_bundles(row: dict) -> list[str]:
    raw = str(row.get("dependency_bundles", "") or "")
    return [x.strip().replace("\\", "/").rsplit("/", 1)[-1] for x in raw.split(";") if x.strip()]


def expand_with_bundle_dependencies(
    direct_rows: Sequence[dict],
    all_rows: Sequence[dict],
) -> tuple[list[dict], list[str]]:
    """递归加入 catalog 已知的依赖 bundle，并返回无法解析的依赖名。"""
    row_by_bundle = {
        str(r.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1].casefold(): dict(r)
        for r in all_rows
        if str(r.get("bundle", "")).strip()
    }

    selected: dict[str, dict] = {}
    queue: list[tuple[dict, str]] = []
    for r in direct_rows:
        item = dict(r)
        item["_selected_reason"] = "query"
        b = str(item.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
        if b:
            selected[b.casefold()] = item
            queue.append((item, "query"))

    unresolved: list[str] = []
    unresolved_seen = set()
    qpos = 0
    while qpos < len(queue):
        row, _reason = queue[qpos]
        qpos += 1
        for dep in _parse_dependency_bundles(row):
            key = dep.casefold()
            if key in selected:
                continue
            dep_row = row_by_bundle.get(key)
            if dep_row is None:
                if key not in unresolved_seen:
                    unresolved_seen.add(key)
                    unresolved.append(dep)
                continue
            dep_row["_selected_reason"] = "dependency"
            selected[key] = dep_row
            queue.append((dep_row, "dependency"))

    result = list(selected.values())
    result.sort(
        key=lambda r: (
            0 if r.get("_selected_reason") == "query" else 1,
            str(r.get("category", "")).casefold(),
            str(r.get("id", "")).casefold(),
            str(r.get("version", "")).casefold(),
            str(r.get("bundle", "")).casefold(),
        )
    )
    return result, unresolved


def bundle_dependency_metadata_status(rows: Sequence[dict]) -> tuple[bool, int]:
    """返回 (metadata_complete_enough, rows_that_claim_multiple_bundles_without_names)."""
    uncertain = 0
    for r in rows:
        try:
            max_count = int(str(r.get("max_dependency_bundle_count", "") or "0"))
        except ValueError:
            max_count = 0
        if max_count > 1 and not _parse_dependency_bundles(r):
            uncertain += 1
    return uncertain == 0, uncertain


def build_local_names(rows: Sequence[dict]) -> list[str]:
    base_names = []
    for r in rows:
        rid = safe_piece(str(r.get("id", "")), "unknown-id")
        cat = safe_piece(str(r.get("category", "")), "unknown-category")
        ver = safe_piece(str(r.get("version", "")), "unknown-version")
        base_names.append(f"{rid}_{cat}_{ver}.bundle")

    counts: dict[str, int] = {}
    result = []
    for name in base_names:
        k = name.casefold()
        counts[k] = counts.get(k, 0) + 1
        n = counts[k]
        if n == 1:
            result.append(name)
        else:
            stem = name[:-7] if name.lower().endswith(".bundle") else name
            result.append(f"{stem}__{n:02d}.bundle")
    return result


# ---------------------------------------------------------------------------
# 多线程下载 + 实时单行状态
# ---------------------------------------------------------------------------

class ProgressTracker:
    """下载阶段只刷新一行，避免多线程时逐文件刷屏。"""
    def __init__(self, total_jobs: int):
        self.total_jobs = total_jobs
        self.completed = 0
        self.failed = 0
        self.skipped = 0
        self.states: dict[str, tuple[int, int | None, str]] = {}
        self.lock = threading.Lock()
        self.start = time.monotonic()
        self.last_render = 0.0
        self._last_width = 0

    def set_file(self, key: str, done: int, total: int | None, force: bool = False) -> None:
        with self.lock:
            old = self.states.get(key)
            status = old[2] if old else "active"
            self.states[key] = (done, total, status)
            self._render_locked(force)

    def mark_done(self, key: str, status: str) -> None:
        with self.lock:
            done, total, _old = self.states.get(key, (0, None, "active"))
            self.states[key] = (done, total, status)
            if status == "exists":
                self.skipped += 1
            elif status == "failed":
                self.failed += 1
            else:
                self.completed += 1
            self._render_locked(True)

    def _render_locked(self, force: bool = False) -> None:
        now = time.monotonic()
        if not force and now - self.last_render < 0.12:
            return
        self.last_render = now

        done_bytes = sum(v[0] for v in self.states.values())
        active = sum(1 for v in self.states.values() if v[2] == "active")
        finished = self.completed + self.failed + self.skipped
        elapsed = max(now - self.start, 0.001)
        speed = done_bytes / elapsed

        # 仅对已经获得 Content-Length 的任务统计字节进度；未开始文件不会造成假百分比。
        known_done = 0
        known_total = 0
        for done, total, _status in self.states.values():
            if total is not None and total > 0:
                known_done += min(done, total)
                known_total += total

        if known_total:
            pct = min(100.0, known_done * 100.0 / known_total)
            bytes_text = f"{human_size(known_done)}/{human_size(known_total)} {pct:6.2f}%"
        else:
            bytes_text = human_size(done_bytes)

        text = (
            f"下载 {finished:,}/{self.total_jobs:,} | 活动 {active} | {bytes_text}"
            f" | {human_size(int(speed))}/s"
            f" | 成功 {self.completed:,} 跳过 {self.skipped:,} 失败 {self.failed:,}"
        )
        width = max(self._last_width, len(text) + 2)
        self._last_width = width
        print("\r" + text.ljust(width), end="", flush=True)

    def finish_line(self) -> None:
        with self.lock:
            self._render_locked(True)
            print()


def download_bundle_job(
    row: dict,
    local_name: str,
    output_dir: Path,
    base_url: str,
    tracker: ProgressTracker,
    *,
    timeout: int,
    retries: int,
    force: bool,
) -> dict:
    bundle = str(row.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
    url = urllib.parse.urljoin(
        base_url.rstrip("/") + "/",
        urllib.parse.quote(bundle, safe=""),
    )
    dest = output_dir / local_name
    key = local_name

    record = {
        "id": row.get("id", ""),
        "category": row.get("category", ""),
        "version": row.get("version", ""),
        "selected_reason": row.get("_selected_reason", "query"),
        "original_bundle": bundle,
        "url": url,
        "local_file": local_name,
        "status": "",
    }

    if dest.exists() and not force:
        size = dest.stat().st_size
        tracker.set_file(key, size, size, True)
        tracker.mark_done(key, "exists")
        record["status"] = "exists"
        return record

    if force:
        dest.unlink(missing_ok=True)

    part = dest.with_name(dest.name + ".part")
    last_error: Exception | None = None

    for attempt in range(1, retries + 1):
        existing = part.stat().st_size if part.exists() else 0
        headers = {"User-Agent": USER_AGENT, "Accept-Encoding": "identity"}
        if existing:
            headers["Range"] = f"bytes={existing}-"

        req = urllib.request.Request(url, headers=headers)

        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                status = getattr(resp, "status", resp.getcode())
                clen = resp.headers.get("Content-Length")
                remain = int(clen) if clen and clen.isdigit() else None

                if existing and status == 206:
                    mode = "ab"
                    done = existing
                    total = existing + remain if remain is not None else None
                else:
                    mode = "wb"
                    done = 0
                    total = remain

                tracker.set_file(key, done, total, True)

                with part.open(mode) as f:
                    while True:
                        chunk = resp.read(1024 * 1024)
                        if not chunk:
                            break
                        f.write(chunk)
                        done += len(chunk)
                        tracker.set_file(key, done, total)

                os.replace(part, dest)
                final_size = dest.stat().st_size
                tracker.set_file(key, final_size, final_size, True)
                tracker.mark_done(key, "downloaded")
                record["status"] = "downloaded"
                return record

        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
            last_error = e
            if attempt < retries:
                time.sleep(min(2 ** attempt, 8))

    tracker.mark_done(key, "failed")
    record["status"] = f"failed: {last_error}"
    return record


def write_manifest(output_dir: Path, records: Sequence[dict]) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / "download_manifest.csv"
    fields = ["id", "category", "version", "selected_reason", "original_bundle", "url", "local_file", "status"]

    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(records)


# ---------------------------------------------------------------------------
# Export：隔离子进程，避免坏资源/native 解码拖垮主程序
# ---------------------------------------------------------------------------

EXPORT_PIP_REQUIREMENTS = ["UnityPy==1.25.3", "Pillow>=10"]
EXPORT_MANIFEST_FIELDS = ["bundle", "asset_path", "exported_path", "type", "path_id", "mode", "status"]


def export_dependency_status() -> tuple[bool, list[str]]:
    problems = []
    for module_name, display_name in (("UnityPy", "UnityPy"), ("PIL", "Pillow")):
        if importlib.util.find_spec(module_name) is None:
            problems.append(display_name + "(未安装)")
            continue
        try:
            importlib.import_module(module_name)
        except Exception as e:
            problems.append(f"{display_name}(导入失败: {type(e).__name__})")
    return not problems, problems


def print_export_dependency_status() -> None:
    ok, missing = export_dependency_status()
    if ok:
        try:
            from importlib.metadata import version
            uv = version("UnityPy")
            pv = version("Pillow")
            print(f"[Export依赖] UnityPy {uv} / Pillow {pv}：齐全")
        except Exception:
            print("[Export依赖] UnityPy / Pillow：齐全")
    else:
        print(f"[Export依赖] 缺少: {', '.join(missing)}。下载功能仍可使用；选择 Export 时可自动安装。")


def ensure_export_dependencies(allow_install: bool = True) -> bool:
    if sys.version_info < (3, 8):
        print("[Export错误] UnityPy 需要 Python 3.8 或更高版本。")
        return False

    ok, missing = export_dependency_status()
    if ok:
        return True

    print(f"Export 需要依赖: {', '.join(missing)}")
    print("建议安装: " + " ".join(EXPORT_PIP_REQUIREMENTS))
    if not allow_install:
        return False

    if not ask_yes_no("是否现在用当前 Python 自动安装缺少的 Export 依赖？", default=True):
        return False

    cmd = [sys.executable, "-m", "pip", "install", *EXPORT_PIP_REQUIREMENTS]
    try:
        subprocess.check_call(cmd)
    except Exception as e:
        print(f"依赖安装失败: {e}")
        return False

    importlib.invalidate_caches()
    ok, missing = export_dependency_status()
    if not ok:
        print(f"安装后仍缺少: {', '.join(missing)}")
        return False
    return True


def _safe_container_relative_path(container_path: str) -> PurePosixPath:
    raw = str(container_path).replace("\\", "/").strip()
    m = re.search(r"(?i)(?:^|/)(Assets/.*)$", raw)
    if m:
        raw = m.group(1)
    else:
        raw = raw.lstrip("/")
        if raw.lower().startswith("assets/"):
            raw = "Assets/" + raw.split("/", 1)[1]

    parts = []
    for part in PurePosixPath(raw).parts:
        if part in ("", "."):
            continue
        if part == "..":
            raise ValueError(f"不安全的 container 路径: {container_path}")
        part = re.sub(r'[<>:"|?*\x00-\x1f]', "_", part)
        parts.append(part)

    if not parts:
        raise ValueError(f"空 container 路径: {container_path}")
    return PurePosixPath(*parts)


def _image_output_relative_path(rel: PurePosixPath) -> PurePosixPath:
    stem, _ext = posixpath.splitext(rel.name)
    return rel.with_name(stem + ".png")


def _replace_last_extension(rel: PurePosixPath, extension: str) -> PurePosixPath:
    stem, _ext = posixpath.splitext(rel.name)
    if extension and not extension.startswith("."):
        extension = "." + extension
    return rel.with_name(stem + extension)


def _textasset_output_relative_path(rel: PurePosixPath) -> PurePosixPath:
    """
    TextAsset 导出文件名规则：
      xxx.atlas.txt  -> xxx.atlas
      xxx.skel.bytes -> xxx.skel
    其他 TextAsset 路径保持原样。
    """
    name_lower = rel.name.lower()
    if name_lower.endswith(".atlas.txt"):
        return rel.with_name(rel.name[:-4])   # 仅去掉末尾 .txt
    if name_lower.endswith(".skel.bytes"):
        return rel.with_name(rel.name[:-6])   # 仅去掉末尾 .bytes
    return rel


def _textasset_bytes(data) -> bytes:
    script = getattr(data, "m_Script", None)
    if script is None:
        script = getattr(data, "script", b"")
    if isinstance(script, str):
        return script.encode("utf-8", "surrogateescape")
    if isinstance(script, memoryview):
        return script.tobytes()
    if isinstance(script, bytearray):
        return bytes(script)
    if isinstance(script, bytes):
        return script
    try:
        return bytes(script)
    except Exception:
        return str(script).encode("utf-8", "surrogateescape")


def _export_id(asset_id: str, asset_root: str) -> str:
    value = (asset_id or "").strip()
    if not value:
        parts = [x for x in str(asset_root).replace("\\", "/").split("/") if x]
        for part in reversed(parts):
            if ID_LIKE_RE.fullmatch(part):
                value = part
                break
    return safe_piece(value, "unknown-id")


def _export_category(category: str, asset_root: str) -> str:
    text = f"{category} {asset_root}".casefold()
    if "character" in text:
        return "character"
    if "still" in text:
        return "still"
    if "sister" in text:
        return "sister"
    value = (category or "").strip().casefold()
    return safe_piece(value, "other")


def _route_export_relative_path(
    source_rel: PurePosixPath,
    layout: str,
    asset_id: str,
    category: str,
    asset_root: str,
) -> PurePosixPath:
    if layout == "original":
        return source_rel
    rid = _export_id(asset_id, asset_root)
    if layout == "id-category":
        return PurePosixPath(rid, _export_category(category, asset_root), source_rel.name)
    if layout == "id-flat":
        return PurePosixPath(rid, source_rel.name)
    raise ValueError(f"未知 Export 布局: {layout}")


def _claim_flat_relative_path(
    out_rel: PurePosixPath,
    registry_path: Path | None,
    source_key: str,
    layout: str,
) -> PurePosixPath:
    """平铺/分类模式避免同名覆盖；原目录模式保持真实路径并允许覆盖。"""
    if layout == "original" or registry_path is None:
        return out_rel
    try:
        registry = json.loads(registry_path.read_text(encoding="utf-8")) if registry_path.exists() else {}
    except Exception:
        registry = {}

    candidate = out_rel
    index = 1
    while True:
        key = candidate.as_posix().casefold()
        owner = registry.get(key)
        if owner is None or owner == source_key:
            registry[key] = source_key
            _atomic_write_json(registry_path, registry)
            return candidate
        index += 1
        stem, ext = posixpath.splitext(out_rel.name)
        candidate = out_rel.with_name(f"{stem}__{index:02d}{ext}")


def _prepare_dest(
    source_rel: PurePosixPath,
    export_root: Path,
    *,
    layout: str,
    asset_id: str,
    category: str,
    asset_root: str,
    registry_path: Path | None,
    source_key: str,
) -> tuple[Path, PurePosixPath]:
    out_rel = _route_export_relative_path(source_rel, layout, asset_id, category, asset_root)
    out_rel = _claim_flat_relative_path(out_rel, registry_path, source_key, layout)
    dest = export_root.joinpath(*out_rel.parts)
    dest.parent.mkdir(parents=True, exist_ok=True)
    return dest, out_rel


def _resolve_object_reader(obj):
    """
    UnityPy env.container 的 value 在当前版本通常是 PPtr，而 env.objects 才是 ObjectReader。
    所有真正需要 reader API（parse_as_object/get_raw_data/type/path_id）的代码先走这里。
    """
    if obj is None:
        raise ValueError("对象为空")
    if hasattr(obj, "get_raw_data") and hasattr(obj, "parse_as_object"):
        return obj
    deref = getattr(obj, "deref", None)
    if callable(deref):
        return deref()
    return obj


def _object_reader_key(obj) -> tuple[int, int]:
    try:
        reader = _resolve_object_reader(obj)
        assets_file = getattr(reader, "assets_file", None)
        path_id = int(getattr(reader, "path_id", 0) or 0)
        return (id(assets_file), path_id)
    except Exception:
        assets_file = getattr(obj, "assets_file", getattr(obj, "assetsfile", None))
        path_id = getattr(obj, "path_id", getattr(obj, "m_PathID", 0))
        try:
            path_id = int(path_id or 0)
        except Exception:
            path_id = 0
        return (id(assets_file), path_id)


def _pptr_valid(ptr) -> bool:
    if ptr is None:
        return False
    path_id = getattr(ptr, "path_id", getattr(ptr, "m_PathID", 0))
    try:
        return int(path_id or 0) != 0
    except Exception:
        return bool(path_id)


def _sprite_render_data(sprite):
    """解析 SpriteAtlas 时使用 atlas 的 RenderData；否则返回 Sprite 自己的 m_RD。"""
    try:
        atlas_ptr = getattr(sprite, "m_SpriteAtlas", None)
        if _pptr_valid(atlas_ptr):
            atlas = atlas_ptr.deref_parse_as_object() if hasattr(atlas_ptr, "deref_parse_as_object") else atlas_ptr.deref().parse_as_object()
            key = getattr(sprite, "m_RenderDataKey", None)
            for pair in getattr(atlas, "m_RenderDataMap", []) or []:
                try:
                    k, value = pair
                    if k == key:
                        return value
                except Exception:
                    continue
    except Exception:
        pass

    # 某些 Sprite 只有 AtlasTags，没有直接 m_SpriteAtlas 指针；按 UnityPy SpriteHelper 的方式回查 SpriteAtlas。
    try:
        tags = getattr(sprite, "m_AtlasTags", None) or []
        assets_file = getattr(sprite, "assets_file", None)
        if tags and assets_file is not None:
            wanted = str(tags[0])
            for reader in getattr(assets_file, "objects", {}).values():
                if getattr(getattr(reader, "type", None), "name", "") != "SpriteAtlas":
                    continue
                try:
                    if reader.peek_name() != wanted:
                        continue
                    atlas = reader.parse_as_object()
                    key = getattr(sprite, "m_RenderDataKey", None)
                    for pair in getattr(atlas, "m_RenderDataMap", []) or []:
                        k, value = pair
                        if k == key:
                            return value
                except Exception:
                    continue
    except Exception:
        pass
    return getattr(sprite, "m_RD", None)


def _full_image_from_sprite(sprite):
    """
    导出 Sprite 底层完整 Texture2D，不调用 sprite.image，因此不做 textureRect crop。
    如果存在 split-alpha，按 UnityPy SpriteHelper 的规则把 alpha 图 R 通道合并到整张纹理。
    """
    rd = _sprite_render_data(sprite)
    if rd is None:
        raise ValueError("Sprite 没有 RenderData")
    texture_ptr = getattr(rd, "texture", None)
    if not _pptr_valid(texture_ptr):
        raise ValueError("Sprite 没有可用 Texture2D 指针")

    tex_reader = texture_ptr.deref() if hasattr(texture_ptr, "deref") else None
    if tex_reader is None:
        raise ValueError("无法解引用 Sprite Texture2D")
    texture = tex_reader.parse_as_object() if hasattr(tex_reader, "parse_as_object") else tex_reader.read()
    image = texture.image

    alpha_ptr = getattr(rd, "alphaTexture", None)
    if _pptr_valid(alpha_ptr):
        try:
            alpha_reader = alpha_ptr.deref()
            alpha_tex = alpha_reader.parse_as_object() if hasattr(alpha_reader, "parse_as_object") else alpha_reader.read()
            alpha_image = alpha_tex.image
            if alpha_image.size == image.size:
                rgb = image.convert("RGBA")
                alpha_channel = alpha_image.split()[0]
                rgb.putalpha(alpha_channel)
                image = rgb
        except Exception:
            # alpha 单独失败时仍保留完整主纹理，绝不退回 Sprite crop。
            pass
    return image


def _detect_payload_extension(data: bytes, *, audio: bool = False, video: bool = False) -> str:
    head = memoryview(data)[:16]
    if head[:4] == b"OggS":
        return ".ogg" if audio else ".ogv"
    if head[:4] == b"RIFF":
        if len(head) >= 12 and bytes(head[8:12]) == b"AVI ":
            return ".avi"
        return ".wav" if audio else ".riff"
    if len(head) >= 8 and bytes(head[4:8]) == b"ftyp":
        return ".m4a" if audio else ".mp4"
    if bytes(head[:4]) == b"\x1aE\xdf\xa3":
        return ".webm"
    if bytes(head[:3]) == b"ID3" or (len(head) >= 2 and head[0] == 0xFF and (head[1] & 0xE0) == 0xE0):
        return ".mp3"
    return ".audio" if audio else ".video" if video else ".bin"


def _resource_payload(data) -> bytes | None:
    """尝试读取对象内嵌数据或 Unity 外部 resource 数据。"""
    for attr in ("m_AudioData", "m_VideoData", "m_Data"):
        value = getattr(data, attr, None)
        if value:
            try:
                raw = bytes(value)
                if raw:
                    return raw
            except Exception:
                pass

    resources = []
    for attr in ("m_Resource", "m_ExternalResource", "m_ExternalResources"):
        value = getattr(data, attr, None)
        if not value:
            continue
        if isinstance(value, (list, tuple)):
            resources.extend(value)
        else:
            resources.append(value)

    reader = getattr(data, "object_reader", None)
    assets_file = getattr(reader, "assets_file", None)
    if assets_file is None:
        return None

    try:
        from UnityPy.helpers.ResourceReader import get_resource_data
    except Exception:
        return None

    for resource in resources:
        source = getattr(resource, "m_Source", getattr(resource, "source", ""))
        offset = getattr(resource, "m_Offset", getattr(resource, "offset", 0))
        size = getattr(resource, "m_Size", getattr(resource, "size", 0))
        if not source or not size:
            continue
        try:
            raw = get_resource_data(source, assets_file, int(offset), int(size))
            if raw:
                return bytes(raw)
        except Exception:
            continue
    return None


def _record(bundle_name: str, asset_path: str, out_rel: PurePosixPath, obj_type: str, path_id, mode: str, status: str = "ok") -> dict:
    return {
        "bundle": bundle_name,
        "asset_path": asset_path,
        "exported_path": out_rel.as_posix(),
        "type": obj_type,
        "path_id": path_id,
        "mode": mode,
        "status": status,
    }


def _safe_peek_name(obj, fallback: str) -> str:
    try:
        name = obj.peek_name()
        if name:
            return safe_piece(str(name), fallback)
    except Exception:
        pass
    return safe_piece(fallback, "unnamed")


def _synthetic_asset_path(obj, asset_root: str) -> str:
    obj_type = getattr(getattr(obj, "type", None), "name", str(getattr(obj, "type", "Unknown")))
    path_id = int(getattr(obj, "path_id", 0) or 0)
    name = _safe_peek_name(obj, f"{obj_type}_{path_id}")
    default_ext = {
        "Texture2D": ".png", "Sprite": ".png", "TextAsset": ".txt",
        "AudioClip": ".audio", "VideoClip": ".video", "MovieTexture": ".video",
        "Font": ".font", "Mesh": ".mesh", "Shader": ".shader",
        "Texture2DArray": ".png", "Material": ".mat", "AnimationClip": ".anim",
        "AnimatorController": ".controller", "RuntimeAnimatorController": ".controller",
        "Avatar": ".avatar", "MonoBehaviour": ".asset", "SpriteAtlas": ".spriteatlas",
        "Cubemap": ".cubemap",
    }.get(obj_type, ".asset")
    if not posixpath.splitext(name)[1]:
        name += default_ext
    root = str(asset_root or "").replace("\\", "/").rstrip("/")
    if root:
        return root + "/" + name
    return f"Assets/__uncontainered__/{safe_piece(obj_type, 'Unknown')}/{name}"


UNCONTAINERED_CONTENT_TYPES = {
    "Texture2D", "Sprite", "TextAsset", "AudioClip", "VideoClip", "MovieTexture",
    "Font", "Mesh", "Shader", "Texture2DArray", "Material", "AnimationClip",
    "AnimatorController", "RuntimeAnimatorController", "Avatar", "MonoBehaviour",
    "SpriteAtlas", "Cubemap", "ComputeShader", "AudioMixer", "AudioMixerGroup",
    "AudioMixerSnapshot", "RenderTexture", "TerrainData", "AssetBundleManifest",
}


def export_one_container_object(
    obj,
    container_path: str,
    export_root: Path,
    bundle_name: str = "",
    *,
    export_layout: str = "original",
    asset_id: str = "",
    category: str = "",
    asset_root: str = "",
    registry_path: Path | None = None,
) -> list[dict]:
    rel = _safe_container_relative_path(container_path)
    # 注意：env.container 的 value 可能是 PPtr。不能直接调用 PPtr.get_raw_data()。
    path_id = getattr(obj, "path_id", getattr(obj, "m_PathID", ""))
    obj_type = type(obj).__name__ or "Unknown"
    source_key = f"{bundle_name}:{path_id}:{container_path}"

    try:
        obj = _resolve_object_reader(obj)
        obj_type = getattr(getattr(obj, "type", None), "name", str(getattr(obj, "type", "Unknown")))
        path_id = getattr(obj, "path_id", path_id)
        source_key = f"{bundle_name}:{path_id}:{container_path}"

        # Texture2D 本身就是完整贴图；Sprite 则只取它引用的完整 Texture2D，绝不使用 sprite.image crop。
        if obj_type == "Texture2D":
            data = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            source_rel = _image_output_relative_path(rel)
            dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                          category=category, asset_root=asset_root, registry_path=registry_path,
                                          source_key=source_key)
            data.image.save(dest, format="PNG")
            return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "png-full-texture")]

        if obj_type == "Sprite":
            sprite = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            source_rel = _image_output_relative_path(rel)
            dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                          category=category, asset_root=asset_root, registry_path=registry_path,
                                          source_key=source_key)
            _full_image_from_sprite(sprite).save(dest, format="PNG")
            return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "png-full-sprite-texture-no-crop")]

        if obj_type == "TextAsset":
            data = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            source_rel = _textasset_output_relative_path(rel)
            dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                          category=category, asset_root=asset_root, registry_path=registry_path,
                                          source_key=source_key)
            dest.write_bytes(_textasset_bytes(data))
            if rel.name.lower().endswith(".atlas.txt"):
                mode = "textasset-atlas"
            elif rel.name.lower().endswith(".skel.bytes"):
                mode = "textasset-skel"
            else:
                mode = "textasset-exact"
            return [_record(bundle_name, container_path, out_rel, obj_type, path_id, mode)]

        if obj_type == "AudioClip":
            clip = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            records = []
            try:
                samples = dict(clip.samples)
            except Exception:
                samples = {}

            if samples:
                base_stem, _ = posixpath.splitext(rel.name)
                multi = len(samples) > 1
                for n, (sample_name, sample_data) in enumerate(samples.items(), 1):
                    sample_name = str(sample_name).replace("\\", "/").rsplit("/", 1)[-1]
                    _sample_stem, sample_ext = posixpath.splitext(sample_name)
                    if not sample_ext:
                        sample_ext = _detect_payload_extension(bytes(sample_data), audio=True)
                    if multi:
                        suffix_name = safe_piece(_sample_stem, f"sample{n}")
                        source_rel = rel.with_name(f"{base_stem}__{suffix_name}{sample_ext}")
                    else:
                        source_rel = rel.with_name(f"{base_stem}{sample_ext}")
                    dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                                  category=category, asset_root=asset_root, registry_path=registry_path,
                                                  source_key=source_key + f":audio:{n}")
                    dest.write_bytes(bytes(sample_data))
                    records.append(_record(bundle_name, container_path, out_rel, obj_type, path_id, "audio-sample"))
                return records

            raw = _resource_payload(clip)
            if raw:
                ext = _detect_payload_extension(raw, audio=True)
                source_rel = _replace_last_extension(rel, ext)
                dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                              category=category, asset_root=asset_root, registry_path=registry_path,
                                              source_key=source_key + ":audio-raw")
                dest.write_bytes(raw)
                return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "audio-resource-raw")]
            raise ValueError("AudioClip 没有可导出的 samples 或 resource 数据")

        if obj_type in {"VideoClip", "MovieTexture"}:
            data = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            raw = _resource_payload(data)
            if not raw:
                raise ValueError("VideoClip 没有找到内嵌/外部视频数据")
            ext = _detect_payload_extension(raw, video=True)
            source_rel = _replace_last_extension(rel, ext)
            dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                          category=category, asset_root=asset_root, registry_path=registry_path,
                                          source_key=source_key + ":video")
            dest.write_bytes(raw)
            return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "video-resource")]

        if obj_type == "Font":
            data = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            raw = bytes(getattr(data, "m_FontData", b"") or b"")
            if raw:
                ext = ".otf" if raw[:4] == b"OTTO" else ".ttf"
                source_rel = _replace_last_extension(rel, ext)
                dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                              category=category, asset_root=asset_root, registry_path=registry_path,
                                              source_key=source_key)
                dest.write_bytes(raw)
                return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "font")]

        if obj_type == "Mesh":
            data = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            source_rel = _replace_last_extension(rel, ".obj")
            dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                          category=category, asset_root=asset_root, registry_path=registry_path,
                                          source_key=source_key)
            with dest.open("w", encoding="utf-8", newline="") as f:
                f.write(data.export())
            return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "mesh-obj")]

        if obj_type == "Shader":
            data = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            source_rel = _replace_last_extension(rel, ".shader")
            dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                          category=category, asset_root=asset_root, registry_path=registry_path,
                                          source_key=source_key)
            shader_text = data.export() if hasattr(data, "export") else str(data)
            with dest.open("w", encoding="utf-8", newline="") as f:
                f.write(shader_text)
            return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "shader-text")]

        if obj_type == "Texture2DArray":
            data = obj.parse_as_object() if hasattr(obj, "parse_as_object") else obj.read()
            images = list(getattr(data, "images", []) or [])
            if not images:
                raise ValueError("Texture2DArray 没有可导出的 images")
            records = []
            base_stem, _ = posixpath.splitext(rel.name)
            for idx, image in enumerate(images):
                source_rel = rel.with_name(f"{base_stem}_{idx}.png")
                dest, out_rel = _prepare_dest(source_rel, export_root, layout=export_layout, asset_id=asset_id,
                                              category=category, asset_root=asset_root, registry_path=registry_path,
                                              source_key=source_key + f":array:{idx}")
                image.save(dest, format="PNG")
                records.append(_record(bundle_name, container_path, out_rel, obj_type, path_id, "png-texture-array"))
            return records

        # 其余已有真实 container 路径的对象保留原路径/原后缀，写原始序列化对象数据。
        dest, out_rel = _prepare_dest(rel, export_root, layout=export_layout, asset_id=asset_id,
                                      category=category, asset_root=asset_root, registry_path=registry_path,
                                      source_key=source_key)
        raw = obj.get_raw_data()
        if isinstance(raw, memoryview):
            raw = raw.tobytes()
        elif not isinstance(raw, (bytes, bytearray)):
            raw = bytes(raw)
        dest.write_bytes(bytes(raw))
        return [_record(bundle_name, container_path, out_rel, obj_type, path_id, "raw-exact")]

    except Exception as e:
        fail_rel = _route_export_relative_path(rel, export_layout, asset_id, category, asset_root)
        return [_record(bundle_name, container_path, fail_rel, obj_type, path_id, "failed",
                        f"failed: {type(e).__name__}: {e}")]

def _write_status_json(path: Path, data: dict) -> None:
    """
    状态文件是临时 UI 数据，允许主进程偶尔读到半写入 JSON 后重试。
    Windows 上不要对它高频 os.replace()：读者短暂打开目标文件时，rename/replace
    需要 delete access，容易触发 WinError 5。
    """
    payload = json.dumps(data, ensure_ascii=False)
    path.parent.mkdir(parents=True, exist_ok=True)
    last_error = None
    for attempt in range(8):
        try:
            with path.open("w", encoding="utf-8", newline="") as f:
                f.write(payload)
                f.flush()
            return
        except (PermissionError, OSError) as e:
            last_error = e
            time.sleep(0.01 * (attempt + 1))
    raise last_error if last_error else OSError(f"无法写入状态文件: {path}")


def _read_json_tolerant(path: Path, default=None, retries: int = 3):
    for attempt in range(max(1, retries)):
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (FileNotFoundError, PermissionError, OSError, json.JSONDecodeError):
            if attempt + 1 < retries:
                time.sleep(0.008)
    return default


def _atomic_write_json(path: Path, data: dict) -> None:
    """
    持久/结果 JSON 仍使用原子替换，但 temp 文件名唯一并对 Windows 文件锁做重试。
    最终若 replace 连续失败，则退回直接写目标文件。
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, ensure_ascii=False)
    tmp = path.with_name(f".{path.name}.{os.getpid()}.{time.time_ns()}.tmp")
    tmp.write_text(payload, encoding="utf-8")
    last_error = None
    for attempt in range(20):
        try:
            os.replace(tmp, path)
            return
        except (PermissionError, OSError) as e:
            last_error = e
            time.sleep(min(0.01 * (attempt + 1), 0.20))
    try:
        path.write_text(payload, encoding="utf-8")
        tmp.unlink(missing_ok=True)
        return
    except Exception:
        tmp.unlink(missing_ok=True)
        if last_error:
            raise last_error
        raise OSError(f"无法原子写入 JSON: {path}")


def _unlink_retry(path: Path, attempts: int = 12) -> bool:
    for i in range(attempts):
        try:
            path.unlink(missing_ok=True)
            return True
        except (PermissionError, OSError):
            time.sleep(min(0.02 * (i + 1), 0.20))
    return not path.exists()


def _rmtree_retry(path: Path, attempts: int = 12) -> bool:
    if not path.exists():
        return True
    for i in range(attempts):
        try:
            shutil.rmtree(path)
            return True
        except (PermissionError, OSError):
            time.sleep(min(0.03 * (i + 1), 0.30))
    return not path.exists()


def _cleanup_known_temp_files(output_dir: Path | None = None, export_root: Path | None = None, work_dir: Path | None = None) -> list[str]:
    """
    只清理本程序明确创建、且完成后无用的临时文件。
    失败下载留下的非空 .part 会保留用于下次断点续传；若最终文件已存在则删除对应 .part。
    """
    removed = []
    if export_root:
        work = export_root / ".export_work"
        if work.exists() and _rmtree_retry(work):
            removed.append(str(work))
        # 历史版本遗留在 export 根目录的明确临时状态文件
        for pat in ("*.status.json.tmp", "*.result.json.tmp", "name_registry.json.*.tmp"):
            for p in export_root.glob(pat):
                if _unlink_retry(p):
                    removed.append(str(p))

    if output_dir and output_dir.exists():
        for part in output_dir.glob("*.part"):
            final = Path(str(part)[:-5])
            try:
                useless = final.exists() or part.stat().st_size == 0
            except OSError:
                useless = False
            if useless and _unlink_retry(part):
                removed.append(str(part))

    if work_dir:
        for p in (work_dir / "catalog.json.tmp", work_dir / "catalog.txt.tmp"):
            if p.exists() and _unlink_retry(p):
                removed.append(str(p))
        # 清理旧版原子 JSON 写入遗留的唯一 tmp
        for p in work_dir.glob(".*.json.*.tmp"):
            if _unlink_retry(p):
                removed.append(str(p))
    return removed


def _container_path_matches_root(asset_path: str, asset_root: str) -> bool:
    if not asset_root:
        return True
    try:
        p = _safe_container_relative_path(asset_path).as_posix().rstrip("/").casefold()
        r = _safe_container_relative_path(asset_root).as_posix().rstrip("/").casefold()
    except Exception:
        return False
    return p == r or p.startswith(r + "/")


def export_worker_main(job_path: Path) -> int:
    """单 bundle Export worker。native crash 时只会终止本子进程。"""
    job = json.loads(job_path.read_text(encoding="utf-8"))
    primary = Path(job["primary_path"])
    dependencies = [Path(x) for x in job.get("dependency_paths", []) if Path(x).is_file()]
    export_root = Path(job["export_root"])
    asset_root = str(job.get("asset_root", "") or "")
    asset_id = str(job.get("asset_id", "") or "")
    category = str(job.get("category", "") or "")
    export_layout = str(job.get("export_layout", "original") or "original")
    registry_path = Path(job["registry_path"]) if job.get("registry_path") else None
    bundle_name = str(job.get("bundle", primary.name))
    status_path = Path(job["status_path"])
    result_path = Path(job["result_path"])

    try:
        import UnityPy
        try:
            from UnityPy.helpers import TypeTreeHelper
            TypeTreeHelper.read_typetree_c = False
        except Exception:
            pass

        env = UnityPy.load(str(primary))
        try:
            env.path = str(primary.parent)
        except Exception:
            pass

        # 必须在加载依赖前保存主 bundle 的对象集合，避免把依赖对象误当成本体导出。
        primary_items = list(env.container.items())
        primary_objects = list(env.objects)

        for dep in dependencies:
            try:
                try:
                    env.load_file(str(dep), is_dependency=True)
                except TypeError:
                    env.load_file(str(dep), True)
            except Exception:
                pass

        # 主 bundle 的所有 container 条目都应导出。旧版按 asset_root 过滤会漏掉同 bundle 中
        # 路径位于其他目录的 AudioClip/共享资源；依赖 bundle 已在 primary_items 捕获之后才加载，因此这里不会混入依赖。
        container_items = primary_items

        # container 只覆盖“有路径”的对象。声音/字体/纹理等有时没有 container 条目，补扫主 bundle objects。
        contained_keys = {_object_reader_key(o) for _p, o in primary_items}
        extra_items = []
        for obj in primary_objects:
            if _object_reader_key(obj) in contained_keys:
                continue
            obj_type = getattr(getattr(obj, "type", None), "name", str(getattr(obj, "type", "Unknown")))
            if obj_type not in UNCONTAINERED_CONTENT_TYPES:
                continue
            extra_items.append((_synthetic_asset_path(obj, asset_root), obj))

        items = container_items + extra_items
        total = len(items)
        ok = failed = 0
        records = []
        _write_status_json(status_path, {"current": 0, "total": total, "ok": 0, "failed": 0, "path": ""})

        for i, (asset_path, obj) in enumerate(items, 1):
            _write_status_json(status_path, {"current": i - 1, "total": total, "ok": ok, "failed": failed, "path": str(asset_path)})
            recs = export_one_container_object(
                obj, asset_path, export_root, bundle_name,
                export_layout=export_layout,
                asset_id=asset_id,
                category=category,
                asset_root=asset_root,
                registry_path=registry_path,
            )
            records.extend(recs)
            if recs and all(r.get("status") == "ok" for r in recs):
                ok += 1
            else:
                failed += 1
            if i == total or i % 4 == 0:
                _write_status_json(status_path, {"current": i, "total": total, "ok": ok, "failed": failed, "path": str(asset_path)})

        _atomic_write_json(result_path, {"status": "ok", "ok": ok, "failed": failed, "records": records})
        _write_status_json(status_path, {"current": total, "total": total, "ok": ok, "failed": failed, "path": "", "done": True})
        return 0
    except BaseException as e:
        try:
            _atomic_write_json(result_path, {
                "status": "worker-error", "ok": 0, "failed": 1, "records": [],
                "error": f"{type(e).__name__}: {e}",
            })
        except Exception:
            pass
        return 70

def verify_downloaded_bundle_dependencies(rows: Sequence[dict], download_records: Sequence[dict]) -> tuple[bool, list[str], int]:
    available = {
        str(r.get("original_bundle", "")).casefold()
        for r in download_records
        if r.get("status") in {"downloaded", "exists"}
    }
    expected = set()
    uncertain = 0
    for row in rows:
        deps = _parse_dependency_bundles(row)
        expected.update(x.casefold() for x in deps)
        try:
            max_count = int(str(row.get("max_dependency_bundle_count", "") or "0"))
        except ValueError:
            max_count = 0
        if max_count > 1 and not deps:
            uncertain += 1

    missing = sorted(x for x in expected if x not in available)
    return not missing, missing, uncertain


def _collect_recursive_dependency_names(row: dict, all_rows: Sequence[dict]) -> list[str]:
    row_by_bundle = {
        str(r.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1].casefold(): r
        for r in all_rows if str(r.get("bundle", "")).strip()
    }
    out = []
    seen = set()
    queue = list(_parse_dependency_bundles(row))
    pos = 0
    while pos < len(queue):
        dep = queue[pos]
        pos += 1
        key = dep.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(dep)
        dep_row = row_by_bundle.get(key)
        if dep_row:
            queue.extend(_parse_dependency_bundles(dep_row))
    return out


def export_downloaded_bundles_isolated(
    direct_rows: Sequence[dict],
    all_rows: Sequence[dict],
    download_records: Sequence[dict],
    output_dir: Path,
    export_root: Path,
    export_layout: str = "original",
) -> tuple[int, int, Path, Path]:
    """逐主 bundle 启动独立 Python worker；任何 native crash 都不会杀死主进程。"""
    local_by_bundle = {
        str(r.get("original_bundle", "")).replace("\\", "/").rsplit("/", 1)[-1].casefold(): output_dir / str(r.get("local_file", ""))
        for r in download_records
        if r.get("status") in {"downloaded", "exists"}
    }

    jobs = []
    for row in direct_rows:
        bundle = str(row.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
        primary = local_by_bundle.get(bundle.casefold())
        if not primary or not primary.is_file():
            continue
        dep_paths = []
        seen = set()
        for dep in _collect_recursive_dependency_names(row, all_rows):
            p = local_by_bundle.get(dep.casefold())
            if p and p.is_file() and str(p).casefold() not in seen:
                seen.add(str(p).casefold())
                dep_paths.append(p)
        jobs.append((row, bundle, primary, dep_paths))

    if not jobs:
        raise RuntimeError("没有可供 Export 的主 bundle 文件")

    export_root.mkdir(parents=True, exist_ok=True)
    work = export_root / ".export_work"
    # 上次异常退出留下的 job/status/log 不参与新任务，启动前先清理。
    _rmtree_retry(work)
    work.mkdir(parents=True, exist_ok=True)
    manifest = export_root / "export_manifest.csv"
    crash_log = export_root / "export_failures.log"
    registry_path = work / "name_registry.json"
    _atomic_write_json(registry_path, {})
    crash_log.unlink(missing_ok=True)

    total_ok = total_failed = 0
    last_width = 0

    with manifest.open("w", encoding="utf-8-sig", newline="") as mf:
        writer = csv.DictWriter(mf, fieldnames=EXPORT_MANIFEST_FIELDS)
        writer.writeheader()

        for idx, (row, bundle, primary, dep_paths) in enumerate(jobs, 1):
            token = f"job_{idx:05d}"
            job_path = work / f"{token}.json"
            status_path = work / f"{token}.status.json"
            result_path = work / f"{token}.result.json"
            log_path = work / f"{token}.log"
            for p in (status_path, result_path, log_path):
                _unlink_retry(p)

            spec = {
                "primary_path": str(primary),
                "dependency_paths": [str(p) for p in dep_paths],
                "export_root": str(export_root),
                "asset_root": str(row.get("asset_root", "") or ""),
                "asset_id": str(row.get("id", "") or ""),
                "category": str(row.get("category", "") or ""),
                "export_layout": export_layout,
                "registry_path": str(registry_path),
                "bundle": bundle,
                "status_path": str(status_path),
                "result_path": str(result_path),
            }
            job_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")

            with log_path.open("wb") as logf:
                proc = subprocess.Popen(
                    [sys.executable, str(Path(__file__).resolve()), "--_export-worker", str(job_path)],
                    stdout=logf,
                    stderr=subprocess.STDOUT,
                    cwd=str(Path(__file__).resolve().parent),
                )

                last_status = {}
                while proc.poll() is None:
                    current_status = _read_json_tolerant(status_path, default=None, retries=2)
                    if isinstance(current_status, dict):
                        last_status = current_status
                    cur = int(last_status.get("current", 0) or 0)
                    tot = int(last_status.get("total", 0) or 0)
                    wok = int(last_status.get("ok", 0) or 0)
                    wfail = int(last_status.get("failed", 0) or 0)
                    path = str(last_status.get("path", ""))
                    short = PurePosixPath(path.replace("\\", "/")).name if path else "准备中"
                    text = (
                        f"Export bundle {idx:,}/{len(jobs):,} | 对象 {cur:,}/{tot:,}"
                        f" | 当前成功 {wok:,} 失败 {wfail:,} | {short[:46]}"
                    )
                    width = max(last_width, len(text) + 2)
                    last_width = width
                    print("\r" + text.ljust(width), end="", flush=True)
                    time.sleep(0.12)

                code = proc.returncode

            result = _read_json_tolerant(result_path, default=None, retries=6)

            if code == 0 and result and result.get("status") == "ok":
                records = result.get("records", [])
                writer.writerows(records)
                mf.flush()
                total_ok += int(result.get("ok", 0) or 0)
                total_failed += int(result.get("failed", 0) or 0)
                _unlink_retry(log_path)
            else:
                total_failed += 1
                err = ""
                if result:
                    err = str(result.get("error", ""))
                try:
                    log_text = log_path.read_text(encoding="utf-8", errors="replace")[-8000:]
                except Exception:
                    log_text = ""
                with crash_log.open("a", encoding="utf-8") as cf:
                    cf.write(f"\n===== {bundle} | exit={code} =====\n{err}\n{log_text}\n")
                writer.writerow({
                    "bundle": bundle,
                    "asset_path": str(row.get("asset_root", "") or ""),
                    "exported_path": "",
                    "type": "BundleWorker",
                    "path_id": "",
                    "mode": "isolated-worker",
                    "status": f"worker-crashed-or-failed: exit={code}; see {crash_log.name}",
                })
                mf.flush()
                _unlink_retry(log_path)

            for p in (job_path, status_path, result_path):
                _unlink_retry(p)

            text = (
                f"Export bundle {idx:,}/{len(jobs):,} | 总成功对象 {total_ok:,}"
                f" | 总失败 {total_failed:,} | 已完成"
            )
            width = max(last_width, len(text) + 2)
            last_width = width
            print("\r" + text.ljust(width), end="", flush=True)

    print()
    _unlink_retry(registry_path)
    _rmtree_retry(work)
    # 没有真正 worker crash 时，不留下空失败日志。
    try:
        if crash_log.exists() and crash_log.stat().st_size == 0:
            _unlink_retry(crash_log)
    except OSError:
        pass
    return total_ok, total_failed, manifest, crash_log


# ---------------------------------------------------------------------------
# 启动流程
# ---------------------------------------------------------------------------

def choose_source_and_load(
    args,
    work_dir: Path,
) -> tuple[list[dict], str]:
    csv_candidates = find_csv_candidates()
    catalog_texts = find_catalog_text_candidates()
    catalog_bundles = find_catalog_bundle_candidates()

    print("\n========== 本地资源检查 ==========")
    if csv_candidates:
        print(f"[CSV]     找到 {len(csv_candidates)} 个，优先: {csv_candidates[0]}")
    else:
        print("[CSV]     未找到")

    if catalog_texts:
        print(f"[catalog] 找到 {len(catalog_texts)} 个明文 catalog，优先: {catalog_texts[0]}")
    elif catalog_bundles:
        print(f"[catalog] 找到 {len(catalog_bundles)} 个 catalog.bundle，优先: {catalog_bundles[0]}")
    else:
        print("[catalog] 未找到")

    # 明确指定参数时优先服从。
    if args.csv:
        path = Path(args.csv).expanduser().resolve()
        print(f"\n使用指定 CSV: {path}")
        return read_mapping_csv(path), "csv"

    if args.catalog:
        path = Path(args.catalog).expanduser().resolve()
        if path.suffix.lower() == ".bundle":
            extracted = work_dir / "catalog.json"
            print(f"\n从指定 catalog.bundle 提取: {path}")
            path = extract_catalog_from_bundle(path, extracted)
        print(f"\n分析指定 catalog: {path}")
        rows = analyze_catalog(path)
        maybe_generate_csv(args, rows, work_dir)
        return rows, "catalog"

    # 若扫描到任何 catalog，按要求询问是否更新。
    have_local_catalog = bool(catalog_texts or catalog_bundles)
    update = False
    if have_local_catalog:
        if args.update_catalog:
            update = True
        elif not args.no_update_prompt:
            update = ask_yes_no(
                "检测到本地 catalog，是否下载最新 catalog.bundle 并更新后使用？",
                default=False,
            )

    if update:
        catalog_path = download_and_extract_catalog(
            work_dir,
            bundle_url=args.catalog_bundle_url,
            timeout=args.timeout,
            retries=args.retries,
        )
        print("正在分析更新后的 catalog...")
        rows = analyze_catalog(catalog_path)
        maybe_generate_csv(args, rows, work_dir)
        return rows, "catalog(updated)"

    # 未更新时，CSV 优先以提高效率。
    if csv_candidates:
        csv_path = csv_candidates[0]
        print(f"\n使用 CSV 快速筛选: {csv_path}")
        print("CSV 模式无需重新解析 catalog，适合日常重复查询。")
        return read_mapping_csv(csv_path), "csv"

    # 无 CSV，但有明文 catalog。
    if catalog_texts:
        catalog_path = catalog_texts[0]
        print(f"\n未找到 CSV，开始分析本地 catalog: {catalog_path}")
        rows = analyze_catalog(catalog_path)
        maybe_generate_csv(args, rows, work_dir)
        return rows, "catalog"

    # 无 CSV、无明文 catalog，但有 catalog.bundle。
    if catalog_bundles:
        bundle_path = catalog_bundles[0]
        catalog_path = work_dir / "catalog.json"
        print(f"\n未找到 CSV/明文 catalog，从本地 catalog.bundle 提取: {bundle_path}")
        extract_catalog_from_bundle(bundle_path, catalog_path)
        rows = analyze_catalog(catalog_path)
        maybe_generate_csv(args, rows, work_dir)
        return rows, "catalog(bundle)"

    # 全部没有：自动下载并解包。
    print("\n没有找到 CSV 或 catalog，将自动下载线上 catalog.bundle。")
    catalog_path = download_and_extract_catalog(
        work_dir,
        bundle_url=args.catalog_bundle_url,
        timeout=args.timeout,
        retries=args.retries,
    )
    print("正在分析 catalog...")
    rows = analyze_catalog(catalog_path)
    maybe_generate_csv(args, rows, work_dir)
    return rows, "catalog(downloaded)"


def maybe_generate_csv(args, rows: Sequence[dict], work_dir: Path) -> None:
    if args.no_csv:
        return
    if args.make_csv:
        generate = True
    else:
        print(
            "\ncatalog 已分析完成。生成对照 CSV 后，"
            "下次启动可以直接读取 CSV，省去 catalog 解析过程，筛选会更快。"
        )
        generate = ask_yes_no("是否生成/更新 bundle_id_mapping.csv？", default=True)

    if generate:
        write_mapping_csv(work_dir / "bundle_id_mapping.csv", rows)


def prompt_queries(args) -> list[str]:
    if args.queries:
        queries = []
        for q in args.queries:
            queries.extend(split_queries(q))
        return queries

    print("\n========== 输入下载条件 ==========")
    print("编号示例: ch_1041005")
    print("类型/目录关键词示例: still  sister  character")
    print("多个条件可用逗号分隔，例如: ch_1041005,still,sister")
    print("可强制模式: id:ch_1041005  key:still")
    typed = input("请输入要下载的编号/类型/路径关键词: ").strip()
    return split_queries(typed)


def choose_export_layout(value: str | None) -> str:
    if value in {"original", "id-category", "id-flat"}:
        return value
    print("\n请选择 Export 存放方式：")
    print("  1. 原目录：保持 Assets/AssetBundles/... 的真实目录结构")
    print("  2. 按编号+类型：<ID>/<character|still|其他类型>/文件（不保留更深层目录）")
    print("  3. 按编号平铺：<ID>/文件（没有任何次级分类目录）")
    while True:
        try:
            choice = input("请选择 1/2/3 [1]: ").strip()
        except EOFError:
            choice = ""
        if choice in {"", "1"}:
            return "original"
        if choice == "2":
            return "id-category"
        if choice == "3":
            return "id-flat"
        print("请输入 1、2 或 3。")


def main(argv: Sequence[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="自动检查 CSV/catalog、下载 Addressables bundle，并可按真实 Assets 路径 Export"
    )
    ap.add_argument("queries", nargs="*", help="ID 或类型/路径关键词")
    ap.add_argument("--csv", help="强制使用指定映射 CSV")
    ap.add_argument("--catalog", help="强制使用指定 catalog.json/txt/catalog.bundle")
    ap.add_argument(
        "--catalog-bundle-url",
        default=CATALOG_BUNDLE_URL,
        help="catalog.bundle 下载地址",
    )
    ap.add_argument(
        "--base-url",
        default=DEFAULT_BUNDLE_BASE_URL,
        help="bundle 下载基址",
    )
    ap.add_argument(
        "-o", "--output-dir",
        default="downloaded_bundles",
        help="所有 bundle 统一保存目录，默认 downloaded_bundles",
    )
    ap.add_argument("--workers", type=int, default=8, help="并行下载线程数，默认 8")
    ap.add_argument("--retries", type=int, default=3, help="失败重试次数，默认 3")
    ap.add_argument("--timeout", type=int, default=60, help="网络超时秒数，默认 60")
    ap.add_argument("--force", action="store_true", help="覆盖已存在的 bundle")
    ap.add_argument("--yes", action="store_true", help="大量匹配时不询问确认")
    ap.add_argument("--dry-run", action="store_true", help="只显示匹配结果，不下载")
    ap.add_argument("--make-csv", action="store_true", help="catalog 分析后自动生成 CSV")
    ap.add_argument("--no-csv", action="store_true", help="catalog 分析后不生成 CSV")
    ap.add_argument("--update-catalog", action="store_true", help="发现本地 catalog 时自动在线更新")
    ap.add_argument("--no-update-prompt", action="store_true", help="不询问 catalog 更新，直接使用本地资源")
    ap.add_argument("--export", dest="export_assets", action="store_true", help="下载完成后自动 Export")
    ap.add_argument("--no-export", action="store_true", help="下载完成后不询问 Export")
    ap.add_argument("--export-dir", default="exported_assets", help="Export 根目录，默认 exported_assets")
    ap.add_argument("--export-layout", choices=["original", "id-category", "id-flat"], help="Export目录布局：original=原Assets路径，id-category=ID/类型/文件，id-flat=ID/文件")
    ap.add_argument("--no-install-export-deps", action="store_true", help="Export 缺依赖时不询问自动 pip 安装")
    ap.add_argument("--no-bundle-deps", action="store_true", help="不自动加入 catalog 已知的依赖 bundle")
    args = ap.parse_args(argv)

    if args.workers < 1:
        args.workers = 1

    work_dir = Path.cwd().resolve()
    print_export_dependency_status()
    output_dir = Path(args.output_dir).expanduser()
    if not output_dir.is_absolute():
        output_dir = work_dir / output_dir
    output_dir = output_dir.resolve()

    # 无论正常退出、提前 return 还是 Ctrl-C，都尽量清理本程序创建的临时工作文件。
    _cleanup_state = {"export_root": None}
    def _exit_cleanup():
        try:
            _cleanup_known_temp_files(
                output_dir=output_dir,
                export_root=_cleanup_state.get("export_root"),
                work_dir=work_dir,
            )
        except Exception:
            pass
    atexit.register(_exit_cleanup)

    try:
        rows, source = choose_source_and_load(args, work_dir)
    except Exception as e:
        print(f"\n[错误] 无法准备映射数据: {e}", file=sys.stderr)
        return 2

    print(f"\n映射数据来源: {source}")
    print(f"可用 bundle 记录: {len(rows):,}")

    queries = prompt_queries(args)
    if not queries:
        print("未输入下载条件。")
        return 1

    direct_matches = filter_rows(rows, queries)
    if not direct_matches:
        print("\n没有找到符合条件的 bundle。")
        print("提示：编号可使用 id:xxx 强制精确匹配；路径关键词可使用 key:xxx。")
        return 1

    unresolved_dependencies: list[str] = []
    if args.no_bundle_deps:
        matches = [dict(r) for r in direct_matches]
        for r in matches:
            r["_selected_reason"] = "query"
    else:
        matches, unresolved_dependencies = expand_with_bundle_dependencies(direct_matches, rows)

    names = build_local_names(matches)

    print("\n========== 匹配结果 ==========")
    print(f"条件: {', '.join(queries)}")
    print(f"直接命中 bundle: {len(direct_matches):,}")
    dep_added = len(matches) - len(direct_matches)
    if dep_added:
        print(f"为保证资源引用完整，自动加入依赖 bundle: {dep_added:,}")
    print(f"本次 bundle 总数: {len(matches):,}")
    if unresolved_dependencies:
        print(f"警告：有 {len(unresolved_dependencies)} 个 catalog 依赖名无法在当前映射表定位。")
    preview = min(len(matches), 8)
    for i, (row, local_name) in enumerate(zip(matches[:preview], names[:preview]), 1):
        print(
            f"{i:>3}. {local_name} | {row.get('bundle','')} | reason={row.get('_selected_reason','query')}"
        )
    if len(matches) > preview:
        print(f"... 共 {len(matches):,} 条，仅预览前 {preview} 条；下载阶段不逐条刷屏。")

    if args.dry_run:
        print("\ndry-run：未执行下载。")
        return 0

    if len(matches) >= 20 and not args.yes:
        if not ask_yes_no(
            f"本次将下载 {len(matches):,} 个 bundle 到同一目录：{output_dir}\n是否继续？",
            default=False,
        ):
            print("已取消。")
            return 0

    output_dir.mkdir(parents=True, exist_ok=True)

    print("\n========== 开始下载 ==========")
    print(f"保存目录: {output_dir}")
    print(f"并行线程: {args.workers}")
    print(f"下载基址: {args.base_url}")

    tracker = ProgressTracker(len(matches))
    records: list[dict] = []
    thread_errors: list[str] = []

    with ThreadPoolExecutor(max_workers=args.workers) as pool:
        futures = [
            pool.submit(
                download_bundle_job,
                row,
                local_name,
                output_dir,
                args.base_url,
                tracker,
                timeout=args.timeout,
                retries=args.retries,
                force=args.force,
            )
            for row, local_name in zip(matches, names)
        ]

        for fut in as_completed(futures):
            try:
                records.append(fut.result())
            except Exception as e:
                thread_errors.append(f"{type(e).__name__}: {e}")

    tracker.finish_line()
    if thread_errors:
        print(f"下载线程异常: {len(thread_errors)} 个（详细信息已省略，失败项见下载清单）")
    records.sort(key=lambda r: str(r.get("local_file", "")).casefold())
    write_manifest(output_dir, records)

    ok = sum(1 for r in records if r.get("status") == "downloaded")
    skipped = sum(1 for r in records if r.get("status") == "exists")
    failed = sum(1 for r in records if str(r.get("status", "")).startswith("failed"))

    print("\n========== 下载完成 ==========")
    print(f"成功: {ok}")
    print(f"已存在跳过: {skipped}")
    print(f"失败: {failed}")
    print(f"清单: {output_dir / 'download_manifest.csv'}")
    cleaned = _cleanup_known_temp_files(output_dir=output_dir, work_dir=work_dir)
    if cleaned:
        print(f"已清理无用临时文件: {len(cleaned)}")

    # ------------------------------------------------------------------
    # 下载完成后按要求询问 Export
    # ------------------------------------------------------------------
    do_export = False
    if args.export_assets:
        do_export = True
    elif not args.no_export:
        print(
            "\nExport 支持三种存放方式：原 Assets 目录、按 ID+类型分类、按 ID 平铺。\n"
            "原目录模式示例：Assets/AssetBundles/Characters/HighQuality/general/ch_1022002/"
            "ch_1022002_b_Atlas.asset\n"
            "Sprite/Texture 会输出完整 PNG（Sprite 不再裁切）；TextAsset 中 .atlas.txt 会导出为 .atlas，.skel.bytes 会导出为 .skel。\n"
            "AudioClip 会优先导出 OGG/WAV/M4A 等可播放文件；还会补扫没有 container 路径的常见资源。"
        )
        do_export = ask_yes_no("是否现在进行 Export？", default=False)

    if do_export:
        print("\n========== Export 前检查 ==========")
        deps_complete, missing_dep_bundles, uncertain_dep_rows = verify_downloaded_bundle_dependencies(
            matches, records
        )
        if deps_complete:
            print("[Bundle依赖] catalog 中可识别的依赖 bundle 均已下载/存在。")
        else:
            print(f"[Bundle依赖] 缺少 {len(missing_dep_bundles)} 个依赖 bundle：")
            for x in missing_dep_bundles[:20]:
                print("  -", x)
            if len(missing_dep_bundles) > 20:
                print(f"  ... 另有 {len(missing_dep_bundles)-20} 个")

        if uncertain_dep_rows:
            print(
                f"[Bundle依赖] 有 {uncertain_dep_rows} 条记录显示存在多 bundle 依赖，"
                "但当前 CSV 没有 dependency_bundles 字段/数据；这通常是旧版 CSV。\n"
                "建议以后用本版 catalog 重新生成 bundle_id_mapping.csv，以便完整校验并自动补下载依赖。"
            )

        if missing_dep_bundles and not ask_yes_no("依赖不完整，仍要继续 Export？", default=False):
            print("已跳过 Export。")
            return 3 if failed else 0

        if not ensure_export_dependencies(allow_install=not args.no_install_export_deps):
            print("Export 依赖不齐全，已跳过 Export；bundle 下载结果保留不变。")
            return 4 if failed == 0 else 3

        export_root = Path(args.export_dir).expanduser()
        if not export_root.is_absolute():
            export_root = work_dir / export_root
        export_root = export_root.resolve()
        _cleanup_state["export_root"] = export_root

        export_layout = choose_export_layout(args.export_layout)
        layout_desc = {
            "original": "原 Assets 目录结构",
            "id-category": "按 ID + 类型分类",
            "id-flat": "按 ID 平铺（无次级目录）",
        }[export_layout]
        print(f"[Export布局] {layout_desc}")

        try:
            ex_ok, ex_failed, ex_manifest, ex_crash_log = export_downloaded_bundles_isolated(
                direct_matches, rows, records, output_dir, export_root, export_layout
            )
        except Exception as e:
            print(f"[Export失败] {e}", file=sys.stderr)
            return 5

        print("\n========== Export 完成 ==========")
        print(f"Export 根目录: {export_root}")
        print(f"成功对象: {ex_ok:,}")
        print(f"失败对象/任务: {ex_failed:,}")
        print(f"Export 清单: {ex_manifest}")
        if ex_crash_log.exists():
            print(f"崩溃/失败日志: {ex_crash_log}")
        print("说明：每个主 bundle 独立子进程导出；Sprite 改为完整底层 Texture2D，不做裁切；AudioClip/Font/Mesh/Shader 等有专用导出，其他 container 对象 raw 保持原后缀。")
        cleaned = _cleanup_known_temp_files(output_dir=output_dir, export_root=export_root, work_dir=work_dir)
        if cleaned:
            print(f"已清理 Export 临时文件: {len(cleaned)}")

    return 0 if failed == 0 else 3


if __name__ == "__main__":
    # 内部 Export worker：由主进程启动，故意与主下载器进程隔离。
    if "--_export-worker" in sys.argv:
        try:
            i = sys.argv.index("--_export-worker")
            raise SystemExit(export_worker_main(Path(sys.argv[i + 1])))
        except BaseException as e:
            if isinstance(e, SystemExit):
                raise
            raise SystemExit(71)
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\n用户中断。")
        raise SystemExit(130)
