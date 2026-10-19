#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Unity Addressables Bundle Downloader / Export V7

正式整合版：
- CSV -> catalog -> 在线 catalog.bundle 自动准备。
- catalog 分析后同时生成 CSV（程序快速读取）和 XLSX（人工查看，带列宽/筛选/冻结首行）。
- 下载默认 20 线程；Export 默认 20 个隔离子进程 worker。
- 模糊筛选：sisters/highquality、ch_100021/highquality；/ 是 AND，逗号/分号是 OR。
- 下载前扫描已有 bundle 和历史 manifest，存在则完全跳过网络。
- 下载时每个线程独立实时进度条，不逐文件刷屏。
- 日志统一保存 logs/<session>/；下载目录只放 bundle，Export 目录只放资源。
- Export 仅两种布局：原 Assets 路径；按 ID 归档，多个版本时自动分版本目录。
- Sprite 导出完整底层 Texture2D，不做 SpriteRect 裁切。
- .atlas.txt -> .atlas；.skel.bytes -> .skel。
- AudioClip/VideoClip/Font/Mesh/Shader/Texture2DArray 等专用导出；PPtr 自动 deref。
- 结束后清理无用临时文件。
"""
from __future__ import annotations

import argparse
import atexit
import base64
import csv
import gzip
import hashlib
import importlib
import importlib.util
import io
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
from typing import Iterable, Sequence
from xml.sax.saxutils import escape as xml_escape

VERSION = "7.0"
CATALOG_BUNDLE_URL = "https://d3mya90gbacu0m.cloudfront.net/prod/StreamingAssets/aa/catalog.bundle"
DEFAULT_BUNDLE_BASE_URL = "https://d3mya90gbacu0m.cloudfront.net/prod/StreamingAssets/aa/WebGL/"
USER_AGENT = f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) AddressablesSmartDownloader/{VERSION}"
DEFAULT_DOWNLOAD_WORKERS = 20
DEFAULT_EXPORT_WORKERS = 20
PREVIEW_MAX = 50
ENTRY_SIZE = 28
ENTRY_STRUCT = struct.Struct("<7i")
CSV_FIELDS = [
    "bundle", "id", "version", "category", "asset_root", "asset_count",
    "confidence", "max_dependency_bundle_count", "dependency_bundles",
    "bundle_internal_id", "bundle_location_index",
]
XLSX_WIDTHS = [44, 20, 26, 20, 82, 13, 18, 22, 82, 76, 20]
EXPORT_PIP_REQUIREMENTS = ["UnityPy==1.25.3", "Pillow>=10"]
EXPORT_MANIFEST_FIELDS = ["bundle", "asset_path", "exported_path", "type", "path_id", "mode", "status"]
SEARCH_FIELDS = ("id", "category", "version", "asset_root", "bundle_internal_id", "bundle")
ID_LIKE_RE = re.compile(r"(?i)^[a-z][a-z0-9]*[_-]\d+[a-z0-9_-]*$")


class CatalogFormatError(RuntimeError):
    pass


class SessionLog:
    def __init__(self, root: Path):
        stamp = time.strftime("%Y%m%d_%H%M%S")
        self.root = root / stamp
        n = 1
        while self.root.exists():
            n += 1
            self.root = root / f"{stamp}_{n:02d}"
        self.root.mkdir(parents=True, exist_ok=True)
        self.app_log = self.root / "app.log"
        self.round_no = 0
        self.lock = threading.Lock()

    def log(self, text: str) -> None:
        with self.lock:
            with self.app_log.open("a", encoding="utf-8") as f:
                f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {text}\n")

    def next_round(self) -> int:
        self.round_no += 1
        return self.round_no

    def download_manifest(self, n: int) -> Path:
        return self.root / f"round_{n:03d}_download_manifest.csv"

    def export_manifest(self, n: int) -> Path:
        return self.root / f"round_{n:03d}_export_manifest.csv"

    def export_failures(self, n: int) -> Path:
        return self.root / f"round_{n:03d}_export_failures.log"

    def export_work(self, n: int) -> Path:
        return self.root / f".export_work_round_{n:03d}"


def clear_screen() -> None:
    if os.name == "nt":
        os.system("cls")
    else:
        print("\033[2J\033[H", end="", flush=True)


def terminal_width() -> int:
    try:
        return max(88, min(136, shutil.get_terminal_size((124, 44)).columns))
    except Exception:
        return 124


def ui_title(text: str, subtitle: str = "") -> None:
    w = terminal_width()
    print("=" * w)
    print(f" {text}")
    if subtitle:
        print(f" {subtitle}")
    print("=" * w)


def ui_section(text: str) -> None:
    print()
    print(f"[ {text} ]")
    print("-" * min(terminal_width(), 112))


def wait_enter(text: str = "按 Enter 继续...") -> None:
    try:
        input(text)
    except EOFError:
        pass


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


def human_size(n: int | None) -> str:
    if n is None:
        return "?"
    f = float(n)
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if f < 1024 or unit == "TB":
            return f"{int(f)}B" if unit == "B" else f"{f:.1f}{unit}"
        f /= 1024
    return str(n)


def safe_piece(value: str, fallback: str) -> str:
    s = (value or "").strip().replace(" ", "_")
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", s)
    s = re.sub(r"_+", "_", s).strip(" ._")
    return s or fallback


def compact_text(value) -> str:
    return re.sub(r"[\s_\-]+", "", str(value or "").casefold())


def truncate(value, width: int) -> str:
    s = str(value or "")
    if len(s) <= width:
        return s
    return s[:max(1, width - 3)] + "..."


def unique_paths(paths: Iterable[Path]) -> list[Path]:
    out, seen = [], set()
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


def cleanup_temp(output_dir: Path | None = None, work_dir: Path | None = None, extra_work: Sequence[Path] = ()) -> list[str]:
    removed = []
    for work in extra_work:
        if work and work.exists() and _rmtree_retry(work):
            removed.append(str(work))
    if output_dir and output_dir.exists():
        for part in output_dir.glob("*.part"):
            final = Path(str(part)[:-5])
            try:
                useless = final.exists() or part.stat().st_size == 0
            except OSError:
                useless = False
            # 非空、最终文件不存在的 part 是有效断点，保留。
            if useless and _unlink_retry(part):
                removed.append(str(part))
    if work_dir:
        for p in (work_dir / "catalog.json.tmp", work_dir / "catalog.txt.tmp", work_dir / "bundle_id_mapping.csv.tmp"):
            if p.exists() and _unlink_retry(p):
                removed.append(str(p))
        for p in work_dir.glob(".*.json.*.tmp"):
            if _unlink_retry(p):
                removed.append(str(p))
    return removed


# ---------------------------- Mapping CSV/XLSX ----------------------------

def write_mapping_csv(path: Path, rows: Sequence[dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    with tmp.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
        w.writeheader()
        for row in rows:
            w.writerow({k: row.get(k, "") for k in CSV_FIELDS})
    os.replace(tmp, path)


def _xlsx_col_letter(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _xlsx_cell(ref: str, value) -> str:
    if value is None:
        value = ""
    if isinstance(value, bool):
        return f'<c r="{ref}" t="b"><v>{1 if value else 0}</v></c>'
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return f'<c r="{ref}" t="n"><v>{value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{xml_escape(str(value))}</t></is></c>'


def write_mapping_xlsx(path: Path, rows: Sequence[dict]) -> None:
    """纯标准库写 XLSX：固定列宽、首行冻结、AutoFilter。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    row_count = len(rows) + 1
    last_col = _xlsx_col_letter(len(CSV_FIELDS))
    dimension = f"A1:{last_col}{row_count}"
    cols_xml = "".join(
        f'<col min="{i}" max="{i}" width="{w}" customWidth="1"/>'
        for i, w in enumerate(XLSX_WIDTHS, 1)
    )
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        f'<dimension ref="{dimension}"/>',
        '<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
        '<sheetFormatPr defaultRowHeight="15"/>',
        f'<cols>{cols_xml}</cols><sheetData>',
    ]
    headers = []
    for c, name in enumerate(CSV_FIELDS, 1):
        ref = f"{_xlsx_col_letter(c)}1"
        headers.append(f'<c r="{ref}" s="1" t="inlineStr"><is><t>{xml_escape(name)}</t></is></c>')
    parts.append('<row r="1" ht="22" customHeight="1">' + "".join(headers) + '</row>')
    for r_idx, row in enumerate(rows, 2):
        cells = []
        for c_idx, field in enumerate(CSV_FIELDS, 1):
            cells.append(_xlsx_cell(f"{_xlsx_col_letter(c_idx)}{r_idx}", row.get(field, "")))
        parts.append(f'<row r="{r_idx}">' + "".join(cells) + '</row>')
    parts.extend(['</sheetData>', f'<autoFilter ref="{dimension}"/>', '</worksheet>'])
    sheet_xml = "".join(parts)

    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>"""
    root_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""
    workbook = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Bundle Mapping" sheetId="1" r:id="rId1"/></sheets>
</workbook>"""
    workbook_rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>"""
    styles = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<fonts count="2"><font><sz val="10"/><name val="Aptos"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Aptos"/></font></fonts>
<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills>
<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf></cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""
    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types)
        zf.writestr("_rels/.rels", root_rels)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        zf.writestr("xl/styles.xml", styles)
        zf.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    os.replace(tmp, path)


def write_mapping_files(work_dir: Path, rows: Sequence[dict]) -> tuple[Path, Path]:
    csv_path = work_dir / "bundle_id_mapping.csv"
    xlsx_path = work_dir / "bundle_id_mapping.xlsx"
    stamp = time.strftime("%Y%m%d_%H%M%S")
    for p in (csv_path, xlsx_path):
        if p.exists():
            try:
                shutil.copy2(p, p.with_name(f"{p.stem}.backup_{stamp}{p.suffix}"))
            except Exception:
                pass
    write_mapping_csv(csv_path, rows)
    write_mapping_xlsx(xlsx_path, rows)
    return csv_path, xlsx_path

# ---------------------------- Source scanning ----------------------------

def _scan_roots() -> list[Path]:
    return unique_paths([p for p in (Path.cwd(), Path(__file__).resolve().parent) if p.exists()])


def looks_like_mapping_csv(path: Path) -> bool:
    try:
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            header = next(csv.reader(f), [])
        cols = {x.strip().casefold() for x in header}
        return {"bundle", "id", "version", "category"}.issubset(cols)
    except Exception:
        return False


def find_csv_candidates() -> list[Path]:
    out = []
    for root in _scan_roots():
        for name in ("bundle_id_mapping.csv", "bundle_id_mapping_compact.csv"):
            p = root / name
            if p.is_file() and looks_like_mapping_csv(p):
                out.append(p)
        for p in root.glob("*.csv"):
            if p.is_file() and looks_like_mapping_csv(p):
                out.append(p)
    out = unique_paths(out)
    out.sort(key=lambda p: (
        3 if p.name.lower() == "bundle_id_mapping.csv" else 2 if p.name.lower() == "bundle_id_mapping_compact.csv" else 1,
        p.stat().st_mtime if p.exists() else 0,
    ), reverse=True)
    return out


def find_catalog_text_candidates() -> list[Path]:
    out = []
    for root in _scan_roots():
        for pat in ("catalog*.json", "catalog*.txt"):
            out.extend(p for p in root.glob(pat) if p.is_file())
    out = unique_paths(out)
    out.sort(key=lambda p: (
        2 if p.name.lower() == "catalog.json" else 1 if p.name.lower() == "catalog.txt" else 0,
        p.stat().st_mtime if p.exists() else 0,
    ), reverse=True)
    return out


def find_catalog_bundle_candidates() -> list[Path]:
    out = []
    for root in _scan_roots():
        for pat in ("catalog.bundle", "catalog_*.bundle", "catalog-*.bundle"):
            out.extend(p for p in root.glob(pat) if p.is_file())
    out = unique_paths(out)
    out.sort(key=lambda p: p.stat().st_mtime if p.exists() else 0, reverse=True)
    return out


# ---------------------------- HTTP / catalog bundle ----------------------------

def download_file_simple(url: str, dest: Path, *, timeout=60, retries=3, force=True) -> Path:
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.exists() and not force:
        return dest
    part = dest.with_name(dest.name + ".part")
    last_error = None
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
                    mode, done = "ab", existing
                    total = existing + remain if remain is not None else None
                else:
                    mode, done, total = "wb", 0, remain
                last = 0.0
                with part.open(mode) as f:
                    while True:
                        chunk = resp.read(1024 * 1024)
                        if not chunk:
                            break
                        f.write(chunk)
                        done += len(chunk)
                        now = time.monotonic()
                        if now - last > 0.15:
                            pct = f"{done * 100 / total:6.2f}%" if total else ""
                            print(f"\rcatalog.bundle {human_size(done)}/{human_size(total)} {pct}".ljust(90), end="", flush=True)
                            last = now
                print()
                os.replace(part, dest)
                return dest
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
            last_error = e
            if attempt < retries:
                time.sleep(min(2 ** attempt, 8))
    raise RuntimeError(f"下载失败: {url}\n{last_error}")


def _read_cstring(data: bytes, off: int) -> tuple[str, int]:
    end = data.find(b"\x00", off)
    if end < 0:
        raise CatalogFormatError("UnityFS 字符串没有终止符")
    return data[off:end].decode("utf-8", errors="replace"), end + 1


def _align(off: int, alignment: int) -> int:
    return (off + alignment - 1) // alignment * alignment


def _lz4_block_decompress(src: bytes, expected_size: int | None = None) -> bytes:
    out = bytearray()
    i, n = 0, len(src)
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
            raise CatalogFormatError("LZ4 offset 损坏")
        match_offset = src[i] | (src[i + 1] << 8)
        i += 2
        if match_offset <= 0 or match_offset > len(out):
            raise CatalogFormatError("LZ4 offset 非法")
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
        raise CatalogFormatError(f"LZ4 长度不符: expected={expected_size}, actual={len(out)}")
    return bytes(out)


def _decompress_unity_block(src: bytes, compression: int, expected_size: int) -> bytes:
    if compression == 0:
        out = src
    elif compression in (2, 3):
        out = _lz4_block_decompress(src, expected_size)
    elif compression == 1:
        errors = []
        for fmt in (lzma.FORMAT_AUTO, lzma.FORMAT_ALONE):
            try:
                out = lzma.decompress(src, format=fmt)
                break
            except Exception as e:
                errors.append(str(e))
        else:
            raise CatalogFormatError("UnityFS LZMA 解压失败: " + " | ".join(errors))
    else:
        raise CatalogFormatError(f"不支持 UnityFS 压缩类型: {compression}")
    if len(out) != expected_size:
        raise CatalogFormatError(f"UnityFS block 长度不符: expected={expected_size}, actual={len(out)}")
    return out


def unpack_unityfs_nodes(data: bytes) -> list[tuple[str, bytes]]:
    if not data.startswith(b"UnityFS\x00"):
        raise CatalogFormatError("不是 UnityFS")
    off = len(b"UnityFS\x00")
    version = struct.unpack_from(">I", data, off)[0]
    off += 4
    _unity_version, off = _read_cstring(data, off)
    _unity_revision, off = _read_cstring(data, off)
    if version < 6:
        raise CatalogFormatError(f"暂不支持 UnityFS version {version}")
    _file_size = struct.unpack_from(">Q", data, off)[0]
    off += 8
    c_info_size = struct.unpack_from(">I", data, off)[0]
    off += 4
    u_info_size = struct.unpack_from(">I", data, off)[0]
    off += 4
    flags = struct.unpack_from(">I", data, off)[0]
    off += 4
    if version >= 7:
        off = _align(off, 16)
    info_at_end = bool(flags & 0x80)
    need_padding = bool(flags & 0x200)
    info_comp = flags & 0x3F
    if info_at_end:
        info_pos, data_pos = len(data) - c_info_size, off
    else:
        info_pos, data_pos = off, off + c_info_size
        if need_padding:
            data_pos = _align(data_pos, 16)
    info = _decompress_unity_block(data[info_pos:info_pos + c_info_size], info_comp, u_info_size)
    p = 16
    block_count = struct.unpack_from(">I", info, p)[0]
    p += 4
    blocks = []
    for _ in range(block_count):
        u_size = struct.unpack_from(">I", info, p)[0]
        c_size = struct.unpack_from(">I", info, p + 4)[0]
        b_flags = struct.unpack_from(">H", info, p + 8)[0]
        p += 10
        blocks.append((u_size, c_size, b_flags & 0x3F))
    node_count = struct.unpack_from(">I", info, p)[0]
    p += 4
    nodes_meta = []
    for _ in range(node_count):
        node_off = struct.unpack_from(">Q", info, p)[0]
        node_size = struct.unpack_from(">Q", info, p + 8)[0]
        node_flags = struct.unpack_from(">I", info, p + 16)[0]
        p += 20
        name, p = _read_cstring(info, p)
        nodes_meta.append((name, node_off, node_size, node_flags))
    stream = bytearray()
    cur = data_pos
    for u_size, c_size, comp in blocks:
        stream.extend(_decompress_unity_block(data[cur:cur + c_size], comp, u_size))
        cur += c_size
    stream_b = bytes(stream)
    nodes = []
    for name, node_off, node_size, _flags in nodes_meta:
        if node_off + node_size <= len(stream_b):
            nodes.append((name, stream_b[node_off:node_off + node_size]))
    if not nodes:
        nodes.append(("unityfs_stream", stream_b))
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
            elif c == 0x5C:
                escaped = True
            elif c == 0x22:
                in_string = False
        else:
            if c == 0x22:
                in_string = True
            elif c == 0x7B:
                depth += 1
            elif c == 0x7D:
                depth -= 1
                if depth == 0:
                    return data[start:i + 1]
    return None


def find_catalog_json_blob(data: bytes) -> bytes | None:
    stripped = data.lstrip(b"\xef\xbb\xbf \t\r\n\x00")
    if stripped.startswith(b"{"):
        try:
            obj = json.loads(stripped.decode("utf-8"))
            if isinstance(obj, dict) and "m_InternalIds" in obj:
                return stripped
        except Exception:
            pass
    for pat in (
        re.compile(rb'\{\s*"m_LocatorId"\s*:'),
        re.compile(rb'\{\s*"m_BuildResultHash"\s*:'),
        re.compile(rb'\{\s*"m_InternalIdPrefixes"\s*:'),
    ):
        for m in pat.finditer(data):
            blob = _extract_balanced_json(data, m.start())
            if not blob:
                continue
            try:
                obj = json.loads(blob.decode("utf-8"))
                if isinstance(obj, dict) and "m_InternalIds" in obj:
                    return blob
            except Exception:
                pass
    return None


def extract_catalog_from_bundle(bundle_path: Path, output_path: Path) -> Path:
    data = bundle_path.read_bytes()
    blob = find_catalog_json_blob(data)
    if blob:
        output_path.write_bytes(blob)
        return output_path
    if data.startswith(b"\x1f\x8b"):
        try:
            blob = find_catalog_json_blob(gzip.decompress(data))
            if blob:
                output_path.write_bytes(blob)
                return output_path
        except Exception:
            pass
    if data.startswith(b"PK\x03\x04"):
        try:
            with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
                for name in zf.namelist():
                    blob = find_catalog_json_blob(zf.read(name))
                    if blob:
                        output_path.write_bytes(blob)
                        return output_path
        except Exception:
            pass
    if data.startswith(b"UnityFS\x00"):
        nodes = unpack_unityfs_nodes(data)
        for _name, raw in nodes:
            blob = find_catalog_json_blob(raw)
            if blob:
                output_path.write_bytes(blob)
                return output_path
        blob = find_catalog_json_blob(b"\n".join(raw for _, raw in nodes))
        if blob:
            output_path.write_bytes(blob)
            return output_path
    raise CatalogFormatError("未能从 catalog.bundle 找到 Addressables catalog JSON")


def download_and_extract_catalog(work_dir: Path, url: str, timeout: int, retries: int) -> Path:
    bundle_path = work_dir / "catalog.bundle"
    catalog_path = work_dir / "catalog.json"
    tmp = work_dir / "catalog.json.tmp"
    tmp.unlink(missing_ok=True)
    print(f"下载：{url}")
    download_file_simple(url, bundle_path, timeout=timeout, retries=retries, force=True)
    print("正在提取 catalog...")
    extract_catalog_from_bundle(bundle_path, tmp)
    with tmp.open("r", encoding="utf-8-sig") as f:
        obj = json.load(f)
    if not isinstance(obj, dict) or "m_InternalIds" not in obj:
        tmp.unlink(missing_ok=True)
        raise CatalogFormatError("提取出的 JSON 不是有效 Addressables catalog")
    os.replace(tmp, catalog_path)
    return catalog_path


# ---------------------------- Catalog analysis ----------------------------

def _find_json_value_span(mm: mmap.mmap, key: str) -> tuple[int, int]:
    needle = f'"{key}"'.encode()
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
    start, first = p, mm[p]
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
        raise CatalogFormatError(f"字符串未结束: {key}")
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
        raise CatalogFormatError(f"JSON 值未结束: {key}")
    while p < n and mm[p] not in b",}\r\n":
        p += 1
    return start, p


def load_required_catalog_fields(path: Path):
    with path.open("rb") as f:
        with mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) as mm:
            vals = {}
            for key in ("m_InternalIdPrefixes", "m_InternalIds", "m_BucketDataString", "m_EntryDataString"):
                a, b = _find_json_value_span(mm, key)
                vals[key] = json.loads(mm[a:b].decode("utf-8"))
    return vals["m_InternalIdPrefixes"], vals["m_InternalIds"], vals["m_BucketDataString"], vals["m_EntryDataString"]


def resolve_internal_id(raw: str, prefixes: Sequence[str]) -> str:
    m = re.match(r"^(\d+)#(.*)$", str(raw), re.S)
    if not m:
        return str(raw).replace("\\", "/")
    idx, suffix = int(m.group(1)), m.group(2)
    if 0 <= idx < len(prefixes):
        return (prefixes[idx] + suffix).replace("\\", "/")
    return str(raw).replace("\\", "/")


def _bundle_filename(raw: str) -> str:
    s = str(raw).replace("\\", "/")
    if "#/" in s:
        s = s.split("#/", 1)[1]
    return s.rsplit("/", 1)[-1]


def _decode_entries(b64: str) -> tuple[bytes, int]:
    data = base64.b64decode(b64, validate=True)
    if len(data) < 4:
        raise CatalogFormatError("EntryData 太短")
    count = struct.unpack_from("<I", data, 0)[0]
    if len(data) != 4 + count * ENTRY_SIZE:
        raise CatalogFormatError("EntryData 大小与 count 不符")
    return data, count


def _decode_buckets(b64: str) -> tuple[bytes, int]:
    data = base64.b64decode(b64, validate=True)
    if len(data) < 4:
        raise CatalogFormatError("BucketData 太短")
    return data, struct.unpack_from("<I", data, 0)[0]


def infer_semantics(asset_paths: Sequence[str]):
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
    category = version = asset_id = ""
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


def analyze_catalog(path: Path) -> list[dict]:
    prefixes, internal_ids, bucket_b64, entry_b64 = load_required_catalog_fields(path)
    entry_data, entry_count = _decode_entries(entry_b64)
    bucket_data, bucket_count = _decode_buckets(bucket_b64)
    bundle_loc_to_internal = {}
    for loc in range(entry_count):
        internal_idx = ENTRY_STRUCT.unpack_from(entry_data, 4 + loc * ENTRY_SIZE)[0]
        if 0 <= internal_idx < len(internal_ids):
            raw = internal_ids[internal_idx]
            if isinstance(raw, str) and raw.lower().endswith(".bundle"):
                bundle_loc_to_internal[loc] = internal_idx
    if not bundle_loc_to_internal:
        raise CatalogFormatError("catalog 中没有 .bundle location")
    bundle_set = set(bundle_loc_to_internal)
    owner_by_key = array("i", [-1]) * bucket_count
    dep_count_by_key = array("I", [0]) * bucket_count
    dep_bundle_locs_by_key = {}
    off = 4
    for bucket_idx in range(bucket_count):
        if off + 8 > len(bucket_data):
            raise CatalogFormatError(f"BucketData 在 {bucket_idx} 截断")
        _key_off, location_count = struct.unpack_from("<II", bucket_data, off)
        off += 8
        byte_count = location_count * 4
        if off + byte_count > len(bucket_data):
            raise CatalogFormatError(f"Bucket locations 在 {bucket_idx} 截断")
        owner = -1
        bundle_locs = []
        for j in range(location_count):
            loc = struct.unpack_from("<I", bucket_data, off + j * 4)[0]
            if loc in bundle_set:
                bundle_locs.append(loc)
                if owner < 0:
                    owner = loc
        owner_by_key[bucket_idx] = owner
        dep_count_by_key[bucket_idx] = len(bundle_locs)
        if bundle_locs:
            dep_bundle_locs_by_key[bucket_idx] = tuple(bundle_locs)
        off += byte_count
    assets_by_bundle = defaultdict(set)
    deps_by_bundle = defaultdict(set)
    max_dep_count = defaultdict(int)
    for loc in range(entry_count):
        internal_idx, _provider, dep_key, _dep_hash, _data_idx, _primary_key, _rtype = ENTRY_STRUCT.unpack_from(entry_data, 4 + loc * ENTRY_SIZE)
        if loc in bundle_set or not (0 <= internal_idx < len(internal_ids)) or not (0 <= dep_key < bucket_count):
            continue
        owner = owner_by_key[dep_key]
        if owner < 0:
            continue
        assets_by_bundle[owner].add(internal_idx)
        for dep_loc in dep_bundle_locs_by_key.get(dep_key, ()):
            if dep_loc != owner:
                deps_by_bundle[owner].add(dep_loc)
        max_dep_count[owner] = max(max_dep_count[owner], dep_count_by_key[dep_key])
    bundle_name_by_loc = {loc: _bundle_filename(internal_ids[idx]) for loc, idx in bundle_loc_to_internal.items()}
    rows = []
    for bundle_loc in sorted(bundle_loc_to_internal):
        internal_idx = bundle_loc_to_internal[bundle_loc]
        raw = internal_ids[internal_idx]
        asset_indexes = sorted(assets_by_bundle.get(bundle_loc, ()))
        asset_paths = [resolve_internal_id(internal_ids[i], prefixes) for i in asset_indexes]
        asset_id, version, category, root, confidence = infer_semantics(asset_paths)
        rows.append({
            "bundle": _bundle_filename(raw),
            "id": asset_id,
            "version": version,
            "category": category,
            "asset_root": root,
            "asset_count": len(asset_paths),
            "confidence": confidence,
            "max_dependency_bundle_count": max_dep_count.get(bundle_loc, 0),
            "dependency_bundles": ";".join(bundle_name_by_loc[x] for x in sorted(deps_by_bundle.get(bundle_loc, ())) if x in bundle_name_by_loc),
            "bundle_internal_id": resolve_internal_id(raw, prefixes),
            "bundle_location_index": bundle_loc,
        })
    return rows


def read_mapping_csv(path: Path) -> list[dict]:
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        missing = {"bundle", "id", "version", "category"} - set(reader.fieldnames or [])
        if missing:
            raise ValueError("CSV 缺少字段: " + ", ".join(sorted(missing)))
        rows = []
        for r in reader:
            item = dict(r)
            for key in CSV_FIELDS:
                item.setdefault(key, "")
            rows.append(item)
        return rows


def maybe_generate_mapping_files(args, rows: Sequence[dict], work_dir: Path) -> None:
    if args.no_csv:
        return
    if args.make_csv:
        generate = True
    else:
        print()
        print("catalog 已分析完成。")
        print("CSV 本身没有列宽属性；本版会同时生成 CSV + XLSX 查看版。")
        print("XLSX 已设置列宽、筛选、冻结首行，便于人工检查。")
        generate = ask_yes_no("是否生成/更新映射表？", default=True)
    if generate:
        csv_path, xlsx_path = write_mapping_files(work_dir, rows)
        print(f"CSV : {csv_path}")
        print(f"XLSX: {xlsx_path}")
        print("下次启动可直接读取 CSV，加快筛选。")


def prepare_mapping(args, work_dir: Path, session: SessionLog):
    csvs = find_csv_candidates()
    catalogs = find_catalog_text_candidates()
    bundles = find_catalog_bundle_candidates()
    clear_screen()
    ui_title("前置准备", f"下载线程 {args.workers} | Export线程 {args.export_workers} | 日志 {session.root}")
    ui_section("本地检查")
    print(f"CSV     : {csvs[0] if csvs else '未找到'}")
    if catalogs:
        print(f"catalog : {catalogs[0]}")
    elif bundles:
        print(f"catalog : {bundles[0]} (bundle)")
    else:
        print("catalog : 未找到")
    if args.csv:
        p = Path(args.csv).expanduser().resolve()
        session.log(f"use explicit csv: {p}")
        return read_mapping_csv(p), f"CSV / {p.name}"
    if args.catalog:
        p = Path(args.catalog).expanduser().resolve()
        if p.suffix.lower() == ".bundle":
            out = work_dir / "catalog.json"
            extract_catalog_from_bundle(p, out)
            p = out
        print(f"\n分析指定 catalog: {p}")
        rows = analyze_catalog(p)
        maybe_generate_mapping_files(args, rows, work_dir)
        return rows, f"catalog / {p.name}"
    have_catalog = bool(catalogs or bundles)
    update = False
    if have_catalog:
        if args.update_catalog:
            update = True
        elif not args.no_update_prompt:
            print()
            update = ask_yes_no("检测到本地 catalog，是否先下载最新 catalog.bundle 更新？", default=False)
    if update:
        clear_screen()
        ui_title("更新 catalog")
        p = download_and_extract_catalog(work_dir, args.catalog_bundle_url, args.timeout, args.retries)
        print("分析更新后的 catalog...")
        rows = analyze_catalog(p)
        maybe_generate_mapping_files(args, rows, work_dir)
        return rows, "catalog / updated"
    if csvs:
        p = csvs[0]
        session.log(f"use csv: {p}")
        return read_mapping_csv(p), f"CSV / {p.name}"
    if catalogs:
        p = catalogs[0]
        print(f"\n没有 CSV，分析本地 catalog: {p}")
        rows = analyze_catalog(p)
        maybe_generate_mapping_files(args, rows, work_dir)
        return rows, f"catalog / {p.name}"
    if bundles:
        p = bundles[0]
        out = work_dir / "catalog.json"
        print(f"\n从本地 catalog.bundle 提取: {p}")
        extract_catalog_from_bundle(p, out)
        rows = analyze_catalog(out)
        maybe_generate_mapping_files(args, rows, work_dir)
        return rows, "catalog / local bundle"
    clear_screen()
    ui_title("获取 catalog", "本地没有 CSV / catalog，将自动下载")
    p = download_and_extract_catalog(work_dir, args.catalog_bundle_url, args.timeout, args.retries)
    rows = analyze_catalog(p)
    maybe_generate_mapping_files(args, rows, work_dir)
    return rows, "catalog / downloaded"

# ---------------------------- Fuzzy query / dependencies ----------------------------

def parse_query_groups(text: str) -> list[list[str]]:
    groups = []
    for group in re.split(r"[,;]+", str(text or "")):
        tokens = [x.strip() for x in group.split("/") if x.strip()]
        if tokens:
            groups.append(tokens)
    return groups


def token_matches_row(row: dict, token: str) -> bool:
    raw = token.strip()
    low = raw.casefold()
    if low.startswith("id:"):
        return compact_text(row.get("id", "")) == compact_text(raw.split(":", 1)[1])
    if raw.startswith("="):
        return compact_text(row.get("id", "")) == compact_text(raw[1:])
    if low.startswith(("key:", "path:", "type:")):
        raw = raw.split(":", 1)[1]
    needle = compact_text(raw)
    if not needle:
        return True
    return any(needle in compact_text(row.get(field, "")) for field in SEARCH_FIELDS)


def fuzzy_filter_rows(rows: Sequence[dict], query: str) -> list[dict]:
    groups = parse_query_groups(query)
    if not groups:
        return []
    out, seen = [], set()
    for row in rows:
        if not any(all(token_matches_row(row, t) for t in group) for group in groups):
            continue
        bundle = str(row.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
        if not bundle:
            continue
        key = bundle.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(dict(row))
    out.sort(key=lambda r: (
        compact_text(r.get("category", "")), compact_text(r.get("id", "")),
        compact_text(r.get("version", "")), compact_text(r.get("bundle", "")),
    ))
    return out


def _parse_dependency_bundles(row: dict) -> list[str]:
    raw = str(row.get("dependency_bundles", "") or "")
    return [x.strip().replace("\\", "/").rsplit("/", 1)[-1] for x in raw.split(";") if x.strip()]


def expand_with_dependencies(direct: Sequence[dict], all_rows: Sequence[dict]):
    row_by_bundle = {
        str(r.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1].casefold(): dict(r)
        for r in all_rows if str(r.get("bundle", "")).strip()
    }
    selected, queue = {}, []
    for r in direct:
        item = dict(r)
        item["_selected_reason"] = "query"
        b = str(item.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
        if b:
            selected[b.casefold()] = item
            queue.append(item)
    unresolved, unresolved_seen = [], set()
    pos = 0
    while pos < len(queue):
        row = queue[pos]
        pos += 1
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
            queue.append(dep_row)
    result = list(selected.values())
    result.sort(key=lambda r: (
        0 if r.get("_selected_reason") == "query" else 1,
        compact_text(r.get("category", "")), compact_text(r.get("id", "")),
        compact_text(r.get("version", "")), compact_text(r.get("bundle", "")),
    ))
    return result, unresolved


def build_local_names(rows: Sequence[dict]) -> list[str]:
    names, used = [], set()
    for r in rows:
        bundle = str(r.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
        if r.get("_selected_reason") == "dependency" or not str(r.get("id", "")).strip():
            name = f"dependency_{safe_piece(Path(bundle).stem, 'bundle')}.bundle"
        else:
            rid = safe_piece(str(r.get("id", "")), "unknown-id")
            cat = safe_piece(str(r.get("category", "")), "unknown-category")
            ver = safe_piece(str(r.get("version", "")), "unknown-version")
            name = f"{rid}_{cat}_{ver}.bundle"
        if name.casefold() in used:
            short = hashlib.sha1(bundle.encode("utf-8")).hexdigest()[:8]
            stem = name[:-7] if name.lower().endswith(".bundle") else name
            name = f"{stem}__{short}.bundle"
        used.add(name.casefold())
        names.append(name)
    return names


def print_query_help() -> None:
    print("示例：")
    print("  ch_100021                  编号相关全部 bundle")
    print("  ch_100021/highquality      该编号中只取 HighQuality")
    print("  sisters/highquality        Sisters 下所有 HighQuality")
    print("  still/highquality          Still 下所有 HighQuality")
    print("  sisters/highquality,still/highquality")
    print("                              两组结果取并集")
    print()
    print('规则： "/" = AND    "," / ";" = OR    id:xxx = 精确 ID')
    print("普通词为模糊包含，并忽略大小写、空格、下划线和连字符差异。")


def print_preview(query: str, direct: Sequence[dict], matches: Sequence[dict], unresolved: Sequence[str], already_count: int) -> None:
    clear_screen()
    ui_title("下载预览", f"筛选：{query}")
    print(f"直接命中       : {len(direct):,}")
    print(f"自动依赖       : {len(matches) - len(direct):,}")
    print(f"本地已存在     : {already_count:,}")
    print(f"预计网络下载   : {max(0, len(matches) - already_count):,}")
    if unresolved:
        print(f"未解析依赖     : {len(unresolved):,}")
    ui_section(f"关键清单（最多 {PREVIEW_MAX} 行，仅显示直接匹配资源）")
    shown = min(len(direct), PREVIEW_MAX)
    print(f"{'#':>3}  {'编号':<20}  {'完整路径':<70}  类型")
    print("-" * min(terminal_width(), 132))
    for i, row in enumerate(direct[:shown], 1):
        typ = " / ".join(x for x in (str(row.get("category", "")).strip(), str(row.get("version", "")).strip()) if x)
        print(f"{i:>3}  {truncate(row.get('id', ''), 20):<20}  {truncate(row.get('asset_root', ''), 70):<70}  {truncate(typ, 30)}")
    if len(direct) > shown:
        print(f"\n... 另有 {len(direct) - shown:,} 条直接匹配未显示。")


# ---------------------------- Existing bundle preflight ----------------------------

def load_historical_bundle_index(output_dir: Path, logs_root: Path) -> dict[str, Path]:
    index = {}
    if output_dir.exists():
        for p in output_dir.glob("*.bundle"):
            try:
                if p.is_file() and p.stat().st_size > 0:
                    index[p.name.casefold()] = p
            except OSError:
                pass
    manifests = []
    old = output_dir / "download_manifest.csv"
    if old.is_file():
        manifests.append(old)
    if logs_root.exists():
        manifests.extend(logs_root.rglob("*download_manifest*.csv"))
    for mf in manifests:
        try:
            with mf.open("r", encoding="utf-8-sig", newline="") as f:
                for r in csv.DictReader(f):
                    original = str(r.get("original_bundle", "") or "").replace("\\", "/").rsplit("/", 1)[-1]
                    local = str(r.get("local_file", "") or "")
                    if not original or not local:
                        continue
                    p = output_dir / local
                    if p.is_file() and p.stat().st_size > 0:
                        index[original.casefold()] = p
        except Exception:
            continue
    return index


def preflight_existing(rows: Sequence[dict], names: Sequence[str], output_dir: Path, logs_root: Path, base_url: str) -> tuple[list[dict], list[tuple[dict, str]]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    index = load_historical_bundle_index(output_dir, logs_root)
    existing_records, need_network = [], []
    for row, desired_name in zip(rows, names):
        bundle = str(row.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
        dest = output_dir / desired_name
        found = None
        try:
            if dest.is_file() and dest.stat().st_size > 0:
                found = dest
        except OSError:
            pass
        if found is None:
            found = index.get(bundle.casefold())
        if found is None:
            hash_named = output_dir / bundle
            try:
                if hash_named.is_file() and hash_named.stat().st_size > 0:
                    found = hash_named
            except OSError:
                pass
        if found is not None:
            if found != dest and not dest.exists():
                try:
                    found.replace(dest)
                    found = dest
                except OSError:
                    pass
            existing_records.append({
                "id": row.get("id", ""), "category": row.get("category", ""), "version": row.get("version", ""),
                "selected_reason": row.get("_selected_reason", "query"), "original_bundle": bundle,
                "url": urllib.parse.urljoin(base_url.rstrip("/") + "/", urllib.parse.quote(bundle, safe="")),
                "local_file": found.name, "status": "exists",
            })
        else:
            need_network.append((dict(row), desired_name))
    return existing_records, need_network


# ---------------------------- Live download UI ----------------------------

class DownloadLiveUI:
    def __init__(self, total_network: int, skipped_local: int, workers: int):
        self.total = total_network
        self.skipped_local = skipped_local
        self.workers = workers
        self.lock = threading.Lock()
        self.states = {i: {"name": "", "done": 0, "total": None, "start_done": 0, "start_time": 0.0, "status": "idle"} for i in range(1, workers + 1)}
        self.completed = 0
        self.failed = 0
        self.stop_event = threading.Event()
        self.renderer = threading.Thread(target=self._render_loop, daemon=True)

    @staticmethod
    def slot() -> int:
        m = re.search(r"_(\d+)$", threading.current_thread().name)
        return (int(m.group(1)) + 1) if m else 1

    def start(self):
        self.renderer.start()

    def begin_file(self, slot: int, name: str, done=0, total=None):
        with self.lock:
            self.states[slot] = {"name": name, "done": done, "total": total, "start_done": done, "start_time": time.monotonic(), "status": "active"}

    def update(self, slot: int, done: int, total):
        with self.lock:
            st = self.states[slot]
            st["done"] = done
            st["total"] = total

    def finish(self, slot: int, success: bool):
        with self.lock:
            self.states[slot]["status"] = "done" if success else "failed"
            if success:
                self.completed += 1
            else:
                self.failed += 1

    def stop(self):
        self.stop_event.set()
        self.renderer.join(timeout=2)
        self.render()
        print()

    def _render_loop(self):
        while not self.stop_event.is_set():
            self.render()
            self.stop_event.wait(0.12)

    def render(self):
        with self.lock:
            states = {k: dict(v) for k, v in self.states.items()}
            completed, failed = self.completed, self.failed
        clear_screen()
        ui_title("Bundle 下载", f"网络任务 {self.total:,} | 本地跳过 {self.skipped_local:,} | 下载线程 {self.workers}")
        print(f"完成 {completed:,}/{self.total:,}  | 失败 {failed:,}  | 本地已存在 {self.skipped_local:,}")
        print()
        bar_w = 22
        for slot in range(1, self.workers + 1):
            st = states[slot]
            name = truncate(st["name"] or "空闲", 42)
            done, total, status = int(st["done"] or 0), st["total"], st["status"]
            if total and total > 0:
                pct = max(0.0, min(100.0, done * 100.0 / total))
                filled = int(bar_w * pct / 100)
                bar = "█" * filled + "░" * (bar_w - filled)
                pct_text = f"{pct:6.2f}%"
                bytes_text = f"{human_size(done)}/{human_size(total)}"
            else:
                bar, pct_text, bytes_text = "░" * bar_w, "   --  ", human_size(done)
            if st["start_time"]:
                elapsed = max(0.001, time.monotonic() - st["start_time"])
                speed = max(0, done - int(st["start_done"] or 0)) / elapsed
            else:
                speed = 0
            label = {"idle": "空闲", "active": "下载", "done": "完成", "failed": "失败"}.get(status, status)
            print(f"T{slot:02d} {label:<4} [{bar}] {pct_text} {bytes_text:>17} {human_size(int(speed)):>8}/s  {name}")


def download_bundle_job(row: dict, local_name: str, output_dir: Path, base_url: str, ui: DownloadLiveUI, session: SessionLog, *, timeout: int, retries: int, force: bool) -> dict:
    bundle = str(row.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
    url = urllib.parse.urljoin(base_url.rstrip("/") + "/", urllib.parse.quote(bundle, safe=""))
    dest = output_dir / local_name
    part = dest.with_name(dest.name + ".part")
    slot = ui.slot()
    record = {
        "id": row.get("id", ""), "category": row.get("category", ""), "version": row.get("version", ""),
        "selected_reason": row.get("_selected_reason", "query"), "original_bundle": bundle,
        "url": url, "local_file": local_name, "status": "",
    }
    if force:
        dest.unlink(missing_ok=True)
    last_error = None
    for attempt in range(1, retries + 1):
        existing = part.stat().st_size if part.exists() else 0
        ui.begin_file(slot, local_name, existing, None)
        headers = {"User-Agent": USER_AGENT, "Accept-Encoding": "identity"}
        if existing:
            headers["Range"] = f"bytes={existing}-"
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=timeout) as resp:
                status = getattr(resp, "status", resp.getcode())
                clen = resp.headers.get("Content-Length")
                remain = int(clen) if clen and clen.isdigit() else None
                if existing and status == 206:
                    mode, done = "ab", existing
                    total = existing + remain if remain is not None else None
                else:
                    mode, done, total = "wb", 0, remain
                ui.begin_file(slot, local_name, done, total)
                with part.open(mode) as f:
                    while True:
                        chunk = resp.read(1024 * 1024)
                        if not chunk:
                            break
                        f.write(chunk)
                        done += len(chunk)
                        ui.update(slot, done, total)
                os.replace(part, dest)
                size = dest.stat().st_size
                ui.update(slot, size, size)
                ui.finish(slot, True)
                record["status"] = "downloaded"
                session.log(f"download ok: {bundle} -> {local_name}")
                return record
        except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError) as e:
            last_error = e
            session.log(f"download attempt {attempt} failed: {bundle}: {e}")
            if attempt < retries:
                time.sleep(min(2 ** attempt, 8))
    ui.finish(slot, False)
    record["status"] = f"failed: {last_error}"
    session.log(f"download failed: {bundle}: {last_error}")
    return record


def write_download_manifest(path: Path, records: Sequence[dict]) -> None:
    fields = ["id", "category", "version", "selected_reason", "original_bundle", "url", "local_file", "status"]
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(records)

# ---------------------------- Export dependencies / object helpers ----------------------------

def export_dependency_status():
    problems = []
    for module, display in (("UnityPy", "UnityPy"), ("PIL", "Pillow")):
        if importlib.util.find_spec(module) is None:
            problems.append(display + "(未安装)")
            continue
        try:
            importlib.import_module(module)
        except Exception as e:
            problems.append(f"{display}(导入失败: {type(e).__name__})")
    return not problems, problems


def ensure_export_dependencies(allow_install=True) -> bool:
    ok, problems = export_dependency_status()
    if ok:
        return True
    print("缺少 Export 依赖: " + ", ".join(problems))
    if not allow_install:
        return False
    if not ask_yes_no("是否自动安装 " + " ".join(EXPORT_PIP_REQUIREMENTS) + "？", default=True):
        return False
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", *EXPORT_PIP_REQUIREMENTS])
    except Exception as e:
        print(f"安装失败: {e}")
        return False
    importlib.invalidate_caches()
    return export_dependency_status()[0]


def _safe_container_relative_path(path: str) -> PurePosixPath:
    raw = str(path).replace("\\", "/").strip()
    m = re.search(r"(?i)(?:^|/)(Assets/.*)$", raw)
    raw = m.group(1) if m else raw.lstrip("/")
    parts = []
    for part in PurePosixPath(raw).parts:
        if part in ("", "."):
            continue
        if part == "..":
            raise ValueError("不安全路径")
        parts.append(re.sub(r'[<>:"|?*\x00-\x1f]', "_", part))
    if not parts:
        raise ValueError("空资源路径")
    return PurePosixPath(*parts)


def _image_output_relative_path(rel: PurePosixPath) -> PurePosixPath:
    stem, _ = posixpath.splitext(rel.name)
    return rel.with_name(stem + ".png")


def _replace_last_extension(rel: PurePosixPath, ext: str) -> PurePosixPath:
    stem, _ = posixpath.splitext(rel.name)
    if ext and not ext.startswith("."):
        ext = "." + ext
    return rel.with_name(stem + ext)


def _textasset_output_relative_path(rel: PurePosixPath) -> PurePosixPath:
    low = rel.name.lower()
    if low.endswith(".atlas.txt"):
        return rel.with_name(rel.name[:-4])
    if low.endswith(".skel.bytes"):
        return rel.with_name(rel.name[:-6])
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


def _resolve_object_reader(obj):
    if obj is None:
        raise ValueError("对象为空")
    if hasattr(obj, "get_raw_data") and hasattr(obj, "parse_as_object"):
        return obj
    deref = getattr(obj, "deref", None)
    if callable(deref):
        return deref()
    return obj


def _object_reader_key(obj):
    try:
        reader = _resolve_object_reader(obj)
        return (id(getattr(reader, "assets_file", None)), int(getattr(reader, "path_id", 0) or 0))
    except Exception:
        return (id(getattr(obj, "assets_file", None)), int(getattr(obj, "path_id", 0) or 0))


def _pptr_valid(ptr) -> bool:
    if ptr is None:
        return False
    try:
        return int(getattr(ptr, "path_id", getattr(ptr, "m_PathID", 0)) or 0) != 0
    except Exception:
        return False


def _sprite_render_data(sprite):
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
                    pass
    except Exception:
        pass
    return getattr(sprite, "m_RD", None)


def _full_image_from_sprite(sprite):
    rd = _sprite_render_data(sprite)
    if rd is None:
        raise ValueError("Sprite 没有 RenderData")
    texture_ptr = getattr(rd, "texture", None)
    if not _pptr_valid(texture_ptr):
        raise ValueError("Sprite 没有 Texture2D")
    texture = texture_ptr.deref().parse_as_object()
    image = texture.image  # 完整 Texture2D；不要调用 sprite.image（会 crop）。
    alpha_ptr = getattr(rd, "alphaTexture", None)
    if _pptr_valid(alpha_ptr):
        try:
            alpha = alpha_ptr.deref().parse_as_object().image
            if alpha.size == image.size:
                rgba = image.convert("RGBA")
                rgba.putalpha(alpha.split()[0])
                image = rgba
        except Exception:
            pass
    return image


def _detect_payload_extension(data: bytes, *, audio=False, video=False) -> str:
    head = memoryview(data)[:16]
    if bytes(head[:4]) == b"OggS":
        return ".ogg" if audio else ".ogv"
    if bytes(head[:4]) == b"RIFF":
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
    # 先取内嵌数据。
    for attr in ("m_AudioData", "m_VideoData", "m_Data"):
        value = getattr(data, attr, None)
        if value:
            try:
                raw = bytes(value)
                if raw:
                    return raw
            except Exception:
                pass

    # 再处理 .resource/.resS 等外部数据，避免 AudioClip/VideoClip 漏导。
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


def _safe_peek_name(obj, fallback: str) -> str:
    try:
        name = obj.peek_name()
        if name:
            return safe_piece(str(name), fallback)
    except Exception:
        pass
    return safe_piece(fallback, "unnamed")


def _synthetic_asset_path(obj, asset_root: str) -> str:
    typ = getattr(getattr(obj, "type", None), "name", "Unknown")
    path_id = int(getattr(obj, "path_id", 0) or 0)
    name = _safe_peek_name(obj, f"{typ}_{path_id}")
    ext = {
        "Texture2D": ".png", "Sprite": ".png", "TextAsset": ".txt",
        "AudioClip": ".audio", "VideoClip": ".video", "MovieTexture": ".video",
        "Font": ".font", "Mesh": ".mesh", "Shader": ".shader",
        "Texture2DArray": ".png", "Material": ".mat", "AnimationClip": ".anim",
        "AnimatorController": ".controller", "RuntimeAnimatorController": ".controller",
        "Avatar": ".avatar", "MonoBehaviour": ".asset", "SpriteAtlas": ".spriteatlas",
        "Cubemap": ".cubemap",
    }.get(typ, ".asset")
    if not posixpath.splitext(name)[1]:
        name += ext
    root = str(asset_root or "").replace("\\", "/").rstrip("/")
    return (root + "/" + name) if root else f"Assets/__uncontainered__/{typ}/{name}"


UNCONTAINERED_CONTENT_TYPES = {
    "Texture2D", "Sprite", "TextAsset", "AudioClip", "VideoClip", "MovieTexture",
    "Font", "Mesh", "Shader", "Texture2DArray", "Material", "AnimationClip",
    "AnimatorController", "RuntimeAnimatorController", "Avatar", "MonoBehaviour",
    "SpriteAtlas", "Cubemap", "ComputeShader", "AudioMixer", "AudioMixerGroup",
    "AudioMixerSnapshot", "RenderTexture", "TerrainData", "AssetBundleManifest",
}


def _quality_token(version: str, root: str) -> str:
    text = compact_text(f"{version} {root}")
    if "highquality" in text:
        return "HighQuality"
    if "lowquality" in text:
        return "LowQuality"
    return ""


def build_variant_subdirs(direct_rows: Sequence[dict]) -> dict[str, str]:
    """为 ID 布局预计算版本目录，避免 High/Low 冲突被简单命名成 __02。"""
    grouped = defaultdict(list)
    for row in direct_rows:
        grouped[str(row.get("id", "") or "").strip()].append(row)
    out = {}
    for _rid, items in grouped.items():
        versions = {str(x.get("version", "") or "").strip() for x in items if str(x.get("version", "") or "").strip()}
        if len(versions) <= 1:
            for x in items:
                out[str(x.get("bundle", "")).casefold()] = ""
            continue
        quality_to_versions = defaultdict(set)
        for x in items:
            ver = str(x.get("version", "") or "").strip()
            q = _quality_token(ver, str(x.get("asset_root", "") or ""))
            quality_to_versions[q or "(other)"].add(ver)
        simple_quality = all(q in {"HighQuality", "LowQuality"} for q in quality_to_versions) and all(len(vs) == 1 for vs in quality_to_versions.values())
        for x in items:
            ver = str(x.get("version", "") or "").strip()
            q = _quality_token(ver, str(x.get("asset_root", "") or ""))
            label = q if simple_quality and q else safe_piece(ver or q or "version", "version")
            out[str(x.get("bundle", "")).casefold()] = label
    return out


def _route_export_path(source_rel: PurePosixPath, layout: str, asset_id: str, variant_subdir: str) -> PurePosixPath:
    if layout == "original":
        return source_rel
    rid = safe_piece(asset_id, "unknown-id")
    return PurePosixPath(rid, safe_piece(variant_subdir, "version"), source_rel.name) if variant_subdir else PurePosixPath(rid, source_rel.name)


def _claim_export_path(out_rel: PurePosixPath, claims_dir: Path, source_key: str, bundle_name: str, layout: str) -> PurePosixPath:
    if layout == "original":
        return out_rel
    claims_dir.mkdir(parents=True, exist_ok=True)
    candidate = out_rel
    attempt = 0
    while True:
        digest = hashlib.sha1(candidate.as_posix().casefold().encode("utf-8", "surrogatepass")).hexdigest()
        claim = claims_dir / f"{digest}.claim"
        try:
            fd = os.open(str(claim), os.O_CREAT | os.O_EXCL | os.O_WRONLY)
            try:
                os.write(fd, source_key.encode("utf-8", "surrogateescape"))
            finally:
                os.close(fd)
            return candidate
        except FileExistsError:
            try:
                owner = claim.read_text(encoding="utf-8", errors="replace")
            except Exception:
                owner = ""
            if owner == source_key:
                return candidate
        attempt += 1
        stem, ext = posixpath.splitext(out_rel.name)
        bundle_short = re.sub(r"[^0-9a-fA-F]", "", Path(bundle_name).stem)[:8] or hashlib.sha1(bundle_name.encode()).hexdigest()[:8]
        suffix = bundle_short if attempt == 1 else f"{bundle_short}_{attempt}"
        candidate = out_rel.with_name(f"{stem}__{suffix}{ext}")


def _prepare_export_dest(source_rel: PurePosixPath, export_root: Path, *, layout: str, asset_id: str, variant_subdir: str, claims_dir: Path, source_key: str, bundle_name: str):
    out_rel = _route_export_path(source_rel, layout, asset_id, variant_subdir)
    out_rel = _claim_export_path(out_rel, claims_dir, source_key, bundle_name, layout)
    dest = export_root.joinpath(*out_rel.parts)
    dest.parent.mkdir(parents=True, exist_ok=True)
    return dest, out_rel


def _record(bundle, asset_path, out_rel, typ, path_id, mode, status="ok"):
    return {"bundle": bundle, "asset_path": asset_path, "exported_path": out_rel.as_posix(), "type": typ, "path_id": path_id, "mode": mode, "status": status}


def export_one_object(obj, container_path: str, export_root: Path, bundle_name: str, *, layout: str, asset_id: str, variant_subdir: str, claims_dir: Path):
    rel = _safe_container_relative_path(container_path)
    path_id = getattr(obj, "path_id", getattr(obj, "m_PathID", ""))
    typ = type(obj).__name__
    source_key = f"{bundle_name}:{path_id}:{container_path}"
    try:
        obj = _resolve_object_reader(obj)
        typ = getattr(getattr(obj, "type", None), "name", str(getattr(obj, "type", "Unknown")))
        path_id = getattr(obj, "path_id", path_id)
        source_key = f"{bundle_name}:{path_id}:{container_path}"
        def prep(source_rel, extra=""):
            return _prepare_export_dest(source_rel, export_root, layout=layout, asset_id=asset_id, variant_subdir=variant_subdir, claims_dir=claims_dir, source_key=source_key + extra, bundle_name=bundle_name)
        if typ == "Texture2D":
            data = obj.parse_as_object()
            source_rel = _image_output_relative_path(rel)
            dest, out_rel = prep(source_rel)
            data.image.save(dest, format="PNG")
            return [_record(bundle_name, container_path, out_rel, typ, path_id, "png-full-texture")]
        if typ == "Sprite":
            sprite = obj.parse_as_object()
            source_rel = _image_output_relative_path(rel)
            dest, out_rel = prep(source_rel)
            _full_image_from_sprite(sprite).save(dest, format="PNG")
            return [_record(bundle_name, container_path, out_rel, typ, path_id, "png-full-sprite-texture-no-crop")]
        if typ == "TextAsset":
            data = obj.parse_as_object()
            source_rel = _textasset_output_relative_path(rel)
            dest, out_rel = prep(source_rel)
            dest.write_bytes(_textasset_bytes(data))
            mode = "textasset-atlas" if rel.name.lower().endswith(".atlas.txt") else "textasset-skel" if rel.name.lower().endswith(".skel.bytes") else "textasset-exact"
            return [_record(bundle_name, container_path, out_rel, typ, path_id, mode)]
        if typ == "AudioClip":
            clip = obj.parse_as_object()
            try:
                samples = dict(clip.samples)
            except Exception:
                samples = {}
            if samples:
                records = []
                base_stem, _ = posixpath.splitext(rel.name)
                multi = len(samples) > 1
                for n, (sample_name, sample_data) in enumerate(samples.items(), 1):
                    sample_name = str(sample_name).replace("\\", "/").rsplit("/", 1)[-1]
                    sample_stem, sample_ext = posixpath.splitext(sample_name)
                    if not sample_ext:
                        sample_ext = _detect_payload_extension(bytes(sample_data), audio=True)
                    src_rel = rel.with_name(f"{base_stem}__{safe_piece(sample_stem, f'sample{n}')}{sample_ext}" if multi else f"{base_stem}{sample_ext}")
                    dest, out_rel = prep(src_rel, f":audio:{n}")
                    dest.write_bytes(bytes(sample_data))
                    records.append(_record(bundle_name, container_path, out_rel, typ, path_id, "audio-sample"))
                return records
            raw = _resource_payload(clip)
            if raw:
                ext = _detect_payload_extension(raw, audio=True)
                src_rel = _replace_last_extension(rel, ext)
                dest, out_rel = prep(src_rel, ":audio-raw")
                dest.write_bytes(raw)
                return [_record(bundle_name, container_path, out_rel, typ, path_id, "audio-resource-raw")]
            raise ValueError("AudioClip 无 samples/resource 数据")
        if typ in {"VideoClip", "MovieTexture"}:
            data = obj.parse_as_object()
            raw = _resource_payload(data)
            if not raw:
                raise ValueError("Video 无内嵌数据")
            src_rel = _replace_last_extension(rel, _detect_payload_extension(raw, video=True))
            dest, out_rel = prep(src_rel, ":video")
            dest.write_bytes(raw)
            return [_record(bundle_name, container_path, out_rel, typ, path_id, "video-resource")]
        if typ == "Font":
            data = obj.parse_as_object()
            raw = bytes(getattr(data, "m_FontData", b"") or b"")
            if raw:
                src_rel = _replace_last_extension(rel, ".otf" if raw[:4] == b"OTTO" else ".ttf")
                dest, out_rel = prep(src_rel)
                dest.write_bytes(raw)
                return [_record(bundle_name, container_path, out_rel, typ, path_id, "font")]
        if typ == "Mesh":
            data = obj.parse_as_object()
            src_rel = _replace_last_extension(rel, ".obj")
            dest, out_rel = prep(src_rel)
            dest.write_text(data.export(), encoding="utf-8")
            return [_record(bundle_name, container_path, out_rel, typ, path_id, "mesh-obj")]
        if typ == "Shader":
            data = obj.parse_as_object()
            src_rel = _replace_last_extension(rel, ".shader")
            dest, out_rel = prep(src_rel)
            dest.write_text(data.export() if hasattr(data, "export") else str(data), encoding="utf-8")
            return [_record(bundle_name, container_path, out_rel, typ, path_id, "shader-text")]
        if typ == "Texture2DArray":
            data = obj.parse_as_object()
            images = list(getattr(data, "images", []) or [])
            records = []
            base_stem, _ = posixpath.splitext(rel.name)
            for idx, image in enumerate(images):
                src_rel = rel.with_name(f"{base_stem}_{idx}.png")
                dest, out_rel = prep(src_rel, f":array:{idx}")
                image.save(dest, format="PNG")
                records.append(_record(bundle_name, container_path, out_rel, typ, path_id, "png-texture-array"))
            if records:
                return records
        dest, out_rel = prep(rel)
        raw = obj.get_raw_data()
        if isinstance(raw, memoryview):
            raw = raw.tobytes()
        elif not isinstance(raw, (bytes, bytearray)):
            raw = bytes(raw)
        dest.write_bytes(bytes(raw))
        return [_record(bundle_name, container_path, out_rel, typ, path_id, "raw-exact")]
    except Exception as e:
        fail_rel = _route_export_path(rel, layout, asset_id, variant_subdir)
        return [_record(bundle_name, container_path, fail_rel, typ, path_id, "failed", f"failed: {type(e).__name__}: {e}")]


def _write_status_json(path: Path, data: dict):
    payload = json.dumps(data, ensure_ascii=False)
    path.parent.mkdir(parents=True, exist_ok=True)
    last = None
    for attempt in range(8):
        try:
            with path.open("w", encoding="utf-8") as f:
                f.write(payload)
                f.flush()
            return
        except (PermissionError, OSError) as e:
            last = e
            time.sleep(0.01 * (attempt + 1))
    if last:
        raise last


def _read_json_tolerant(path: Path, default=None, retries=3):
    for i in range(max(1, retries)):
        try:
            return json.loads(path.read_text(encoding="utf-8"))
        except (FileNotFoundError, PermissionError, OSError, json.JSONDecodeError):
            if i + 1 < retries:
                time.sleep(0.008)
    return default


def _atomic_write_json(path: Path, data: dict):
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(data, ensure_ascii=False)
    tmp = path.with_name(f".{path.name}.{os.getpid()}.{time.time_ns()}.tmp")
    tmp.write_text(payload, encoding="utf-8")
    last = None
    for attempt in range(20):
        try:
            os.replace(tmp, path)
            return
        except (PermissionError, OSError) as e:
            last = e
            time.sleep(min(0.01 * (attempt + 1), 0.20))
    try:
        path.write_text(payload, encoding="utf-8")
        tmp.unlink(missing_ok=True)
    except Exception:
        tmp.unlink(missing_ok=True)
        if last:
            raise last


def export_worker_main(job_path: Path) -> int:
    job = json.loads(job_path.read_text(encoding="utf-8"))
    primary = Path(job["primary_path"])
    dependencies = [Path(x) for x in job.get("dependency_paths", []) if Path(x).is_file()]
    export_root = Path(job["export_root"])
    asset_root = str(job.get("asset_root", "") or "")
    asset_id = str(job.get("asset_id", "") or "")
    variant_subdir = str(job.get("variant_subdir", "") or "")
    layout = str(job.get("layout", "original") or "original")
    claims_dir = Path(job["claims_dir"])
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
        contained_keys = {_object_reader_key(o) for _, o in primary_items}
        extra_items = []
        for obj in primary_objects:
            if _object_reader_key(obj) in contained_keys:
                continue
            typ = getattr(getattr(obj, "type", None), "name", "Unknown")
            if typ in UNCONTAINERED_CONTENT_TYPES:
                extra_items.append((_synthetic_asset_path(obj, asset_root), obj))
        items = primary_items + extra_items
        total, ok, failed, records = len(items), 0, 0, []
        _write_status_json(status_path, {"current": 0, "total": total, "ok": 0, "failed": 0, "path": ""})
        for i, (asset_path, obj) in enumerate(items, 1):
            if i == 1 or i % 4 == 0:
                _write_status_json(status_path, {"current": i - 1, "total": total, "ok": ok, "failed": failed, "path": str(asset_path)})
            recs = export_one_object(obj, asset_path, export_root, bundle_name, layout=layout, asset_id=asset_id, variant_subdir=variant_subdir, claims_dir=claims_dir)
            records.extend(recs)
            if recs and all(r.get("status") == "ok" for r in recs):
                ok += 1
            else:
                failed += 1
        _atomic_write_json(result_path, {"status": "ok", "ok": ok, "failed": failed, "records": records})
        _write_status_json(status_path, {"current": total, "total": total, "ok": ok, "failed": failed, "path": "", "done": True})
        return 0
    except BaseException as e:
        try:
            _atomic_write_json(result_path, {"status": "worker-error", "ok": 0, "failed": 1, "records": [], "error": f"{type(e).__name__}: {e}"})
        except Exception:
            pass
        return 70


def _collect_recursive_dependency_names(row: dict, all_rows: Sequence[dict]):
    row_by_bundle = {str(r.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1].casefold(): r for r in all_rows if str(r.get("bundle", "")).strip()}
    out, seen, queue = [], set(), list(_parse_dependency_bundles(row))
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

# ---------------------------- Parallel Export UI ----------------------------

_export_thread_slot_lock = threading.Lock()
_export_thread_slot_map = {}
_export_next_slot = 1


def _export_slot() -> int:
    global _export_next_slot
    ident = threading.get_ident()
    with _export_thread_slot_lock:
        if ident not in _export_thread_slot_map:
            _export_thread_slot_map[ident] = _export_next_slot
            _export_next_slot += 1
        return _export_thread_slot_map[ident]


class ExportLiveUI:
    def __init__(self, total_jobs: int, workers: int):
        self.total = total_jobs
        self.workers = workers
        self.lock = threading.Lock()
        self.slots = {}
        self.status_paths = {}
        self.completed = 0
        self.failed_jobs = 0
        self.stop_event = threading.Event()
        self.renderer = threading.Thread(target=self._loop, daemon=True)

    def register(self, job_idx: int, slot: int, status_path: Path, bundle: str):
        with self.lock:
            self.slots[slot] = {"job_idx": job_idx, "bundle": bundle, "status": "active"}
            self.status_paths[job_idx] = status_path

    def finish(self, slot: int, success: bool):
        with self.lock:
            if slot in self.slots:
                self.slots[slot]["status"] = "done" if success else "failed"
            self.completed += 1
            if not success:
                self.failed_jobs += 1

    def start(self):
        self.renderer.start()

    def stop(self):
        self.stop_event.set()
        self.renderer.join(timeout=2)
        self.render()
        print()

    def _loop(self):
        while not self.stop_event.is_set():
            self.render()
            self.stop_event.wait(0.12)

    def render(self):
        with self.lock:
            slots = {k: dict(v) for k, v in self.slots.items()}
            status_paths = dict(self.status_paths)
            completed = self.completed
            failed_jobs = self.failed_jobs
        clear_screen()
        ui_title("Export", f"主 bundle {self.total:,} | 并发 worker {self.workers} | 每个 bundle 独立子进程")
        print(f"完成 {completed:,}/{self.total:,} | worker 失败 {failed_jobs:,}")
        print()
        for slot in range(1, self.workers + 1):
            info = slots.get(slot)
            if not info:
                print(f"E{slot:02d} 空闲")
                continue
            st = _read_json_tolerant(status_paths.get(info["job_idx"], Path("__missing__")), default={}, retries=1) or {}
            cur, tot = int(st.get("current", 0) or 0), int(st.get("total", 0) or 0)
            ok, fail = int(st.get("ok", 0) or 0), int(st.get("failed", 0) or 0)
            name = PurePosixPath(str(st.get("path", "")).replace("\\", "/")).name or info["bundle"]
            pct = cur * 100 / tot if tot else 0.0
            bar_w = 20
            filled = int(bar_w * min(100, pct) / 100)
            bar = "█" * filled + "░" * (bar_w - filled)
            print(f"E{slot:02d} [{bar}] {pct:6.2f}% {cur:>5}/{tot:<5} OK {ok:<5} ERR {fail:<4} {truncate(name, 48)}")


def run_export_subprocess(info: dict, ui: ExportLiveUI):
    slot = _export_slot()
    ui.register(info["job_idx"], slot, Path(info["status_path"]), info["bundle"])
    log_path = Path(info["log_path"])
    with log_path.open("wb") as logf:
        proc = subprocess.Popen(
            [sys.executable, "-u", str(Path(__file__).resolve()), "--_export-worker", str(info["job_path"])],
            stdout=logf, stderr=subprocess.STDOUT, cwd=str(Path(__file__).resolve().parent),
        )
        code = proc.wait()
    ui.finish(slot, code == 0)
    return code, info


def verify_downloaded_dependencies(rows: Sequence[dict], records: Sequence[dict]):
    available = {str(r.get("original_bundle", "")).casefold() for r in records if r.get("status") in {"downloaded", "exists"}}
    expected, uncertain = set(), 0
    for row in rows:
        deps = _parse_dependency_bundles(row)
        expected.update(x.casefold() for x in deps)
        try:
            count = int(str(row.get("max_dependency_bundle_count", "") or "0"))
        except ValueError:
            count = 0
        if count > 1 and not deps:
            uncertain += 1
    missing = sorted(x for x in expected if x not in available)
    return not missing, missing, uncertain


def export_parallel(direct_rows: Sequence[dict], all_rows: Sequence[dict], download_records: Sequence[dict], output_dir: Path, export_root: Path, layout: str, workers: int, session: SessionLog, round_no: int):
    local_by_bundle = {
        str(r.get("original_bundle", "")).replace("\\", "/").rsplit("/", 1)[-1].casefold(): output_dir / str(r.get("local_file", ""))
        for r in download_records if r.get("status") in {"downloaded", "exists"}
    }
    variant_subdirs = build_variant_subdirs(direct_rows)
    jobs = []
    for row in direct_rows:
        bundle = str(row.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
        primary = local_by_bundle.get(bundle.casefold())
        if not primary or not primary.is_file():
            continue
        dep_paths, seen = [], set()
        for dep in _collect_recursive_dependency_names(row, all_rows):
            p = local_by_bundle.get(dep.casefold())
            if p and p.is_file() and str(p).casefold() not in seen:
                seen.add(str(p).casefold())
                dep_paths.append(p)
        jobs.append((row, bundle, primary, dep_paths))
    if not jobs:
        raise RuntimeError("没有可供 Export 的主 bundle")
    export_root.mkdir(parents=True, exist_ok=True)
    work = session.export_work(round_no)
    _rmtree_retry(work)
    work.mkdir(parents=True, exist_ok=True)
    claims_dir = work / "claims"
    claims_dir.mkdir(parents=True, exist_ok=True)
    manifest = session.export_manifest(round_no)
    fail_log = session.export_failures(round_no)
    fail_log.unlink(missing_ok=True)
    infos = []
    for idx, (row, bundle, primary, deps) in enumerate(jobs, 1):
        token = f"job_{idx:05d}"
        job_path = work / f"{token}.json"
        status_path = work / f"{token}.status.json"
        result_path = work / f"{token}.result.json"
        log_path = work / f"{token}.worker.log"
        spec = {
            "primary_path": str(primary), "dependency_paths": [str(x) for x in deps],
            "export_root": str(export_root), "asset_root": str(row.get("asset_root", "") or ""),
            "asset_id": str(row.get("id", "") or ""), "variant_subdir": variant_subdirs.get(bundle.casefold(), ""),
            "layout": layout, "claims_dir": str(claims_dir), "bundle": bundle,
            "status_path": str(status_path), "result_path": str(result_path),
        }
        job_path.write_text(json.dumps(spec, ensure_ascii=False), encoding="utf-8")
        infos.append({"job_idx": idx, "row": row, "bundle": bundle, "job_path": str(job_path), "status_path": str(status_path), "result_path": str(result_path), "log_path": str(log_path)})
    max_workers = max(1, min(workers, len(infos)))
    global _export_thread_slot_map, _export_next_slot
    _export_thread_slot_map, _export_next_slot = {}, 1
    ui = ExportLiveUI(len(infos), max_workers)
    ui.start()
    total_ok = total_failed = 0
    with manifest.open("w", encoding="utf-8-sig", newline="") as mf:
        writer = csv.DictWriter(mf, fieldnames=EXPORT_MANIFEST_FIELDS)
        writer.writeheader()
        with ThreadPoolExecutor(max_workers=max_workers, thread_name_prefix="EXP") as pool:
            futures = [pool.submit(run_export_subprocess, info, ui) for info in infos]
            for fut in as_completed(futures):
                code, info = fut.result()
                result = _read_json_tolerant(Path(info["result_path"]), default=None, retries=6)
                if code == 0 and result and result.get("status") == "ok":
                    writer.writerows(result.get("records", []))
                    mf.flush()
                    total_ok += int(result.get("ok", 0) or 0)
                    total_failed += int(result.get("failed", 0) or 0)
                    _unlink_retry(Path(info["log_path"]))
                else:
                    total_failed += 1
                    err = str(result.get("error", "")) if result else ""
                    try:
                        worker_log = Path(info["log_path"]).read_text(encoding="utf-8", errors="replace")[-10000:]
                    except Exception:
                        worker_log = ""
                    with fail_log.open("a", encoding="utf-8") as f:
                        f.write(f"\n===== {info['bundle']} | exit={code} =====\n{err}\n{worker_log}\n")
                    writer.writerow({
                        "bundle": info["bundle"], "asset_path": str(info["row"].get("asset_root", "") or ""),
                        "exported_path": "", "type": "BundleWorker", "path_id": "", "mode": "isolated-worker-parallel",
                        "status": f"worker-failed: exit={code}; see {fail_log.name}",
                    })
                    mf.flush()
    ui.stop()
    _rmtree_retry(work)
    try:
        if fail_log.exists() and fail_log.stat().st_size == 0:
            _unlink_retry(fail_log)
    except OSError:
        pass
    return total_ok, total_failed, manifest, fail_log


def choose_export_layout(value=None) -> str:
    if value in {"original", "id"}:
        return value
    clear_screen()
    ui_title("Export 存放方式")
    print()
    print("  1  原目录")
    print("     Assets/AssetBundles/... 保留真实路径")
    print()
    print("  2  按编号归档")
    print("     单版本：ch_100021/文件")
    print("     多版本：ch_100021/HighQuality/文件")
    print("             ch_100021/LowQuality/文件")
    print("     若同 Quality 还有 adult/general 等多个 version，会自动使用完整 version 目录。")
    print()
    while True:
        c = input("请选择 [1/2，默认 1]：").strip()
        if c in {"", "1"}:
            return "original"
        if c == "2":
            return "id"


# ---------------------------- Interactive round / main ----------------------------

def run_download_round(args, rows, source: str, work_dir: Path, output_dir: Path, export_root: Path, logs_root: Path, session: SessionLog):
    while True:
        clear_screen()
        ui_title("Bundle 筛选 / 下载", f"数据源 {source} | 记录 {len(rows):,} | 下载线程 {args.workers} | Export线程 {args.export_workers}")
        print_query_help()
        print()
        q = input("请输入筛选条件（R重新准备 / Q退出）：").strip()
        if q.casefold() in {"q", "quit", "exit"}:
            return "quit"
        if q.casefold() in {"r", "reload"}:
            return "reload"
        if not q:
            continue
        direct = fuzzy_filter_rows(rows, q)
        if not direct:
            print("\n没有匹配结果。")
            wait_enter()
            continue
        if args.no_bundle_deps:
            matches = [dict(x) for x in direct]
            for x in matches:
                x["_selected_reason"] = "query"
            unresolved = []
        else:
            matches, unresolved = expand_with_dependencies(direct, rows)
        names = build_local_names(matches)
        if args.force:
            existing_records, need_network = [], list(zip([dict(x) for x in matches], names))
        else:
            existing_records, need_network = preflight_existing(matches, names, output_dir, logs_root, args.base_url)
        print_preview(q, direct, matches, unresolved, len(existing_records))
        if args.dry_run:
            wait_enter()
            continue
        print()
        c = input("Enter=开始下载 | B=返回筛选 | Q=退出：").strip().casefold()
        if c in {"q", "quit"}:
            return "quit"
        if c in {"b", "back"}:
            continue
        if c:
            continue
        round_no = session.next_round()
        records = list(existing_records)
        if need_network:
            ui = DownloadLiveUI(len(need_network), len(existing_records), args.workers)
            ui.start()
            with ThreadPoolExecutor(max_workers=args.workers, thread_name_prefix="DL") as pool:
                futures = [
                    pool.submit(download_bundle_job, row, name, output_dir, args.base_url, ui, session, timeout=args.timeout, retries=args.retries, force=args.force)
                    for row, name in need_network
                ]
                for fut in as_completed(futures):
                    try:
                        records.append(fut.result())
                    except Exception as e:
                        session.log(f"download worker exception: {type(e).__name__}: {e}")
            ui.stop()
        records.sort(key=lambda r: str(r.get("local_file", "")).casefold())
        dl_manifest = session.download_manifest(round_no)
        write_download_manifest(dl_manifest, records)
        ok = sum(1 for r in records if r.get("status") == "downloaded")
        skipped = sum(1 for r in records if r.get("status") == "exists")
        failed = sum(1 for r in records if str(r.get("status", "")).startswith("failed"))
        cleaned = cleanup_temp(output_dir=output_dir, work_dir=work_dir)
        clear_screen()
        ui_title("下载完成", f"筛选：{q}")
        print(f"新下载     : {ok:,}")
        print(f"本地已存在 : {skipped:,}")
        print(f"失败       : {failed:,}")
        print(f"下载目录   : {output_dir}")
        print(f"下载清单   : {dl_manifest}")
        print(f"统一日志   : {session.root}")
        if cleaned:
            print(f"临时清理   : {len(cleaned):,} 项")
        do_export = args.export_assets
        if not args.no_export and not args.export_assets:
            print()
            do_export = ask_yes_no("是否 Export 本轮直接匹配资源？", default=False)
        if do_export:
            _complete, missing, uncertain = verify_downloaded_dependencies(matches, records)
            if missing:
                print(f"\n缺少 {len(missing)} 个依赖 bundle。")
                if not ask_yes_no("仍继续 Export？", default=False):
                    do_export = False
            if uncertain:
                print(f"注意：有 {uncertain} 条旧 CSV 记录无法完整确认依赖。")
        if do_export:
            if not ensure_export_dependencies(allow_install=not args.no_install_export_deps):
                print("Export 依赖未满足，跳过。")
            else:
                layout = choose_export_layout(args.export_layout)
                clear_screen()
                ui_title("Export", f"{'原 Assets 路径' if layout == 'original' else '按编号/版本归档'} | worker {args.export_workers}")
                try:
                    ex_ok, ex_failed, ex_manifest, ex_fail_log = export_parallel(direct, rows, records, output_dir, export_root, layout, args.export_workers, session, round_no)
                    cleaned2 = cleanup_temp(output_dir=output_dir, work_dir=work_dir)
                    clear_screen()
                    ui_title("Export 完成")
                    print(f"成功对象/项 : {ex_ok:,}")
                    print(f"失败对象/任务: {ex_failed:,}")
                    print(f"Export目录   : {export_root}")
                    print(f"Export清单   : {ex_manifest}")
                    if ex_fail_log.exists():
                        print(f"失败日志     : {ex_fail_log}")
                    print(f"统一日志目录 : {session.root}")
                    if cleaned2:
                        print(f"临时清理     : {len(cleaned2):,} 项")
                except Exception as e:
                    session.log(f"export exception: {type(e).__name__}: {e}")
                    print(f"\nExport 失败: {type(e).__name__}: {e}")
        print()
        print("Enter 继续下载    R 重新检查 CSV/catalog    Q 退出")
        c = input("请选择：").strip().casefold()
        if c in {"q", "quit"}:
            return "quit"
        if c in {"r", "reload"}:
            return "reload"
        continue


def build_parser():
    ap = argparse.ArgumentParser(description="Unity Addressables V7：20线程下载 + 20线程Export + AND/OR模糊路径筛选")
    ap.add_argument("queries", nargs="*", help="可选命令行筛选；双击 BAT 时使用交互界面")
    ap.add_argument("--csv")
    ap.add_argument("--catalog")
    ap.add_argument("--catalog-bundle-url", default=CATALOG_BUNDLE_URL)
    ap.add_argument("--base-url", default=DEFAULT_BUNDLE_BASE_URL)
    ap.add_argument("-o", "--output-dir", default="downloaded_bundles")
    ap.add_argument("--export-dir", default="exported_assets")
    ap.add_argument("--logs-dir", default="logs")
    ap.add_argument("--workers", type=int, default=DEFAULT_DOWNLOAD_WORKERS)
    ap.add_argument("--export-workers", type=int, default=DEFAULT_EXPORT_WORKERS)
    ap.add_argument("--timeout", type=int, default=60)
    ap.add_argument("--retries", type=int, default=3)
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--make-csv", action="store_true")
    ap.add_argument("--no-csv", action="store_true")
    ap.add_argument("--update-catalog", action="store_true")
    ap.add_argument("--no-update-prompt", action="store_true")
    ap.add_argument("--export", dest="export_assets", action="store_true")
    ap.add_argument("--no-export", action="store_true")
    ap.add_argument("--export-layout", choices=["original", "id"])
    ap.add_argument("--no-install-export-deps", action="store_true")
    ap.add_argument("--no-bundle-deps", action="store_true")
    return ap


def commandline_once(args, rows, source, work_dir, output_dir, export_root, logs_root, session):
    query = ",".join(args.queries)
    direct = fuzzy_filter_rows(rows, query)
    if not direct:
        print(f"没有匹配: {query}")
        return 1
    if args.no_bundle_deps:
        matches = [dict(x) for x in direct]
        for x in matches:
            x["_selected_reason"] = "query"
        unresolved = []
    else:
        matches, unresolved = expand_with_dependencies(direct, rows)
    names = build_local_names(matches)
    if args.force:
        existing, network = [], list(zip([dict(x) for x in matches], names))
    else:
        existing, network = preflight_existing(matches, names, output_dir, logs_root, args.base_url)
    print_preview(query, direct, matches, unresolved, len(existing))
    if args.dry_run:
        return 0
    if not ask_yes_no("开始下载？", default=True):
        return 0
    round_no = session.next_round()
    records = list(existing)
    if network:
        ui = DownloadLiveUI(len(network), len(existing), args.workers)
        ui.start()
        with ThreadPoolExecutor(max_workers=args.workers, thread_name_prefix="DL") as pool:
            futs = [pool.submit(download_bundle_job, row, name, output_dir, args.base_url, ui, session, timeout=args.timeout, retries=args.retries, force=args.force) for row, name in network]
            for f in as_completed(futs):
                records.append(f.result())
        ui.stop()
    records.sort(key=lambda r: str(r.get("local_file", "")).casefold())
    write_download_manifest(session.download_manifest(round_no), records)
    if args.export_assets and ensure_export_dependencies(not args.no_install_export_deps):
        export_parallel(direct, rows, records, output_dir, export_root, args.export_layout or "original", args.export_workers, session, round_no)
    cleanup_temp(output_dir=output_dir, work_dir=work_dir)
    return 0


def migrate_legacy_logs(output_dir: Path, export_root: Path, session: SessionLog) -> None:
    """把旧版散落在下载/Export目录里的日志移到本次统一日志目录。"""
    candidates = [
        output_dir / "download_manifest.csv",
        export_root / "export_manifest.csv",
        export_root / "export_failures.log",
    ]
    for src in candidates:
        if not src.is_file():
            continue
        dest = session.root / f"legacy_{src.name}"
        n = 1
        while dest.exists():
            n += 1
            dest = session.root / f"legacy_{n:02d}_{src.name}"
        try:
            shutil.move(str(src), str(dest))
            session.log(f"migrated legacy log: {src} -> {dest}")
        except Exception as e:
            session.log(f"legacy log migrate failed: {src}: {e}")
    old_work = export_root / ".export_work"
    if old_work.exists():
        _rmtree_retry(old_work)


def main(argv=None):
    args = build_parser().parse_args(argv)
    args.workers = max(1, int(args.workers))
    args.export_workers = max(1, int(args.export_workers))
    work_dir = Path.cwd().resolve()
    output_dir = Path(args.output_dir).expanduser()
    export_root = Path(args.export_dir).expanduser()
    logs_root = Path(args.logs_dir).expanduser()
    if not output_dir.is_absolute(): output_dir = work_dir / output_dir
    if not export_root.is_absolute(): export_root = work_dir / export_root
    if not logs_root.is_absolute(): logs_root = work_dir / logs_root
    output_dir, export_root, logs_root = output_dir.resolve(), export_root.resolve(), logs_root.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    logs_root.mkdir(parents=True, exist_ok=True)
    session = SessionLog(logs_root)
    session.log(f"start V{VERSION}, python={sys.version}")
    migrate_legacy_logs(output_dir, export_root, session)
    def exit_cleanup():
        try:
            cleanup_temp(output_dir=output_dir, work_dir=work_dir)
        except Exception:
            pass
    atexit.register(exit_cleanup)
    try:
        rows, source = prepare_mapping(args, work_dir, session)
    except Exception as e:
        session.log(f"prepare mapping failed: {type(e).__name__}: {e}")
        clear_screen()
        ui_title("准备失败")
        print(f"{type(e).__name__}: {e}")
        return 2
    if args.queries:
        return commandline_once(args, rows, source, work_dir, output_dir, export_root, logs_root, session)
    while True:
        action = run_download_round(args, rows, source, work_dir, output_dir, export_root, logs_root, session)
        if action == "quit":
            clear_screen()
            ui_title("已退出")
            print(f"日志：{session.root}")
            return 0
        if action == "reload":
            try:
                rows, source = prepare_mapping(args, work_dir, session)
            except Exception as e:
                session.log(f"reload failed: {type(e).__name__}: {e}")
                print(f"\n重新准备失败: {e}")
                wait_enter()


if __name__ == "__main__":
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
