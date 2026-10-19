# -*- coding: utf-8 -*-
"""Smart Bundle Downloader GUI - Tkinter front-end for the complete V7 engine."""
from __future__ import annotations

import csv
import json
import math
import os
import queue
import re
import subprocess
import sys
import threading
import time
import traceback
import urllib.error
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path, PurePosixPath
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

ROOT_DIR = Path(__file__).resolve().parent
CORE_FILE = ROOT_DIR / "core" / "smart_bundle_downloader_export_v7.py"
SPINE_ENGINE_FILE = ROOT_DIR / "core" / "spine_global_actions.py"
CONFIG_FILE = ROOT_DIR / "config.json"


def load_core():
    import importlib.util
    spec = importlib.util.spec_from_file_location("smart_bundle_downloader_core", CORE_FILE)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载下载核心：{CORE_FILE}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


core = load_core()


def load_spine_engine():
    import importlib.util
    spec = importlib.util.spec_from_file_location("smart_bundle_spine_global_actions", SPINE_ENGINE_FILE)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"无法加载 Spine 全局动作模块：{SPINE_ENGINE_FILE}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


spine_engine = load_spine_engine()

DEFAULT_CONFIG = {
    "base_url": core.DEFAULT_BUNDLE_BASE_URL,
    "catalog_bundle_url": core.CATALOG_BUNDLE_URL,
    "output_dir": "downloaded_bundles",
    "export_dir": "exported_assets",
    "logs_dir": "logs",
    "workers": core.DEFAULT_DOWNLOAD_WORKERS,
    "export_workers": core.DEFAULT_EXPORT_WORKERS,
    "timeout": 60,
    "retries": 3,
    "force": False,
    "expand_dependencies": True,
    "skip_low_quality": True,
    "export_layout": "character",
    "resource_file": "",
    "remember_ui_state": True,
    "filter_search": "",
    "filter_category": "全部分类",
    "filter_version": "全部版本",
    "sort_column": "id",
    "sort_descending": False,
    "checked_bundles": [],
    "window_geometry": "1420x860",
    # Spine global-action injection is integrated, but post-export automation is OFF by default.
    "spine_auto_enabled": False,
    "spine_source_dir": "",
    "spine_output_dir": "exported_assets/spine_global_actions",
    "spine_converter_path": "",
    "spine_root_bone": "",
    "spine_make_skel": False,
    "spine_remove_curve": False,
    "spine_keep_json": False,
    "spine_backup": True,
    "spine_actions": spine_engine.DEFAULT_ACTIONS,
    "config_version": 3,
}

ALIASES = {
    "bundle": {"bundle", "bundlename", "bundle_name", "assetbundle", "asset_bundle", "资源包", "bundle文件", "bundle名称", "文件名"},
    "id": {"id", "assetid", "asset_id", "resourceid", "resource_id", "编号", "资源id", "资源编号"},
    "version": {"version", "quality", "版本", "版本号", "品质"},
    "category": {"category", "type", "kind", "分类", "类型", "资源类型"},
    "asset_root": {"asset_root", "assetroot", "asset_path", "assetpath", "path", "root", "资源路径", "路径", "资源根路径"},
    "asset_count": {"asset_count", "assetcount", "count", "资源数量", "数量"},
    "confidence": {"confidence", "置信度"},
    "max_dependency_bundle_count": {"max_dependency_bundle_count", "dependency_count", "dependencycount", "依赖数量"},
    "dependency_bundles": {"dependency_bundles", "dependencies", "dependency", "依赖", "依赖包", "依赖bundles"},
    "bundle_internal_id": {"bundle_internal_id", "internal_id", "内部路径", "内部id"},
    "bundle_location_index": {"bundle_location_index", "location_index", "locationindex", "location索引"},
}


def _norm_header(value) -> str:
    return re.sub(r"[\s\-./\\()（）]+", "", str(value or "").strip().casefold())


NORMALIZED_ALIASES = {field: {_norm_header(x) for x in values | {field}} for field, values in ALIASES.items()}


def normalize_rows(raw_rows):
    rows = []
    for raw in raw_rows:
        if not isinstance(raw, dict):
            continue
        by_norm = {_norm_header(k): v for k, v in raw.items()}
        row = {field: "" for field in core.CSV_FIELDS}
        for field, aliases in NORMALIZED_ALIASES.items():
            for alias in aliases:
                if alias in by_norm and by_norm[alias] is not None:
                    row[field] = str(by_norm[alias]).strip()
                    break
        # Preserve supported engine columns that have already been named correctly.
        for field in core.CSV_FIELDS:
            if field in raw and raw[field] is not None:
                row[field] = str(raw[field]).strip()
        if row.get("bundle"):
            rows.append(row)
    if not rows:
        raise ValueError("资源表中没有可用记录。请确认表头中包含 bundle / Bundle名称 / 资源包等字段。")
    return rows


def _column_index(cell_ref: str) -> int:
    letters = "".join(ch for ch in cell_ref if ch.isalpha()).upper()
    value = 0
    for ch in letters:
        value = value * 26 + ord(ch) - ord("A") + 1
    return max(0, value - 1)


def read_xlsx(path: Path):
    """Read the first worksheet from XLSX with stdlib only (no openpyxl dependency)."""
    ns_main = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    ns_rel = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    ns_pkg = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall(f"{ns_main}si"):
                shared.append("".join(t.text or "" for t in si.iter(f"{ns_main}t")))
        sheet_path = "xl/worksheets/sheet1.xml"
        if "xl/workbook.xml" in names and "xl/_rels/workbook.xml.rels" in names:
            wb = ET.fromstring(zf.read("xl/workbook.xml"))
            sheets = wb.find(f"{ns_main}sheets")
            if sheets is not None and len(sheets):
                rel_id = sheets[0].attrib.get(ns_rel + "id")
                rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
                target = next((r.attrib.get("Target") for r in rels.findall(f"{ns_pkg}Relationship") if r.attrib.get("Id") == rel_id), None)
                if target:
                    target = target.replace("\\", "/")
                    sheet_path = target.lstrip("/") if target.startswith("/") else str(PurePosixPath("xl") / target)
                    sheet_path = str(PurePosixPath(sheet_path))
        if sheet_path not in names:
            candidates = sorted(n for n in names if n.startswith("xl/worksheets/") and n.endswith(".xml"))
            if not candidates:
                raise ValueError("XLSX 文件没有可读取的工作表。")
            sheet_path = candidates[0]
        sheet = ET.fromstring(zf.read(sheet_path))
        rows_data = []
        for row_el in sheet.iter(f"{ns_main}row"):
            values = {}
            max_col = -1
            for cell in row_el.findall(f"{ns_main}c"):
                idx = _column_index(cell.attrib.get("r", "A1"))
                max_col = max(max_col, idx)
                cell_type = cell.attrib.get("t", "")
                v_el = cell.find(f"{ns_main}v")
                if cell_type == "inlineStr":
                    val = "".join(t.text or "" for t in cell.iter(f"{ns_main}t"))
                elif v_el is None or v_el.text is None:
                    val = ""
                elif cell_type == "s":
                    try:
                        val = shared[int(v_el.text)]
                    except (ValueError, IndexError):
                        val = ""
                else:
                    val = v_el.text
                values[idx] = val
            rows_data.append([values.get(i, "") for i in range(max_col + 1)])
        if not rows_data:
            return []
        headers = [str(x).strip() for x in rows_data[0]]
        out = []
        for values in rows_data[1:]:
            if not any(str(v).strip() for v in values):
                continue
            out.append({headers[i] if i < len(headers) and headers[i] else f"column_{i+1}": values[i] if i < len(values) else "" for i in range(len(headers))})
        return out


def read_resource_file(path: Path):
    suffix = path.suffix.casefold()
    if suffix == ".csv" or suffix == ".txt":
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            sample = f.read(8192)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            raw = list(csv.DictReader(f, dialect=dialect))
    elif suffix == ".xlsx":
        raw = read_xlsx(path)
    elif suffix == ".json":
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if isinstance(data, dict):
            for key in ("rows", "data", "resources", "items"):
                if isinstance(data.get(key), list):
                    data = data[key]
                    break
            else:
                # Support bundle-keyed JSON maps.
                data = [dict(v, bundle=k) if isinstance(v, dict) else {"bundle": k, "id": "", "category": "", "version": ""} for k, v in data.items()]
        if not isinstance(data, list):
            raise ValueError("JSON 资源表应为数组、或包含 rows/data/resources/items 数组的对象。")
        raw = data
    else:
        raise ValueError("支持的资源表格式：CSV、XLSX、JSON、TXT（分隔文本）。")
    return normalize_rows(raw)


class GuiDownloadProgress:
    """Bridge used by the unchanged V7 transfer engine to report progress to Tk."""
    def __init__(self, app, workers: int):
        self.app = app
        self.workers = max(1, workers)
        self.states = {}
        self.lock = threading.Lock()

    @staticmethod
    def slot() -> int:
        m = re.search(r"_(\d+)$", threading.current_thread().name)
        return (int(m.group(1)) + 1) if m else 1

    def check_cancel(self):
        if self.app.cancel_event.is_set():
            raise core.DownloadCancelled("用户取消")
        while not self.app.resume_event.wait(0.2):
            if self.app.cancel_event.is_set():
                raise core.DownloadCancelled("用户取消")
        if self.app.cancel_event.is_set():
            raise core.DownloadCancelled("用户取消")

    def begin_file(self, slot: int, name: str, done=0, total=None):
        with self.lock:
            self.states[slot] = {"name": name, "done": int(done or 0), "total": total, "started": time.monotonic(), "start_bytes": int(done or 0)}
        self.app.post({"type": "progress", "name": name, "status": "下载中", "done": int(done or 0), "total": total, "speed": 0})

    def update(self, slot: int, done: int, total):
        with self.lock:
            state = self.states.setdefault(slot, {"name": "未知任务", "started": time.monotonic(), "start_bytes": 0})
            state["done"], state["total"] = int(done or 0), total
            elapsed = max(0.001, time.monotonic() - state.get("started", time.monotonic()))
            speed = max(0, int(done or 0) - int(state.get("start_bytes", 0))) / elapsed
            name = state.get("name", "未知任务")
        self.app.post({"type": "progress", "name": name, "status": "下载中", "done": int(done or 0), "total": total, "speed": speed})

    def finish(self, slot: int, success: bool):
        with self.lock:
            state = self.states.get(slot, {})
            name = state.get("name", "未知任务")
            done, total = state.get("done", 0), state.get("total")
        status = "完成" if success else ("已取消/失败" if self.app.cancel_event.is_set() else "失败")
        self.app.post({"type": "progress", "name": name, "status": status, "done": done, "total": total, "speed": 0})


ACTIVE_APP = None


class GuiExportProgress:
    """Non-terminal ExportLiveUI replacement; worker subprocess/export logic is untouched."""
    def __init__(self, total_jobs: int, workers: int):
        self.total = total_jobs
        self.workers = workers
        self.app = ACTIVE_APP
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
        if self.app:
            self.app.post({"type": "progress", "name": "EXPORT " + bundle, "status": "导出中", "done": 0, "total": None, "speed": 0})

    def finish(self, slot: int, success: bool):
        with self.lock:
            bundle = self.slots.get(slot, {}).get("bundle", "未知资源")
            if slot in self.slots:
                self.slots[slot]["status"] = "done" if success else "failed"
            self.completed += 1
            if not success:
                self.failed_jobs += 1
            completed = self.completed
        if self.app:
            self.app.post({"type": "progress", "name": "EXPORT " + bundle, "status": "完成" if success else "失败", "done": 1, "total": 1, "speed": 0})
            self.app.post({"type": "log", "text": f"Export worker 完成 {completed}/{self.total}: {bundle} ({'成功' if success else '失败'})"})

    def start(self):
        self.renderer.start()

    def stop(self):
        self.stop_event.set()
        self.renderer.join(timeout=2)

    def _loop(self):
        last_values = {}
        while not self.stop_event.wait(0.25):
            with self.lock:
                status_paths = dict(self.status_paths)
                slots = dict(self.slots)
            for job_idx, status_path in status_paths.items():
                if not status_path.exists():
                    continue
                try:
                    data = core._read_json_tolerant(status_path, default={}, retries=1) or {}
                    current, total = int(data.get("current", 0) or 0), int(data.get("total", 0) or 0)
                    name = PurePosixPath(str(data.get("path", "")).replace("\\", "/")).name or "Export"
                    marker = (current, total, name)
                    if marker != last_values.get(job_idx):
                        last_values[job_idx] = marker
                        if self.app:
                            self.app.post({"type": "progress", "name": "EXPORT " + slots.get(next((s for s, v in slots.items() if v.get("job_idx") == job_idx), -1), {}).get("bundle", name), "status": "导出中", "done": current, "total": total or None, "speed": 0})
                except Exception:
                    pass


core.ExportLiveUI = GuiExportProgress


class SmartBundleApp:
    def __init__(self, root):
        global ACTIVE_APP
        self.root = root
        self.root.title("Smart Bundle Downloader · Unity Addressables")
        self.root.geometry("1420x860")
        self.root.minsize(1050, 680)
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        ACTIVE_APP = self

        self.events = queue.Queue()
        self.busy = False
        self.current_task = ""
        self.cancel_event = threading.Event()
        self.resume_event = threading.Event()
        self.resume_event.set()
        self.rows = []
        self.filtered_indices = []
        self.checked_indices = set()
        self.item_to_index = {}
        self.download_tree_items = {}
        self.last_source = "尚未载入资源表"
        self.loaded_count = 0
        self._filter_job = None
        self._config = self.read_config()
        self.sort_column = str(self._config.get("sort_column", "id") or "id")
        self.sort_descending = bool(self._config.get("sort_descending", False))
        if bool(self._config.get("remember_ui_state", True)):
            try:
                self.root.geometry(str(self._config.get("window_geometry", "1420x860") or "1420x860"))
            except tk.TclError:
                pass
        self._build_style()
        self._build_ui()
        self._apply_config_to_vars()
        self._restore_log()
        self.root.after(100, self._drain_events)
        self.root.after(500, self.auto_load_source)

    def read_config(self):
        cfg = dict(DEFAULT_CONFIG)
        try:
            loaded = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                # Upgrade the previous GUI config: its default layout was
                # "original", while the new default is character-ID grouping.
                # Preserve all other user's settings during the migration.
                try:
                    old_version = int(loaded.get("config_version", 0) or 0)
                except (TypeError, ValueError):
                    old_version = 0
                if old_version < 2:
                    if loaded.get("export_layout", "original") == "original":
                        loaded["export_layout"] = "character"
                    loaded.setdefault("skip_low_quality", True)
                    loaded.setdefault("remember_ui_state", True)
                    loaded["config_version"] = 2
                loaded["config_version"] = max(3, old_version)
                cfg.update(loaded)
        except (OSError, json.JSONDecodeError):
            pass
        return cfg

    def save_config(self, silent=False):
        cfg = dict(DEFAULT_CONFIG)
        cfg.update(self._config)
        for key, var in self.vars.items():
            value = var.get()
            cfg[key] = value
        for key in ("workers", "export_workers", "timeout", "retries"):
            try:
                cfg[key] = max(1, int(cfg[key]))
            except (TypeError, ValueError):
                cfg[key] = DEFAULT_CONFIG[key]
        for key in ("force", "expand_dependencies", "skip_low_quality", "remember_ui_state", "spine_auto_enabled", "spine_make_skel", "spine_remove_curve", "spine_keep_json", "spine_backup"):
            cfg[key] = bool(cfg[key])
        remember_state = bool(cfg.get("remember_ui_state", True))
        if remember_state:
            if hasattr(self, "search_var"):
                cfg["filter_search"] = self.search_var.get()
                cfg["filter_category"] = self.category_var.get()
                cfg["filter_version"] = self.version_var.get()
            if hasattr(self, "sort_column"):
                cfg["sort_column"] = self.sort_column
                cfg["sort_descending"] = bool(self.sort_descending)
            if hasattr(self, "rows") and hasattr(self, "checked_indices"):
                checked = []
                for idx in sorted(self.checked_indices):
                    if 0 <= idx < len(self.rows):
                        key = self._bundle_key(self.rows[idx].get("bundle", ""))
                        if key:
                            checked.append(key)
                cfg["checked_bundles"] = sorted(set(checked))
            if hasattr(self, "root"):
                try:
                    cfg["window_geometry"] = self.root.geometry()
                except tk.TclError:
                    pass
        else:
            cfg.update({
                "filter_search": "", "filter_category": "全部分类", "filter_version": "全部版本",
                "sort_column": "id", "sort_descending": False, "checked_bundles": [],
                "window_geometry": "1420x860",
            })
        if hasattr(self, "spine_action_tree"):
            cfg["spine_actions"] = self._collect_spine_actions(silent_invalid=True)
        try:
            CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
            self._config = cfg
            if not silent:
                self.log("设置已保存。")
        except OSError as e:
            if not silent:
                messagebox.showerror("保存设置失败", str(e), parent=self.root)

    def _build_style(self):
        """Modern light theme using ttk's native widgets (no third-party UI dependency)."""
        style = ttk.Style(self.root)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        bg = "#F3F6FB"
        surface = "#FFFFFF"
        text = "#172033"
        muted = "#64748B"
        border = "#DCE3EF"
        primary = "#4F46E5"
        primary_hover = "#4338CA"
        soft = "#E9EDFF"
        danger = "#B42318"

        self.ui_colors = {
            "bg": bg, "surface": surface, "text": text, "muted": muted,
            "border": border, "primary": primary, "primary_hover": primary_hover,
            "soft": soft, "danger": danger,
        }
        self.root.configure(background=bg)
        style.configure(".", font=("Segoe UI", 10), background=bg, foreground=text)
        style.configure("TFrame", background=bg)
        style.configure("Card.TFrame", background=surface, relief="flat")
        style.configure("Header.TFrame", background=surface)
        style.configure("TLabel", background=bg, foreground=text)
        style.configure("Card.TLabel", background=surface, foreground=text)
        style.configure("Muted.TLabel", foreground=muted)
        style.configure("Title.TLabel", font=("Segoe UI Semibold", 20), background=surface, foreground="#111827")
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10), background=surface, foreground=muted)
        style.configure("Status.TLabel", padding=(10, 6), background="#E8EDF5", foreground="#334155")

        style.configure("TButton", padding=(11, 7), relief="flat", borderwidth=0,
                        background="#E8EDF5", foreground="#263248", focusthickness=0)
        style.map("TButton",
                  background=[("disabled", "#E9EDF3"), ("pressed", "#D4DDEC"), ("active", "#DDE5F1")],
                  foreground=[("disabled", "#94A3B8")])
        style.configure("Accent.TButton", padding=(13, 8), background=primary,
                        foreground="#FFFFFF", font=("Segoe UI Semibold", 10), borderwidth=0)
        style.map("Accent.TButton", background=[("disabled", "#A5B4FC"), ("pressed", primary_hover), ("active", primary_hover)],
                  foreground=[("disabled", "#F8FAFC"), ("!disabled", "#FFFFFF")])
        style.configure("Quiet.TButton", padding=(10, 7), background=surface,
                        foreground="#334155", borderwidth=1, relief="solid")
        style.map("Quiet.TButton", background=[("active", "#F8FAFC"), ("pressed", "#E2E8F0")])
        style.configure("Danger.TButton", padding=(10, 7), background="#FEE4E2", foreground=danger)
        style.map("Danger.TButton", background=[("active", "#FECDCA"), ("pressed", "#FDA29B")])

        style.configure("TEntry", padding=(8, 7), fieldbackground=surface, foreground=text,
                        bordercolor=border, lightcolor=border, darkcolor=border, insertcolor=text)
        style.map("TEntry", bordercolor=[("focus", primary)])
        style.configure("TCombobox", padding=(7, 6), fieldbackground=surface, background=surface,
                        foreground=text, bordercolor=border, arrowcolor=muted)
        style.map("TCombobox", fieldbackground=[("readonly", surface)], foreground=[("readonly", text)],
                  bordercolor=[("focus", primary)])
        style.configure("TCheckbutton", background=bg, foreground=text, padding=(2, 4))
        style.map("TCheckbutton", background=[("active", bg)])
        style.configure("TLabelframe", background=bg, bordercolor=border, relief="solid", borderwidth=1)
        style.configure("TLabelframe.Label", background=bg, foreground="#344054",
                        font=("Segoe UI Semibold", 10))
        style.configure("TNotebook", background=bg, borderwidth=0, tabmargins=(0, 2, 0, 0))
        style.configure("TNotebook.Tab", padding=(17, 10), background="#E8EDF5", foreground="#475467",
                        font=("Segoe UI Semibold", 10), borderwidth=0)
        style.map("TNotebook.Tab", background=[("selected", surface), ("active", "#DDE5F3")],
                  foreground=[("selected", primary), ("active", text)])
        style.configure("Treeview", rowheight=29, font=("Segoe UI", 9), background=surface,
                        fieldbackground=surface, foreground=text, bordercolor=border, relief="flat")
        style.configure("Treeview.Heading", font=("Segoe UI Semibold", 9), background="#E9EEF6",
                        foreground="#344054", padding=(8, 8), relief="flat")
        style.map("Treeview", background=[("selected", "#DDE6FF")], foreground=[("selected", "#1D2B53")])
        style.map("Treeview.Heading", background=[("active", "#DDE5F1")])
        style.configure("Horizontal.TProgressbar", background=primary, troughcolor="#DFE6F1", borderwidth=0,
                        lightcolor=primary, darkcolor=primary)
        style.configure("Vertical.TScrollbar", background="#D9E1ED", troughcolor=bg,
                        borderwidth=0, arrowsize=12)
        style.configure("Horizontal.TScrollbar", background="#D9E1ED", troughcolor=bg,
                        borderwidth=0, arrowsize=12)

    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=(16, 13, 16, 12))
        outer.pack(fill="both", expand=True)
        head = ttk.Frame(outer, style="Header.TFrame", padding=(16, 12))
        head.pack(fill="x", pady=(0, 12))
        title_area = ttk.Frame(head, style="Header.TFrame")
        title_area.pack(side="left", fill="x", expand=True)
        ttk.Label(title_area, text="Smart Bundle Downloader", style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_area, text="资源索引  ·  依赖展开  ·  多线程下载  ·  Unity 资源导出", style="Subtitle.TLabel").pack(anchor="w", pady=(3, 0))
        ttk.Label(head, text="RESOURCE WORKSPACE", font=("Segoe UI Semibold", 9),
                  background="#EEF2FF", foreground="#4F46E5", padding=(12, 8)).pack(side="right", padx=(12, 0))

        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True)
        self.resources_tab = ttk.Frame(self.notebook, padding=8)
        self.tasks_tab = ttk.Frame(self.notebook, padding=8)
        self.settings_tab = ttk.Frame(self.notebook, padding=8)
        self.spine_tab = ttk.Frame(self.notebook, padding=8)
        self.logs_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.resources_tab, text="资源库")
        self.notebook.add(self.tasks_tab, text="下载 / 导出任务")
        self.notebook.add(self.spine_tab, text="Spine 全局动作")
        self.notebook.add(self.settings_tab, text="设置")
        self.notebook.add(self.logs_tab, text="日志")
        self._build_resources_tab()
        self._build_tasks_tab()
        self._build_settings_tab()
        self._build_spine_tab()
        self._build_logs_tab()

        self.status_var = tk.StringVar(value="就绪")
        status = ttk.Frame(outer)
        status.pack(fill="x", pady=(8, 0))
        ttk.Label(status, textvariable=self.status_var, style="Status.TLabel", anchor="w").pack(side="left", fill="x", expand=True)
        self.global_progress = ttk.Progressbar(status, mode="indeterminate", length=180)
        self.global_progress.pack(side="right", padx=(8, 0))

    def _build_resources_tab(self):
        toolbar = ttk.Frame(self.resources_tab)
        toolbar.pack(fill="x", pady=(0, 8))
        self.import_btn = ttk.Button(toolbar, text="导入资源表", command=self.choose_resource_file)
        self.import_btn.pack(side="left")
        self.scan_btn = ttk.Button(toolbar, text="自动扫描本地资源", command=self.auto_load_source)
        self.scan_btn.pack(side="left", padx=(6, 0))
        self.catalog_btn = ttk.Button(toolbar, text="分析 Catalog 文件", command=self.choose_catalog_file)
        self.catalog_btn.pack(side="left", padx=(6, 0))
        self.update_catalog_btn = ttk.Button(toolbar, text="在线更新 Catalog", command=self.update_catalog)
        self.update_catalog_btn.pack(side="left", padx=(6, 0))
        self.source_var = tk.StringVar(value=self.last_source)
        ttk.Label(toolbar, textvariable=self.source_var, anchor="e").pack(side="right", fill="x", expand=True, padx=(12, 0))

        filters = ttk.Frame(self.resources_tab)
        filters.pack(fill="x", pady=(0, 7))
        ttk.Label(filters, text="快速搜索").pack(side="left")
        remember_view = bool(self._config.get("remember_ui_state", True))
        self.search_var = tk.StringVar(value=str(self._config.get("filter_search", "") if remember_view else ""))
        search_entry = ttk.Entry(filters, textvariable=self.search_var, width=34)
        search_entry.pack(side="left", padx=(6, 10))
        search_entry.bind("<Escape>", lambda _e: self.search_var.set(""))
        ttk.Label(filters, text="分类").pack(side="left")
        category_initial = str(self._config.get("filter_category", "全部分类") if remember_view else "全部分类")
        self.category_var = tk.StringVar(value=category_initial)
        self.category_combo = ttk.Combobox(filters, textvariable=self.category_var, state="readonly", width=15, values=["全部分类"])
        self.category_combo.pack(side="left", padx=(6, 10))
        self.category_combo.bind("<<ComboboxSelected>>", lambda _e: self.apply_filter())
        ttk.Label(filters, text="版本").pack(side="left")
        version_initial = str(self._config.get("filter_version", "全部版本") if remember_view else "全部版本")
        self.version_var = tk.StringVar(value=version_initial)
        self.version_combo = ttk.Combobox(filters, textvariable=self.version_var, state="readonly", width=17, values=["全部版本"])
        self.version_combo.pack(side="left", padx=(6, 10))
        self.version_combo.bind("<<ComboboxSelected>>", lambda _e: self.apply_filter())
        self.count_var = tk.StringVar(value="0 条资源")
        ttk.Label(filters, textvariable=self.count_var).pack(side="right")
        self.search_var.trace_add("write", self._schedule_filter)

        tree_frame = ttk.Frame(self.resources_tab)
        tree_frame.pack(fill="both", expand=True)
        columns = ("check", "id", "category", "version", "bundle", "asset_root", "asset_count", "dependency", "local")
        self.resource_tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")
        headings = {
            "check": "勾选", "id": "资源 ID", "category": "分类", "version": "版本", "bundle": "Bundle", "asset_root": "资源路径", "asset_count": "Asset数", "dependency": "依赖数", "local": "本地状态"
        }
        widths = {"check": 50, "id": 115, "category": 105, "version": 130, "bundle": 260, "asset_root": 390, "asset_count": 75, "dependency": 75, "local": 95}
        self.resource_heading_labels = dict(headings)
        for col in columns:
            self.resource_tree.heading(col, text=headings[col], command=lambda c=col: self.sort_resource_rows(c))
            self.resource_tree.column(col, width=widths[col], minwidth=45, anchor="w" if col not in {"check", "asset_count", "dependency"} else "center", stretch=col in {"bundle", "asset_root"})
        self._update_sort_heading()
        yscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.resource_tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.resource_tree.xview)
        self.resource_tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.resource_tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)
        self.resource_tree.bind("<Button-1>", self._resource_tree_click, add=True)
        self.resource_tree.bind("<Double-1>", self._resource_tree_double_click, add=True)

        actions = ttk.Frame(self.resources_tab)
        actions.pack(fill="x", pady=(8, 0))
        ttk.Button(actions, text="勾选当前行", command=self.check_selected).pack(side="left")
        ttk.Button(actions, text="取消勾选当前行", command=self.uncheck_selected).pack(side="left", padx=(5, 0))
        ttk.Button(actions, text="全选筛选结果", command=self.check_all_filtered).pack(side="left", padx=(5, 0))
        ttk.Button(actions, text="反选筛选结果", command=self.invert_filtered).pack(side="left", padx=(5, 0))
        ttk.Button(actions, text="清空勾选", command=self.clear_checks).pack(side="left", padx=(5, 0))
        ttk.Separator(actions, orient="vertical").pack(side="left", fill="y", padx=9)
        self.preflight_btn = ttk.Button(actions, text="预检查（不下载）", command=self.preflight_selected)
        self.preflight_btn.pack(side="left")
        self.download_btn = ttk.Button(actions, text="下载选中资源", style="Accent.TButton", command=lambda: self.download_selected(False))
        self.download_btn.pack(side="left")
        self.download_export_btn = ttk.Button(actions, text="下载并导出", style="Accent.TButton", command=lambda: self.download_selected(True))
        self.download_export_btn.pack(side="left", padx=(5, 0))
        self.export_btn = ttk.Button(actions, text="仅导出本地资源", command=self.export_selected)
        self.export_btn.pack(side="left", padx=(5, 0))

    def _build_tasks_tab(self):
        controls = ttk.Frame(self.tasks_tab)
        controls.pack(fill="x", pady=(0, 8))
        ttk.Button(controls, text="打开下载目录", command=lambda: self.open_directory("output_dir")).pack(side="left")
        ttk.Button(controls, text="打开导出目录", command=lambda: self.open_directory("export_dir")).pack(side="left", padx=(5, 0))
        ttk.Button(controls, text="打开日志目录", command=lambda: self.open_directory("logs_dir")).pack(side="left", padx=(5, 0))
        self.pause_btn = ttk.Button(controls, text="暂停下载", command=self.toggle_pause, state="disabled")
        self.pause_btn.pack(side="left", padx=(15, 0))
        self.cancel_btn = ttk.Button(controls, text="取消当前下载", command=self.cancel_current, state="disabled")
        self.cancel_btn.pack(side="left", padx=(5, 0))
        self.task_summary_var = tk.StringVar(value="尚无运行任务")
        ttk.Label(controls, textvariable=self.task_summary_var).pack(side="right")

        frame = ttk.Frame(self.tasks_tab)
        frame.pack(fill="both", expand=True)
        columns = ("name", "status", "progress", "size", "speed")
        self.task_tree = ttk.Treeview(frame, columns=columns, show="headings", height=18)
        for col, title, width in (("name", "Bundle / 资源", 570), ("status", "状态", 120), ("progress", "进度", 110), ("size", "已处理 / 总量", 190), ("speed", "速度", 130)):
            self.task_tree.heading(col, text=title)
            self.task_tree.column(col, width=width, minwidth=80, anchor="w" if col == "name" else "center", stretch=col == "name")
        sy = ttk.Scrollbar(frame, orient="vertical", command=self.task_tree.yview)
        sx = ttk.Scrollbar(frame, orient="horizontal", command=self.task_tree.xview)
        self.task_tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        self.task_tree.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        sx.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        ttk.Label(self.tasks_tab, text="提示：暂停会在当前下载分块边界生效；取消会保留有效的 .part 断点文件，后续可续传。Export 任务由隔离子进程执行。", wraplength=1260).pack(anchor="w", pady=(8, 0))

    def _build_settings_tab(self):
        self.vars = {}
        self.settings_widgets = {}
        form = ttk.Frame(self.settings_tab)
        form.pack(fill="both", expand=True, anchor="nw")
        form.columnconfigure(1, weight=1)
        row = 0

        def add_entry(key, label, browse=None, width=70):
            nonlocal row
            ttk.Label(form, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=7)
            var = tk.StringVar()
            self.vars[key] = var
            entry = ttk.Entry(form, textvariable=var, width=width)
            entry.grid(row=row, column=1, sticky="ew", pady=7)
            self.settings_widgets[key] = entry
            if browse:
                ttk.Button(form, text="浏览…", command=lambda k=key, d=browse: self.browse_path(k, d)).grid(row=row, column=2, padx=(8, 0), pady=7)
            row += 1

        add_entry("base_url", "Bundle 下载基础 URL")
        add_entry("catalog_bundle_url", "Catalog.bundle URL")
        add_entry("output_dir", "Bundle 下载目录", "dir")
        add_entry("export_dir", "资源导出目录", "dir")
        add_entry("logs_dir", "日志目录", "dir")

        numeric = ttk.Frame(form)
        numeric.grid(row=row, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        for key, label, default in (("workers", "下载线程", 20), ("export_workers", "导出线程", 20), ("timeout", "超时秒数", 60), ("retries", "失败重试", 3)):
            block = ttk.Frame(numeric)
            block.pack(side="left", padx=(0, 18))
            ttk.Label(block, text=label).pack(anchor="w")
            var = tk.StringVar(value=str(default))
            self.vars[key] = var
            ttk.Spinbox(block, from_=1, to=128 if "workers" in key else 3600, textvariable=var, width=8).pack(anchor="w", pady=(3, 0))
        row += 1

        flags = ttk.Frame(form)
        flags.grid(row=row, column=0, columnspan=3, sticky="w", pady=8)
        self.vars["force"] = tk.BooleanVar(value=False)
        ttk.Checkbutton(flags, text="强制重新下载（覆盖已存在 Bundle）", variable=self.vars["force"]).pack(anchor="w", pady=2)
        self.vars["expand_dependencies"] = tk.BooleanVar(value=True)
        ttk.Checkbutton(flags, text="下载时自动展开 Bundle 依赖（推荐）", variable=self.vars["expand_dependencies"]).pack(anchor="w", pady=2)
        self.vars["skip_low_quality"] = tk.BooleanVar(value=True)
        ttk.Checkbutton(flags, text="默认跳过 LowQuality 版本（未标记高/低质量的资源仍照常下载）", variable=self.vars["skip_low_quality"]).pack(anchor="w", pady=2)
        self.vars["remember_ui_state"] = tk.BooleanVar(value=True)
        ttk.Checkbutton(flags, text="记住筛选条件、排序、已勾选资源和窗口布局", variable=self.vars["remember_ui_state"]).pack(anchor="w", pady=2)
        row += 1

        ttk.Label(form, text="Export 目录布局").grid(row=row, column=0, sticky="w", padx=(0, 12), pady=7)
        self.vars["export_layout"] = tk.StringVar(value="character")
        layout = ttk.Combobox(form, textvariable=self.vars["export_layout"], state="readonly", values=["character", "original", "id"], width=18)
        layout.grid(row=row, column=1, sticky="w", pady=7)
        ttk.Label(form, text="character = 特征ID/服饰ID分层；original = 原 Assets 分类；id = 资源ID归档", wraplength=720).grid(row=row, column=1, sticky="w", padx=(160, 0))
        row += 1

        buttons = ttk.Frame(form)
        buttons.grid(row=row, column=0, columnspan=3, sticky="w", pady=(18, 0))
        ttk.Button(buttons, text="保存设置", command=self.save_config).pack(side="left")
        ttk.Button(buttons, text="修复历史导出分类 / 清理残留", command=self.repair_existing_export_classification).pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text="安装 / 检查导出依赖", command=self.install_export_dependencies).pack(side="left", padx=(6, 0))
        ttk.Button(buttons, text="恢复默认值", command=self.reset_settings).pack(side="left", padx=(6, 0))
        ttk.Label(form, text="导出功能使用原 V7 的 UnityPy / Pillow 处理流程；下载和资源索引读取不要求安装这两个包。", wraplength=1120).grid(row=row+1, column=0, columnspan=3, sticky="w", pady=(18, 0))

    def _build_spine_tab(self):
        """Embedded manual SD/CH Spine processor and optional post-export hook."""
        outer = ttk.Frame(self.spine_tab)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(0, weight=1)
        self.spine_files = []

        # Separate setup, file selection, and action editing into full-height workspaces.
        # This preserves all controls at 100% Windows display scaling on laptop-sized screens.
        self.spine_workspace = ttk.Notebook(outer)
        self.spine_workspace.grid(row=0, column=0, sticky="nsew", pady=(0, 7))
        self.spine_paths_page = ttk.Frame(self.spine_workspace, padding=8)
        self.spine_files_page = ttk.Frame(self.spine_workspace, padding=5)
        self.spine_actions_page = ttk.Frame(self.spine_workspace, padding=5)
        self.spine_workspace.add(self.spine_paths_page, text="路径与运行")
        self.spine_workspace.add(self.spine_files_page, text="目标文件")
        self.spine_workspace.add(self.spine_actions_page, text="动作模板与关键帧")

        paths = ttk.LabelFrame(self.spine_paths_page, text="输入、转换器和输出", padding=10)
        paths.pack(fill="x", pady=(0, 8))
        paths.columnconfigure(1, weight=1)

        def path_row(row, key, label, browse_type):
            ttk.Label(paths, text=label).grid(row=row, column=0, sticky="w", padx=(0, 10), pady=4)
            var = tk.StringVar(value=str(self._config.get(key, DEFAULT_CONFIG.get(key, ""))))
            self.vars[key] = var
            ttk.Entry(paths, textvariable=var).grid(row=row, column=1, sticky="ew", pady=4)
            command = (lambda k=key: self._browse_spine_path(k, "file" if browse_type == "file" else "dir"))
            ttk.Button(paths, text="浏览…", command=command).grid(row=row, column=2, padx=(7, 0), pady=4)

        path_row(0, "spine_source_dir", "扫描目录（留空使用导出目录）", "dir")
        path_row(1, "spine_converter_path", "SpineSkeletonDataConverter.exe", "file")
        path_row(2, "spine_output_dir", "动作输出目录（不会覆盖原始 SKEL）", "dir")
        ttk.Label(paths, text="Root 骨骼名（留空自动识别）").grid(row=3, column=0, sticky="w", padx=(0, 10), pady=4)
        self.vars["spine_root_bone"] = tk.StringVar(value=str(self._config.get("spine_root_bone", "")))
        ttk.Entry(paths, textvariable=self.vars["spine_root_bone"]).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(paths, text="只扫描/处理文件名或路径明确包含 sd_ / ch_ 的 .skel；不会处理其他 Spine。", style="Muted.TLabel").grid(row=4, column=0, columnspan=3, sticky="w", pady=(5, 1))

        switches = ttk.LabelFrame(self.spine_paths_page, text="运行模式", padding=10)
        switches.pack(fill="x", pady=(0, 8))
        self.vars["spine_auto_enabled"] = tk.BooleanVar(value=bool(self._config.get("spine_auto_enabled", False)))
        ttk.Checkbutton(switches, text="导出完成后自动处理 SD/CH Spine（默认关闭）", variable=self.vars["spine_auto_enabled"]).pack(side="left", padx=(0, 12))
        self.vars["spine_make_skel"] = tk.BooleanVar(value=bool(self._config.get("spine_make_skel", False)))
        ttk.Checkbutton(switches, text="同时生成修改后的 SKEL", variable=self.vars["spine_make_skel"]).pack(side="left", padx=(0, 10))
        self.vars["spine_remove_curve"] = tk.BooleanVar(value=bool(self._config.get("spine_remove_curve", False)))
        ttk.Checkbutton(switches, text="Converter 使用 --remove-curve", variable=self.vars["spine_remove_curve"]).pack(side="left", padx=(0, 10))
        self.vars["spine_keep_json"] = tk.BooleanVar(value=bool(self._config.get("spine_keep_json", False)))
        ttk.Checkbutton(switches, text="保留中间 JSON", variable=self.vars["spine_keep_json"]).pack(side="left", padx=(0, 10))
        self.vars["spine_backup"] = tk.BooleanVar(value=bool(self._config.get("spine_backup", True)))
        ttk.Checkbutton(switches, text="覆盖输出前备份 .bak", variable=self.vars["spine_backup"]).pack(side="left")

        candidates = ttk.LabelFrame(self.spine_files_page, text="SD / CH Spine 文件列表（可单独执行，不依赖下载任务）", padding=7)
        candidates.pack(fill="both", expand=True)
        candidates.rowconfigure(0, weight=1)
        candidates.columnconfigure(0, weight=1)
        self.spine_file_tree = ttk.Treeview(candidates, columns=("path", "type"), show="headings", selectmode="extended", height=4)
        self.spine_file_tree.heading("path", text="SKEL 文件")
        self.spine_file_tree.heading("type", text="类型")
        self.spine_file_tree.column("path", width=540, minwidth=200, stretch=True)
        self.spine_file_tree.column("type", width=70, anchor="center", stretch=False)
        fs_y = ttk.Scrollbar(candidates, orient="vertical", command=self.spine_file_tree.yview)
        fs_x = ttk.Scrollbar(candidates, orient="horizontal", command=self.spine_file_tree.xview)
        self.spine_file_tree.configure(yscrollcommand=fs_y.set, xscrollcommand=fs_x.set)
        self.spine_file_tree.grid(row=0, column=0, sticky="nsew")
        fs_y.grid(row=0, column=1, sticky="ns")
        fs_x.grid(row=1, column=0, sticky="ew")
        file_buttons = ttk.Frame(candidates)
        file_buttons.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(file_buttons, text="扫描目录", command=self.scan_spine_files).pack(side="left")
        ttk.Button(file_buttons, text="添加 .skel", command=self.add_spine_files).pack(side="left", padx=(5, 0))
        ttk.Button(file_buttons, text="添加文件夹", command=self.add_spine_folder).pack(side="left", padx=(5, 0))
        ttk.Button(file_buttons, text="移除选中", command=self.remove_spine_files).pack(side="left", padx=(5, 0))
        ttk.Button(file_buttons, text="清空列表", command=self.clear_spine_files).pack(side="left", padx=(5, 0))

        actions_box = ttk.LabelFrame(self.spine_actions_page, text="全局动作模板", padding=10)
        actions_box.pack(fill="both", expand=True)
        actions_box.columnconfigure(0, weight=1)
        actions_box.rowconfigure(0, weight=1)
        self.spine_action_tree = ttk.Treeview(
            actions_box,
            columns=("name", "kind", "dur", "amp", "count", "repeat", "direction", "segments", "keys"),
            displaycolumns=("name", "kind", "dur", "amp", "count", "repeat", "direction", "keys"),
            show="headings", selectmode="browse", height=6,
        )
        for col, title, width in (("name", "动作名", 105), ("kind", "类型", 82), ("dur", "时长", 54), ("amp", "幅度", 54), ("count", "次数", 44), ("repeat", "重复", 48), ("direction", "方向", 52), ("segments", "关键帧数据", 0), ("keys", "关键帧", 58)):
            self.spine_action_tree.heading(col, text=title)
            self.spine_action_tree.column(col, width=width, minwidth=0 if col == "segments" else 36, anchor="center", stretch=(col in {"name", "keys"}))
        self.spine_action_tree.grid(row=0, column=0, sticky="nsew")
        self.spine_action_tree.bind("<<TreeviewSelect>>", self._select_spine_action)
        editor = ttk.Frame(actions_box)
        editor.grid(row=1, column=0, sticky="ew", pady=(7, 2))
        for col in range(7):
            editor.columnconfigure(col, weight=1)
        self.spine_action_name = tk.StringVar(value="NewAction")
        self.spine_action_kind = tk.StringVar(value="跳跃")
        self.spine_action_dur = tk.StringVar(value="0.2")
        self.spine_action_amp = tk.StringVar(value="48")
        self.spine_action_count = tk.StringVar(value="1")
        self.spine_action_repeat = tk.StringVar(value="1")
        self.spine_action_direction = tk.StringVar(value="上")
        self.spine_action_segments = tk.StringVar(value="[]")
        self.spine_segment_status = tk.StringVar(value="自定义关键帧：0 个")
        # Tracks which type's template the form's current segment data belongs to.
        # Loading a saved action resets this tracker without regenerating its data.
        self._spine_form_kind_for_segments = self.spine_action_kind.get()
        edit_fields = (("动作名", self.spine_action_name, 11, "entry"), ("类型", self.spine_action_kind, 8, "kind"), ("时长秒", self.spine_action_dur, 6, "entry"), ("幅度", self.spine_action_amp, 5, "entry"), ("次数", self.spine_action_count, 5, "entry"), ("重复", self.spine_action_repeat, 5, "entry"), ("方向", self.spine_action_direction, 6, "direction"))
        for c, (label, var, width, field_type) in enumerate(edit_fields):
            block = ttk.Frame(editor)
            block.grid(row=0, column=c, sticky="ew", padx=(0, 3))
            ttk.Label(block, text=label).pack(anchor="w")
            if field_type == "kind":
                self.spine_action_kind_combo = ttk.Combobox(
                    block, textvariable=var, values=spine_engine.ACTION_TYPES,
                    state="readonly", width=width,
                )
                self.spine_action_kind_combo.pack(fill="x")
                self.spine_action_kind_combo.bind("<<ComboboxSelected>>", self._on_spine_action_kind_changed)
            elif field_type == "direction":
                ttk.Combobox(block, textvariable=var, values=spine_engine.DIRECTIONS, state="readonly", width=width).pack(fill="x")
            else:
                ttk.Entry(block, textvariable=var, width=width).pack(fill="x")
        custom_keys_row = ttk.Frame(actions_box)
        custom_keys_row.grid(row=2, column=0, sticky="ew", pady=(2, 0))
        ttk.Label(custom_keys_row, textvariable=self.spine_segment_status, style="Muted.TLabel").pack(side="left")
        ttk.Button(custom_keys_row, text="关键帧编辑器…", style="Quiet.TButton", command=self.edit_spine_keyframes).pack(side="left", padx=(8, 0))
        ttk.Button(custom_keys_row, text="套用预设", style="Accent.TButton", command=self._apply_spine_preset_to_form).pack(side="left", padx=(6, 0))
        ttk.Label(custom_keys_row, text="切换类型会载入对应关键帧模板；自定义类型保留当前关键帧。", style="Muted.TLabel").pack(side="left", padx=(9, 0))
        action_buttons = ttk.Frame(actions_box)
        action_buttons.grid(row=3, column=0, sticky="ew", pady=(5, 0))
        ttk.Button(action_buttons, text="新增动作", style="Accent.TButton", command=self.add_spine_action).pack(side="left")
        ttk.Button(action_buttons, text="更新选中", command=self.update_spine_action).pack(side="left", padx=(6, 0))
        ttk.Button(action_buttons, text="复制选中", command=self.duplicate_spine_action).pack(side="left", padx=(6, 0))
        ttk.Button(action_buttons, text="删除动作", style="Danger.TButton", command=self.remove_spine_action).pack(side="left", padx=(6, 0))
        ttk.Button(action_buttons, text="全部按类型套用", command=self._apply_spine_presets_to_all).pack(side="left", padx=(6, 0))
        ttk.Button(action_buttons, text="恢复默认动作", command=self.reset_spine_actions).pack(side="left", padx=(6, 0))
        ttk.Button(action_buttons, text="保存模板", style="Quiet.TButton", command=self.save_config).pack(side="right")

        bottom = ttk.Frame(outer)
        bottom.grid(row=1, column=0, sticky="ew", pady=(1, 0))
        ttk.Button(bottom, text="执行所选 SD/CH Spine", style="Accent.TButton", command=self.run_spine_actions_independent).pack(side="left")
        ttk.Label(bottom, text="自动处理默认关闭；可独立执行，不必先下载。", style="Muted.TLabel").pack(side="left", padx=(10, 0))

        action_list = self._config.get("spine_actions", spine_engine.DEFAULT_ACTIONS)
        if not isinstance(action_list, list):
            action_list = spine_engine.DEFAULT_ACTIONS
        for action in action_list:
            if isinstance(action, dict):
                self._insert_spine_action(action)

    def _browse_spine_path(self, key, mode):
        raw = str(self.vars[key].get() or "").strip()
        initial = Path(raw).expanduser() if raw else ROOT_DIR
        if not initial.is_absolute():
            initial = ROOT_DIR / initial
        if mode == "file":
            picked = filedialog.askopenfilename(parent=self.root, title="选择 SpineSkeletonDataConverter.exe", initialdir=str(initial.parent if initial.parent.exists() else ROOT_DIR), filetypes=[("执行文件", "*.exe"), ("全部文件", "*.*")])
        else:
            picked = filedialog.askdirectory(parent=self.root, title="选择 Spine 资源目录", initialdir=str(initial if initial.is_dir() else ROOT_DIR))
        if picked:
            self.vars[key].set(picked)
            self.save_config(silent=True)

    def _spine_resolved_path(self, value, default=""):
        raw = str(value or default or "").strip()
        if not raw:
            return None
        path = Path(raw).expanduser()
        if not path.is_absolute():
            path = ROOT_DIR / path
        return path.resolve()

    def _spine_source_root(self):
        source_var = self.vars.get("spine_source_dir")
        export_var = self.vars.get("export_dir")
        raw = str(source_var.get() if source_var is not None else "").strip()
        if not raw:
            raw = str(export_var.get() if export_var is not None else DEFAULT_CONFIG["export_dir"])
        return self._spine_resolved_path(raw)

    def _add_spine_candidate(self, path):
        item = Path(path).expanduser().resolve()
        if not spine_engine.is_sd_ch_spine(item):
            return False
        if item not in self.spine_files:
            self.spine_files.append(item)
            tag = "CH" if item.stem.casefold().startswith("ch") else "SD"
            self.spine_file_tree.insert("", "end", iid=str(len(self.spine_files) - 1), values=(str(item), tag))
        return True

    def _refresh_spine_file_tree(self):
        for item in self.spine_file_tree.get_children():
            self.spine_file_tree.delete(item)
        unique = []
        seen = set()
        for path in self.spine_files:
            key = str(Path(path).resolve()).casefold()
            if key not in seen and spine_engine.is_sd_ch_spine(path):
                unique.append(Path(path).resolve()); seen.add(key)
        self.spine_files = unique
        for idx, path in enumerate(self.spine_files):
            tag = "CH" if path.stem.casefold().startswith("ch") else "SD"
            self.spine_file_tree.insert("", "end", iid=str(idx), values=(str(path), tag))

    def scan_spine_files(self):
        try:
            source = self._spine_source_root()
            output = self._spine_resolved_path(self.vars["spine_output_dir"].get(), DEFAULT_CONFIG["spine_output_dir"])
            found = spine_engine.scan_target_skel_files(source, output)
            self.spine_files = list(found)
            self._refresh_spine_file_tree()
            self.log(f"Spine 扫描完成：{len(found)} 个 SD/CH .skel；扫描目录：{source}")
            if not found:
                messagebox.showinfo("扫描完成", "没有发现文件名或路径中带 sd_ / ch_ 标识的 .skel。其他 Spine 文件已按要求过滤。", parent=self.root)
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc), parent=self.root)

    def add_spine_files(self):
        names = filedialog.askopenfilenames(parent=self.root, title="添加 SD / CH Spine .skel", filetypes=[("Spine SKEL", "*.skel")])
        accepted = 0
        rejected = 0
        for name in names:
            if self._add_spine_candidate(name):
                accepted += 1
            else:
                rejected += 1
        self._refresh_spine_file_tree()
        if rejected:
            messagebox.showwarning("已过滤非目标资源", f"加入 {accepted} 个 SD/CH Spine 文件；过滤 {rejected} 个非 SD/CH .skel 文件。", parent=self.root)

    def add_spine_folder(self):
        folder = filedialog.askdirectory(parent=self.root, title="选择包含 SD / CH Spine 的目录")
        if not folder:
            return
        try:
            output = self._spine_resolved_path(self.vars["spine_output_dir"].get(), DEFAULT_CONFIG["spine_output_dir"])
            found = spine_engine.scan_target_skel_files(folder, output)
            known = {str(x).casefold() for x in self.spine_files}
            self.spine_files.extend(x for x in found if str(x).casefold() not in known)
            self._refresh_spine_file_tree()
            self.log(f"添加目录扫描结果：新增 {len(self.spine_files) - len(known)} 个目标文件。")
        except Exception as exc:
            messagebox.showerror("扫描失败", str(exc), parent=self.root)

    def remove_spine_files(self):
        selected = {Path(self.spine_file_tree.item(iid, "values")[0]).resolve() for iid in self.spine_file_tree.selection()}
        self.spine_files = [x for x in self.spine_files if x.resolve() not in selected]
        self._refresh_spine_file_tree()

    def clear_spine_files(self):
        self.spine_files = []
        self._refresh_spine_file_tree()

    def _normalise_spine_segments(self, action):
        """Return a validated, compact list of editable keyframes for an action row."""
        kind = str(action.get("kind", "跳跃"))
        try:
            duration = float(action.get("dur", 0.2))
        except (TypeError, ValueError):
            duration = 0.2
        try:
            amplitude = float(action.get("amp", 48))
        except (TypeError, ValueError):
            amplitude = 48.0
        try:
            count = max(1, int(action.get("count", 1)))
        except (TypeError, ValueError):
            count = 1
        direction = str(action.get("direction", "上"))
        segments = action.get("segments", [])
        if isinstance(segments, str):
            try:
                segments = json.loads(segments or "[]")
            except json.JSONDecodeError:
                segments = []
        if not isinstance(segments, list):
            segments = []
        if len(segments) < 2:
            if kind in spine_engine.FACT:
                # Like the original standalone tool, a preset is real editable keyframe data,
                # not just a label that gets regenerated later during export.
                segments = spine_engine.FACT[kind](duration, amplitude, count, direction)
            else:
                duration = duration if duration > 0 else 0.2
                segments = [{"time": 0, "x": 0, "y": 0}, {"time": duration, "x": 0, "y": 0}]
        normalized = []
        for segment in segments:
            if not isinstance(segment, dict):
                continue
            normalized.append({
                "time": float(segment.get("time", segment.get("Time", 0))),
                "x": float(segment.get("x", segment.get("X", 0))),
                "y": float(segment.get("y", segment.get("Y", 0))),
            })
        normalized.sort(key=lambda item: item["time"])
        return normalized

    def _spine_action_values(self, action):
        segments = self._normalise_spine_segments(action)
        vals = (
            str(action.get("name", "Action")), str(action.get("kind", "跳跃")),
            str(action.get("dur", 0.2)), str(action.get("amp", 48)),
            str(action.get("count", 1)), str(action.get("repeat", 1)),
            str(action.get("direction", "上")),
            json.dumps(segments, ensure_ascii=False, separators=(",", ":")),
            str(len(segments)),
        )
        return vals

    def _insert_spine_action(self, action, iid=None):
        vals = self._spine_action_values(action)
        if iid and self.spine_action_tree.exists(iid):
            self.spine_action_tree.item(iid, values=vals)
        else:
            self.spine_action_tree.insert("", "end", values=vals)

    def _select_spine_action(self, _event=None):
        selected = self.spine_action_tree.selection()
        if not selected:
            return
        values = self.spine_action_tree.item(selected[0], "values")
        if len(values) >= 8:
            for var, value in zip((self.spine_action_name, self.spine_action_kind, self.spine_action_dur,
                                   self.spine_action_amp, self.spine_action_count, self.spine_action_repeat,
                                   self.spine_action_direction, self.spine_action_segments), values):
                var.set(str(value))
            try:
                segments = json.loads(self.spine_action_segments.get() or "[]")
            except json.JSONDecodeError:
                segments = []
            self._spine_form_kind_for_segments = str(values[1])
            self.spine_segment_status.set(f"关键帧：{len(segments) if isinstance(segments, list) else 0} 个")

    def _on_spine_action_kind_changed(self, _event=None):
        """Load the newly selected built-in template into the edit form immediately."""
        new_kind = str(self.spine_action_kind.get())
        previous_kind = str(getattr(self, "_spine_form_kind_for_segments", new_kind))
        if new_kind == previous_kind:
            return True
        try:
            current = {
                "kind": new_kind,
                "dur": float(self.spine_action_dur.get()),
                "amp": float(self.spine_action_amp.get()),
                "count": int(self.spine_action_count.get()),
                "repeat": int(self.spine_action_repeat.get()),
                "direction": self.spine_action_direction.get(),
                "segments": json.loads(self.spine_action_segments.get() or "[]"),
            }
            synced = spine_engine.sync_action_type_template(current, previous_kind)
            segments = synced.get("segments", [])
            self.spine_action_segments.set(json.dumps(segments, ensure_ascii=False, separators=(",", ":")))
            self.spine_action_dur.set(str(synced.get("dur", current["dur"])))
            self._spine_form_kind_for_segments = new_kind
            if new_kind in spine_engine.FACT:
                self.spine_segment_status.set(f"已载入【{new_kind}】模板 · {len(segments)} 个关键帧；可继续编辑")
            else:
                self.spine_segment_status.set(f"自定义类型 · 保留当前关键帧 {len(segments)} 个")
        except (ValueError, TypeError, json.JSONDecodeError) as exc:
            # Revert the selector so the visible type always agrees with the data.
            self.spine_action_kind.set(previous_kind)
            messagebox.showerror("无法切换动作模板", str(exc), parent=self.root)
            return False
        return True

    def _current_spine_form_action(self):
        action = {
            "name": self.spine_action_name.get().strip() or "Action",
            "kind": self.spine_action_kind.get(),
            "dur": float(self.spine_action_dur.get()),
            "amp": float(self.spine_action_amp.get()),
            "count": int(self.spine_action_count.get()),
            "repeat": int(self.spine_action_repeat.get()),
            "direction": self.spine_action_direction.get(),
            "segments": json.loads(self.spine_action_segments.get() or "[]"),
        }
        if action["dur"] <= 0 or action["count"] <= 0 or action["repeat"] <= 0:
            raise ValueError("时长、次数和重复次数必须大于 0。")
        if action["kind"] not in spine_engine.ACTION_TYPES:
            raise ValueError("不支持的动作类型。")
        if action["direction"] not in spine_engine.DIRECTIONS:
            raise ValueError("不支持的动作方向。")
        return action

    def _apply_spine_preset_to_form(self):
        try:
            action = self._current_spine_form_action()
            if action["kind"] not in spine_engine.FACT:
                raise ValueError("当前动作是【自定义】类型，不会被内置预设覆盖。")
            segments = spine_engine.FACT[action["kind"]](action["dur"], action["amp"], action["count"], action["direction"])
            if len(segments) < 2:
                raise ValueError("预设没有生成足够的关键帧。")
            action["segments"] = segments
            action["dur"] = float(segments[-1]["time"])
            self.spine_action_segments.set(json.dumps(segments, ensure_ascii=False, separators=(",", ":")))
            self.spine_action_dur.set(str(action["dur"]))
            self._spine_form_kind_for_segments = action["kind"]
            self.spine_segment_status.set(f"已生成【{action['kind']}】预设 · {len(segments)} 个关键帧")
            selected = self.spine_action_tree.selection()
            if selected:
                self._insert_spine_action(action, selected[0])
                self.spine_action_tree.selection_set(selected[0])
                self.spine_action_tree.focus(selected[0])
            self.save_config(silent=True)
        except (ValueError, TypeError, json.JSONDecodeError) as exc:
            messagebox.showerror("无法套用预设", str(exc), parent=self.root)

    def _apply_spine_presets_to_all(self):
        applied = 0
        skipped = 0
        for iid in self.spine_action_tree.get_children():
            values = list(self.spine_action_tree.item(iid, "values"))
            if len(values) < 8:
                continue
            kind = str(values[1])
            if kind not in spine_engine.FACT:
                skipped += 1
                continue
            try:
                duration = float(values[2]); amp = float(values[3]); count = int(values[4]); direction = str(values[6])
                if duration <= 0 or count <= 0:
                    raise ValueError("时长和次数必须大于 0")
                segments = spine_engine.FACT[kind](duration, amp, count, direction)
                values[2] = str(segments[-1]["time"])
                values[7] = json.dumps(segments, ensure_ascii=False, separators=(",", ":"))
                if len(values) > 8:
                    values[8] = str(len(segments))
                self.spine_action_tree.item(iid, values=values)
                applied += 1
            except Exception as exc:
                self.log(f"[Spine 预设] 跳过动作 {values[0]}：{exc}")
        selected = self.spine_action_tree.selection()
        if selected:
            self._select_spine_action()
        self.save_config(silent=True)
        messagebox.showinfo("批量套用预设", f"已重新生成 {applied} 个动作的关键帧。" + (f"\n跳过自定义/不适用动作：{skipped} 个。" if skipped else ""), parent=self.root)

    def edit_spine_keyframes(self):
        try:
            action = self._current_spine_form_action()
            original = self._normalise_spine_segments(action)
            if len(original) < 2:
                raise ValueError("动作至少需要两个关键帧。")
        except Exception as exc:
            messagebox.showerror("关键帧参数错误", str(exc), parent=self.root)
            return

        win = tk.Toplevel(self.root)
        win.title(f"关键帧编辑器 · {action['name']}")
        self.root.update_idletasks()
        screen_w, screen_h = win.winfo_screenwidth(), win.winfo_screenheight()
        win_w = min(1120, max(900, screen_w - 40))
        win_h = min(700, max(640, screen_h - 80))
        parent_x, parent_y = self.root.winfo_rootx(), self.root.winfo_rooty()
        parent_w = max(900, self.root.winfo_width())
        parent_h = max(640, self.root.winfo_height())
        pos_x = max(0, min(screen_w - win_w, parent_x + (parent_w - win_w) // 2))
        pos_y = max(0, min(screen_h - win_h, parent_y + (parent_h - win_h) // 2))
        win.geometry(f"{win_w}x{win_h}+{pos_x}+{pos_y}")
        win.minsize(min(900, win_w), min(640, win_h))
        win.transient(self.root)
        win.grab_set()
        colors = getattr(self, "ui_colors", {})
        win.configure(background=colors.get("bg", "#F3F6FB"))
        keys = [dict(k) for k in original]
        selected_index = {"value": 0}

        header = ttk.Frame(win, padding=(16, 4))
        header.pack(fill="x")
        ttk.Label(header, text="关键帧时间线", font=("Segoe UI Semibold", 16)).pack(anchor="w")
        ttk.Label(header, text=f"{action['name']}  ·  {action['kind']}  ·  双击/选中节点后可调整时间与位移；保存后将直接用于导出。", style="Muted.TLabel").pack(anchor="w", pady=(3, 0))

        timeline = tk.Canvas(win, height=52, background=colors.get("surface", "#FFFFFF"),
                             highlightthickness=1, highlightbackground=colors.get("border", "#DCE3EF"))
        timeline.pack(fill="x", padx=16, pady=(0, 6))

        body = ttk.Frame(win, padding=(16, 0, 16, 0))
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=3)
        body.columnconfigure(1, weight=2)
        body.rowconfigure(0, weight=1)

        table_box = ttk.LabelFrame(body, text="关键帧列表", padding=8)
        table_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        table_box.rowconfigure(0, weight=1)
        table_box.columnconfigure(0, weight=1)
        tree = ttk.Treeview(
            table_box,
            columns=("index", "time", "progress", "x", "y", "dx", "dy", "gap", "phase"),
            show="headings", selectmode="browse", height=4,
        )
        columns = (
            ("index", "#", 34), ("time", "时间(s)", 68), ("progress", "进度", 52),
            ("x", "X", 54), ("y", "Y", 54), ("dx", "ΔX", 54), ("dy", "ΔY", 54),
            ("gap", "区间(s)", 64), ("phase", "阶段", 58),
        )
        for col, title, width in columns:
            tree.heading(col, text=title)
            tree.column(col, width=width, minwidth=width, anchor="center", stretch=(col in {"time", "x", "y"}))
        tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll = ttk.Scrollbar(table_box, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        ttk.Label(table_box, text="双击时间 / X / Y 单元格可直接编辑；时间必须递增，至少保留两个节点。", style="Muted.TLabel").grid(row=1, column=0, columnspan=2, sticky="w", pady=(5, 0))

        edit_box = ttk.LabelFrame(body, text="编辑所选关键帧", padding=6)
        edit_box.grid(row=0, column=1, sticky="nsew")
        edit_box.columnconfigure(1, weight=1)
        time_var = tk.StringVar()
        x_var = tk.StringVar()
        y_var = tk.StringVar()
        step_time_var = tk.StringVar(value="0.01")
        step_pos_var = tk.StringVar(value="5")
        for idx, (label, var) in enumerate((("时间 (秒)", time_var), ("X 位移", x_var), ("Y 位移", y_var))):
            ttk.Label(edit_box, text=label).grid(row=idx, column=0, sticky="w", pady=1, padx=(0, 8))
            ttk.Entry(edit_box, textvariable=var, width=16).grid(row=idx, column=1, sticky="ew", pady=1)
        ttk.Button(edit_box, text="应用数值", style="Accent.TButton", command=lambda: apply_editor_values()).grid(row=3, column=0, columnspan=2, sticky="ew", pady=(3, 4))
        ttk.Separator(edit_box).grid(row=4, column=0, columnspan=2, sticky="ew", pady=2)
        ttk.Label(edit_box, text="快速步进", font=("Segoe UI Semibold", 9)).grid(row=5, column=0, columnspan=2, sticky="w", pady=(1, 2))
        ttk.Label(edit_box, text="时间步进").grid(row=6, column=0, sticky="w", pady=1)
        ttk.Entry(edit_box, textvariable=step_time_var, width=10).grid(row=6, column=1, sticky="ew", pady=1)
        ttk.Label(edit_box, text="位置步进").grid(row=7, column=0, sticky="w", pady=1)
        ttk.Entry(edit_box, textvariable=step_pos_var, width=10).grid(row=7, column=1, sticky="ew", pady=1)
        time_buttons = ttk.Frame(edit_box)
        time_buttons.grid(row=8, column=0, columnspan=2, sticky="ew", pady=(3, 1))
        ttk.Button(time_buttons, text="时间 −", command=lambda: nudge("time", -1)).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(time_buttons, text="时间 +", command=lambda: nudge("time", 1)).pack(side="left", expand=True, fill="x", padx=(3, 0))
        xy_buttons = ttk.Frame(edit_box)
        xy_buttons.grid(row=9, column=0, columnspan=2, sticky="ew", pady=1)
        ttk.Button(xy_buttons, text="X −", command=lambda: nudge("x", -1)).pack(side="left", expand=True, fill="x", padx=(0, 3))
        ttk.Button(xy_buttons, text="X +", command=lambda: nudge("x", 1)).pack(side="left", expand=True, fill="x", padx=3)
        ttk.Button(xy_buttons, text="Y −", command=lambda: nudge("y", -1)).pack(side="left", expand=True, fill="x", padx=3)
        ttk.Button(xy_buttons, text="Y +", command=lambda: nudge("y", 1)).pack(side="left", expand=True, fill="x", padx=(3, 0))
        ttk.Button(edit_box, text="所选 X/Y 归零", command=lambda: zero_selected()).grid(row=10, column=0, columnspan=2, sticky="ew", pady=(2, 0))

        operations = ttk.LabelFrame(win, text="关键帧操作", padding=(10, 5))
        operations.pack(fill="x", padx=16, pady=(6, 4))
        operation_specs = [
            ("添加末尾", lambda: add_end()), ("插入中点", lambda: insert_midpoint()), ("复制选中", lambda: duplicate_selected()),
            ("删除选中", lambda: delete_selected()), ("均匀分布时间", lambda: uniform_times()), ("全部 X/Y 归零", lambda: zero_all()),
        ]
        if action["kind"] in spine_engine.FACT:
            operation_specs.append(("重新生成预设", lambda: regenerate_preset()))
        for i, (label, callback) in enumerate(operation_specs):
            ttk.Button(operations, text=label, command=callback).grid(row=0, column=i, sticky="ew", padx=2, pady=1)
            operations.columnconfigure(i, weight=1)
        footer = ttk.Frame(win, padding=(16, 0, 16, 8))
        footer.pack(fill="x")
        status_var = tk.StringVar(value=f"{len(keys)} 个关键帧")
        ttk.Label(footer, textvariable=status_var, style="Muted.TLabel").pack(side="left")
        ttk.Button(footer, text="取消", command=win.destroy).pack(side="right", padx=(6, 0))
        ttk.Button(footer, text="保存关键帧", style="Accent.TButton", command=lambda: save_editor()).pack(side="right")

        def draw_timeline():
            timeline.delete("all")
            width = max(300, timeline.winfo_width())
            height = max(42, timeline.winfo_height())
            left, right = 26, width - 24
            mid = int(height * 0.52)
            timeline.create_line(left, mid, right, mid, fill=colors.get("border", "#DCE3EF"), width=2)
            end_t = float(keys[-1]["time"]) if keys else 1.0
            if end_t <= 0:
                end_t = 1.0
            for idx, key in enumerate(keys):
                xx = left + (right - left) * float(key["time"]) / end_t
                selected = idx == selected_index["value"]
                r = 5 if selected else 4
                fill = colors.get("primary", "#4F46E5") if selected else "#93A4C7"
                timeline.create_oval(xx-r, mid-r, xx+r, mid+r, fill=fill, outline=fill)
                label = f"K{idx+1} · {float(key['time']):.3f}s"
                label_y = 13 if idx % 2 == 0 else height - 15
                if idx == 0:
                    timeline.create_text(left + 2, label_y, text=label, anchor="w",
                                         fill=colors.get("text", "#172033"), font=("Segoe UI", 8))
                elif idx == len(keys) - 1:
                    timeline.create_text(right - 2, label_y, text=label, anchor="e",
                                         fill=colors.get("text", "#172033"), font=("Segoe UI", 8))
                else:
                    timeline.create_text(xx, label_y, text=label, anchor="center",
                                         fill=colors.get("text", "#172033"), font=("Segoe UI", 8))
            timeline.create_text(left, height - 5, text="0s", anchor="w", fill=colors.get("muted", "#64748B"), font=("Segoe UI", 8))
            timeline.create_text(right, height - 5, text=f"{end_t:.3f}s", anchor="e", fill=colors.get("muted", "#64748B"), font=("Segoe UI", 8))

        def refresh(select_idx=None):
            if select_idx is not None:
                selected_index["value"] = max(0, min(int(select_idx), len(keys)-1)) if keys else 0
            for iid in tree.get_children():
                tree.delete(iid)
            end_time = max(float(keys[-1]["time"]), 0.000001) if keys else 1.0
            for idx, key in enumerate(keys):
                tm = float(key["time"])
                previous = keys[idx - 1] if idx > 0 else {"x": 0.0, "y": 0.0}
                next_key = keys[idx + 1] if idx < len(keys) - 1 else None
                dx = float(key["x"]) - float(previous["x"])
                dy = float(key["y"]) - float(previous["y"])
                gap = max(0.0, float(next_key["time"]) - tm) if next_key else 0.0
                phase = "起点" if idx == 0 else ("结束" if idx == len(keys) - 1 else f"变化 {idx}")
                iid = str(idx)
                tree.insert(
                    "", "end", iid=iid,
                    values=(idx + 1, f"{tm:.4f}", f"{100.0 * tm / end_time:.1f}%",
                            f"{float(key['x']):.3f}", f"{float(key['y']):.3f}",
                            f"{dx:.3f}", f"{dy:.3f}", f"{gap:.4f}", phase),
                    tags=("odd" if idx % 2 else "even",),
                )
            tree.tag_configure("odd", background="#F8FAFD")
            tree.tag_configure("even", background="#FFFFFF")
            if keys:
                idx = selected_index["value"]
                iid = str(idx)
                if tree.exists(iid):
                    tree.selection_set(iid)
                    tree.focus(iid)
                    tree.see(iid)
                set_form_from_key(idx)
            status_var.set(f"{len(keys)} 个关键帧  ·  总时长 {float(keys[-1]['time']):.4f}s" if keys else "没有关键帧")
            draw_timeline()

        def set_form_from_key(idx):
            if not 0 <= idx < len(keys):
                return
            time_var.set(f"{float(keys[idx]['time']):.4f}")
            x_var.set(f"{float(keys[idx]['x']):.3f}")
            y_var.set(f"{float(keys[idx]['y']):.3f}")

        def on_select(_event=None):
            selection = tree.selection()
            if not selection:
                return
            idx = int(selection[0])
            selected_index["value"] = idx
            set_form_from_key(idx)
            draw_timeline()

        def apply_editor_values():
            if not keys:
                return False
            idx = selected_index["value"]
            try:
                tm, xx, yy = float(time_var.get()), float(x_var.get()), float(y_var.get())
                if not all(math.isfinite(value) for value in (tm, xx, yy)):
                    raise ValueError("时间和 X/Y 必须是有限数字，不能填写 NaN 或 Infinity。")
                if tm < 0:
                    raise ValueError("时间不能小于 0。")
                if idx > 0 and tm <= float(keys[idx-1]["time"]):
                    raise ValueError("当前关键帧时间必须大于上一个关键帧。")
                if idx < len(keys)-1 and tm >= float(keys[idx+1]["time"]):
                    raise ValueError("当前关键帧时间必须小于下一个关键帧。")
                keys[idx] = {"time": tm, "x": xx, "y": yy}
                refresh(idx)
                return True
            except (ValueError, TypeError) as exc:
                messagebox.showerror("关键帧值无效", str(exc), parent=win)
                return False

        def selected_idx():
            sel = tree.selection()
            return int(sel[0]) if sel else selected_index["value"]

        def nudge(axis, sign):
            idx = selected_idx()
            if not 0 <= idx < len(keys):
                return
            try:
                step = float(step_time_var.get() if axis == "time" else step_pos_var.get())
                if step <= 0:
                    raise ValueError("步进必须大于 0。")
                key = dict(keys[idx])
                if axis == "time":
                    target = float(key["time"]) + sign * step
                    prev = float(keys[idx-1]["time"]) if idx > 0 else -0.000001
                    nxt = float(keys[idx+1]["time"]) if idx < len(keys)-1 else float("inf")
                    if idx == 0:
                        target = max(0.0, target)
                    else:
                        target = max(prev + 0.000001, target)
                    if idx < len(keys)-1:
                        target = min(nxt - 0.000001, target)
                    key["time"] = round(target, 6)
                else:
                    key[axis] = round(float(key[axis]) + sign * step, 6)
                keys[idx] = key
                refresh(idx)
            except (ValueError, TypeError) as exc:
                messagebox.showerror("无效步进", str(exc), parent=win)

        def zero_selected():
            idx = selected_idx()
            if 0 <= idx < len(keys):
                keys[idx]["x"] = 0.0
                keys[idx]["y"] = 0.0
                refresh(idx)

        def add_end():
            last = dict(keys[-1]) if keys else {"time": 0.0, "x": 0.0, "y": 0.0}
            try:
                step = max(0.05, float(step_time_var.get()))
            except ValueError:
                step = 0.05
            new_time = float(last["time"]) + step
            keys.append({"time": round(new_time, 6), "x": 0.0, "y": 0.0})
            refresh(len(keys)-1)

        def insert_midpoint():
            idx = selected_idx()
            if idx >= len(keys)-1:
                messagebox.showinfo("插入中点", "请选中一个非末尾关键帧。", parent=win)
                return
            a, b = keys[idx], keys[idx+1]
            mid_key = {axis: (float(a[axis]) + float(b[axis])) / 2 for axis in ("time", "x", "y")}
            if mid_key["time"] <= float(a["time"]) or mid_key["time"] >= float(b["time"]):
                return
            keys.insert(idx+1, mid_key)
            refresh(idx+1)

        def duplicate_selected():
            idx = selected_idx()
            if not 0 <= idx < len(keys):
                return
            source_key = dict(keys[idx])
            if idx < len(keys)-1:
                next_time = (float(source_key["time"]) + float(keys[idx+1]["time"])) / 2
            else:
                try:
                    next_time = float(source_key["time"]) + max(0.05, float(step_time_var.get()))
                except ValueError:
                    next_time = float(source_key["time"]) + 0.05
            if idx < len(keys)-1 and not (float(source_key["time"]) < next_time < float(keys[idx+1]["time"])):
                messagebox.showwarning("无法复制", "两个关键帧之间没有足够的时间插入新节点。", parent=win)
                return
            clone = {"time": next_time, "x": float(source_key["x"]), "y": float(source_key["y"])}
            keys.insert(idx+1, clone)
            refresh(idx+1)

        def delete_selected():
            if len(keys) <= 2:
                messagebox.showinfo("不能删除", "至少需要保留两个关键帧。", parent=win)
                return
            idx = selected_idx()
            if 0 <= idx < len(keys):
                keys.pop(idx)
                refresh(min(idx, len(keys)-1))

        def uniform_times():
            if len(keys) < 2:
                return
            start_t = float(keys[0]["time"])
            end_t = float(keys[-1]["time"])
            if end_t <= start_t:
                messagebox.showerror("无法均匀分布", "首尾关键帧时间必须有间隔。", parent=win)
                return
            for idx in range(1, len(keys)-1):
                keys[idx]["time"] = start_t + (end_t-start_t) * idx / (len(keys)-1)
            refresh(selected_idx())

        def zero_all():
            for key in keys:
                key["x"] = 0.0
                key["y"] = 0.0
            refresh(selected_idx())

        def regenerate_preset():
            try:
                current = self._current_spine_form_action()
                if current["kind"] not in spine_engine.FACT:
                    raise ValueError("自定义动作不能重新生成内置预设。")
                new_keys = spine_engine.FACT[current["kind"]](current["dur"], current["amp"], current["count"], current["direction"])
                if len(new_keys) < 2:
                    raise ValueError("预设没有生成有效的关键帧。")
                keys[:] = [dict(k) for k in new_keys]
                refresh(0)
            except Exception as exc:
                messagebox.showerror("生成预设失败", str(exc), parent=win)

        def save_editor():
            # Commit the visible fields too, so Save never silently discards the last typed edit.
            if not apply_editor_values():
                return
            try:
                if len(keys) < 2:
                    raise ValueError("至少需要两个关键帧。")
                normalized = []
                prev = -1.0
                for idx, key in enumerate(keys, 1):
                    tm, xx, yy = float(key["time"]), float(key["x"]), float(key["y"])
                    if not all(math.isfinite(value) for value in (tm, xx, yy)):
                        raise ValueError(f"第 {idx} 个关键帧包含无效数值。")
                    if tm < 0 or tm <= prev:
                        raise ValueError(f"第 {idx} 个关键帧时间必须非负且严格递增。")
                    normalized.append({"time": tm, "x": xx, "y": yy})
                    prev = tm
                if prev <= 0:
                    raise ValueError("最后一个关键帧时间必须大于 0。")
                self.spine_action_segments.set(json.dumps(normalized, ensure_ascii=False, separators=(",", ":")))
                self.spine_action_dur.set(str(round(prev, 6)))
                self.spine_segment_status.set(f"关键帧已保存 · {len(normalized)} 个")
                selected = self.spine_action_tree.selection()
                if selected:
                    current_action = {
                        "name": self.spine_action_name.get().strip() or "Action",
                        "kind": self.spine_action_kind.get(), "dur": prev,
                        "amp": float(self.spine_action_amp.get()), "count": int(self.spine_action_count.get()),
                        "repeat": int(self.spine_action_repeat.get()), "direction": self.spine_action_direction.get(),
                        "segments": normalized,
                    }
                    self._insert_spine_action(current_action, selected[0])
                    self.spine_action_tree.selection_set(selected[0])
                    self.spine_action_tree.focus(selected[0])
                    self._select_spine_action()
                self.save_config(silent=True)
                win.destroy()
            except Exception as exc:
                messagebox.showerror("关键帧参数错误", str(exc), parent=win)

        def begin_cell_edit(event):
            item = tree.identify_row(event.y)
            column = tree.identify_column(event.x)
            editable = {"#2": ("time", 1), "#4": ("x", 3), "#5": ("y", 4)}
            if not item or column not in editable:
                return
            box = tree.bbox(item, column)
            if not box:
                return
            idx = int(item)
            field, value_index = editable[column]
            x, y, width, height = box
            editor = ttk.Entry(tree, justify="center")
            editor.place(x=x, y=y, width=width, height=height)
            editor.insert(0, tree.set(item, field))
            editor.select_range(0, "end")
            editor.focus_set()
            committed = {"done": False}

            def finish(commit=True):
                if committed["done"]:
                    return
                committed["done"] = True
                raw = editor.get().strip()
                try:
                    if commit:
                        number = float(raw)
                        if not math.isfinite(number):
                            raise ValueError("请输入有限数字，不能使用 NaN 或 Infinity。")
                        if field == "time":
                            prev = float(keys[idx - 1]["time"]) if idx > 0 else -1e-9
                            nxt = float(keys[idx + 1]["time"]) if idx < len(keys) - 1 else float("inf")
                            if number < 0 or number <= prev or number >= nxt:
                                raise ValueError("时间必须非负，并严格位于相邻关键帧之间。")
                        keys[idx][field] = number
                        refresh(idx)
                except (ValueError, TypeError) as exc:
                    messagebox.showerror("关键帧值无效", str(exc), parent=win)
                    refresh(idx)
                finally:
                    if editor.winfo_exists():
                        editor.destroy()

            editor.bind("<Return>", lambda _e: finish(True))
            editor.bind("<Escape>", lambda _e: finish(False))
            editor.bind("<FocusOut>", lambda _e: finish(True))

        def select_timeline_key(event):
            if not keys:
                return
            width = max(300, timeline.winfo_width())
            left, right = 26, width - 24
            if right <= left:
                return
            ratio = min(1.0, max(0.0, (event.x - left) / (right - left)))
            end_t = max(float(keys[-1]["time"]), 0.000001)
            target_time = ratio * end_t
            nearest = min(range(len(keys)), key=lambda idx: abs(float(keys[idx]["time"]) - target_time))
            refresh(nearest)

        tree.bind("<<TreeviewSelect>>", on_select)
        tree.bind("<Double-1>", begin_cell_edit)
        timeline.bind("<Button-1>", select_timeline_key)
        timeline.bind("<Configure>", lambda _e: draw_timeline())
        refresh(0)

    def _read_spine_action_form(self):
        # The event binding handles normal UI interaction; this guard also covers
        # programmatic type changes and ensures Add/Update can never save a stale template.
        if str(self.spine_action_kind.get()) != str(getattr(self, "_spine_form_kind_for_segments", self.spine_action_kind.get())):
            if not self._on_spine_action_kind_changed():
                raise ValueError("动作类型尚未同步成功，请检查参数后重试。")
        action = self._current_spine_form_action()
        action["segments"] = self._normalise_spine_segments(action)
        if action["kind"] == "自定义" and len(action["segments"]) < 2:
            raise ValueError("自定义动作至少需要两个关键帧。")
        return action

    def add_spine_action(self):
        try:
            action = self._read_spine_action_form()
            # Use the form's synchronized keyframes as-is. Regenerating on every Add
            # would discard manual edits made after the selected template was loaded.
            self._insert_spine_action(action)
            last = self.spine_action_tree.get_children()[-1]
            self.spine_action_tree.selection_set(last)
            self.spine_action_tree.focus(last)
            self._select_spine_action()
            self.save_config(silent=True)
        except (ValueError, TypeError, json.JSONDecodeError) as exc:
            messagebox.showerror("动作参数错误", str(exc), parent=self.root)

    def update_spine_action(self):
        selected = self.spine_action_tree.selection()
        if not selected:
            messagebox.showinfo("选择动作", "请先在动作列表中选中要更新的动作。", parent=self.root)
            return
        try:
            action = self._read_spine_action_form()
            # Match the original editor: changing total duration scales the whole timeline
            # instead of leaving a stale final keyframe behind.
            segments = action["segments"]
            last_time = float(segments[-1]["time"]) if segments else 0.0
            requested_duration = float(action["dur"])
            if last_time <= 0:
                raise ValueError("最后一个关键帧时间必须大于 0。")
            if abs(requested_duration - last_time) > 1e-6:
                ratio = requested_duration / last_time
                for key in segments:
                    key["time"] = round(float(key["time"]) * ratio, 6)
                action["segments"] = segments
            self._insert_spine_action(action, selected[0])
            self.spine_action_tree.selection_set(selected[0])
            self.spine_action_tree.focus(selected[0])
            self._select_spine_action()
            self.save_config(silent=True)
        except (ValueError, TypeError, json.JSONDecodeError) as exc:
            messagebox.showerror("动作参数错误", str(exc), parent=self.root)

    def duplicate_spine_action(self):
        selected = self.spine_action_tree.selection()
        if not selected:
            messagebox.showinfo("复制动作", "请先选择要复制的动作。", parent=self.root)
            return
        values = self.spine_action_tree.item(selected[0], "values")
        if len(values) < 8:
            return
        action = {
            "name": str(values[0]) + "_Copy", "kind": str(values[1]), "dur": float(values[2]),
            "amp": float(values[3]), "count": int(values[4]), "repeat": int(values[5]),
            "direction": str(values[6]), "segments": json.loads(values[7] or "[]"),
        }
        existing_names = {str(self.spine_action_tree.item(i, "values")[0]) for i in self.spine_action_tree.get_children()}
        base = action["name"]
        suffix = 2
        while action["name"] in existing_names:
            action["name"] = f"{base}{suffix}"
            suffix += 1
        self._insert_spine_action(action)
        last = self.spine_action_tree.get_children()[-1]
        self.spine_action_tree.selection_set(last)
        self.spine_action_tree.focus(last)
        self._select_spine_action()
        self.save_config(silent=True)

    def remove_spine_action(self):
        for iid in self.spine_action_tree.selection():
            self.spine_action_tree.delete(iid)
        self.save_config(silent=True)

    def reset_spine_actions(self):
        if not messagebox.askyesno("恢复默认动作", "将用内置的 Jump / DoubleJump / Sway 替换当前动作列表。继续？", parent=self.root):
            return
        for iid in self.spine_action_tree.get_children():
            self.spine_action_tree.delete(iid)
        for action in spine_engine.DEFAULT_ACTIONS:
            self._insert_spine_action(action)
        children = self.spine_action_tree.get_children()
        if children:
            self.spine_action_tree.selection_set(children[0])
            self._select_spine_action()
        self.save_config(silent=True)

    def _collect_spine_actions(self, silent_invalid=False):
        result = []
        for iid in self.spine_action_tree.get_children() if hasattr(self, "spine_action_tree") else []:
            values = self.spine_action_tree.item(iid, "values")
            try:
                action = {
                    "name": str(values[0]).strip() or "Action", "kind": str(values[1]),
                    "dur": float(values[2]), "amp": float(values[3]), "count": int(values[4]),
                    "repeat": int(values[5]), "direction": str(values[6]),
                    "segments": json.loads(values[7] or "[]"),
                }
                if action["kind"] not in spine_engine.ACTION_TYPES or action["direction"] not in spine_engine.DIRECTIONS:
                    raise ValueError("动作类型或方向不支持。")
                if action["dur"] <= 0 or action["count"] <= 0 or action["repeat"] <= 0:
                    raise ValueError("时长、次数和重复次数必须大于 0。")
                action["segments"] = self._normalise_spine_segments(action)
                if len(action["segments"]) < 2:
                    raise ValueError("动作至少需要两个关键帧。")
                result.append(action)
            except Exception:
                if not silent_invalid:
                    raise ValueError(f"动作行参数无效：{values}")
        return result

    def run_spine_actions_independent(self):
        if self.busy:
            messagebox.showinfo("任务运行中", "当前已有任务正在执行，请等待结束后再启动 Spine 动作处理。", parent=self.root)
            return
        exe_text = str(self.vars["spine_converter_path"].get() or "").strip()
        exe = self._spine_resolved_path(exe_text)
        if not exe or not exe.is_file():
            messagebox.showerror("缺少 Converter", "请先选择有效的 SpineSkeletonDataConverter.exe。", parent=self.root)
            return
        source = self._spine_source_root()
        output = self._spine_resolved_path(self.vars["spine_output_dir"].get(), DEFAULT_CONFIG["spine_output_dir"])
        if not output:
            messagebox.showerror("缺少输出目录", "请指定动作输出目录。", parent=self.root)
            return
        files = [Path(self.spine_file_tree.item(iid, "values")[0]).resolve() for iid in self.spine_file_tree.selection()]
        if not files:
            files = list(self.spine_files)
        if not files:
            try:
                files = spine_engine.scan_target_skel_files(source, output)
            except Exception as exc:
                messagebox.showerror("扫描失败", str(exc), parent=self.root)
                return
        files = [p for p in files if spine_engine.is_sd_ch_spine(p)]
        if not files:
            messagebox.showwarning("没有目标文件", "请先扫描/添加 SD 或 CH 的 .skel 文件。", parent=self.root)
            return
        try:
            actions = self._collect_spine_actions()
            if not actions:
                raise ValueError("至少需要一个全局动作。")
        except Exception as exc:
            messagebox.showerror("动作参数错误", str(exc), parent=self.root)
            return
        options = {
            "root_bone": str(self.vars["spine_root_bone"].get() or "").strip(),
            "make_skel": bool(self.vars["spine_make_skel"].get()),
            "remove_curve": bool(self.vars["spine_remove_curve"].get()),
            "keep_json": bool(self.vars["spine_keep_json"].get()),
            "backup": bool(self.vars["spine_backup"].get()),
        }
        self.save_config(silent=True)
        self.notebook.select(self.spine_tab)

        def work():
            spine_engine.run_batch(files, source, output, exe, actions, **options,
                                   log=lambda msg: self.post({"type": "log", "text": msg}))

        self.launch_task("Spine 全局动作（独立运行）", work)

    def _run_auto_spine_actions_after_export(self, export_root):
        """Optional post-export hook; its setting is false unless explicitly enabled."""
        if not bool(self._config.get("spine_auto_enabled", False)):
            return
        converter = self._spine_resolved_path(self._config.get("spine_converter_path", ""))
        if not converter or not converter.is_file():
            self.post({"type": "warning", "text": "已启用导出后自动处理 SD/CH Spine，但没有配置有效的 SpineSkeletonDataConverter.exe；跳过动作处理。"})
            return
        output = self._spine_resolved_path(self._config.get("spine_output_dir", DEFAULT_CONFIG["spine_output_dir"]), DEFAULT_CONFIG["spine_output_dir"])
        if not output:
            self.post({"type": "warning", "text": "导出后 Spine 动作处理未运行：输出目录无效。"})
            return
        try:
            files = spine_engine.scan_target_skel_files(export_root, output)
        except Exception as exc:
            self.post({"type": "warning", "text": f"导出后 Spine 扫描失败：{exc}"})
            return
        if not files:
            self.post({"type": "log", "text": "导出后 Spine 自动处理：没有发现 SD/CH .skel 文件，跳过。"})
            return
        actions = self._config.get("spine_actions", spine_engine.DEFAULT_ACTIONS)
        if not isinstance(actions, list) or not actions:
            self.post({"type": "warning", "text": "导出后 Spine 自动处理未运行：没有配置全局动作。"})
            return
        options = {
            "root_bone": str(self._config.get("spine_root_bone", "") or "").strip(),
            "make_skel": bool(self._config.get("spine_make_skel", False)),
            "remove_curve": bool(self._config.get("spine_remove_curve", False)),
            "keep_json": bool(self._config.get("spine_keep_json", False)),
            "backup": bool(self._config.get("spine_backup", True)),
        }
        self.post({"type": "log", "text": f"导出后 Spine 自动处理：发现 {len(files)} 个 SD/CH 文件，输出到 {output}"})
        result = spine_engine.run_batch(files, export_root, output, converter, actions, **options,
                                        log=lambda msg: self.post({"type": "log", "text": msg}))
        self.post({"type": "summary", "text": f"Spine 动作处理完成 · 成功 {result['ok']} · 失败 {result['failed']}"})

    def _build_logs_tab(self):
        frame = ttk.Frame(self.logs_tab)
        frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(frame, wrap="word", state="disabled", font=("Consolas", 9), undo=False)
        sy = ttk.Scrollbar(frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=sy.set)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        sy.grid(row=0, column=1, sticky="ns")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        bottom = ttk.Frame(self.logs_tab)
        bottom.pack(fill="x", pady=(7, 0))
        ttk.Button(bottom, text="清空界面日志", command=self.clear_log_view).pack(side="left")
        ttk.Button(bottom, text="保存日志文本…", command=self.save_log_view).pack(side="left", padx=(6, 0))

    def _apply_config_to_vars(self):
        for key, var in self.vars.items():
            val = self._config.get(key, DEFAULT_CONFIG.get(key, ""))
            try:
                if isinstance(var, tk.BooleanVar):
                    var.set(bool(val))
                else:
                    var.set(str(val))
            except tk.TclError:
                pass

    def _restore_log(self):
        self.log(f"程序启动。Python {sys.version.split()[0]}；核心版本 V{core.VERSION}。")
        self.log("下载/解析/导出沿用 V7 核心；角色路径遵循 ID 归档并去除冗余包装层；无编号资源按类型归档；LowQuality 默认跳过。SD/CH Spine 全局动作已整合，导出后自动执行默认关闭。")

    def log(self, text):
        timestamp = time.strftime("%H:%M:%S")
        if not hasattr(self, "log_text"):
            return
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{timestamp}] {text}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def clear_log_view(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def save_log_view(self):
        path = filedialog.asksaveasfilename(parent=self.root, title="保存日志", defaultextension=".txt", filetypes=[("文本文件", "*.txt"), ("全部文件", "*.*")])
        if not path:
            return
        try:
            Path(path).write_text(self.log_text.get("1.0", "end-1c"), encoding="utf-8")
            self.log(f"界面日志已保存：{path}")
        except OSError as e:
            messagebox.showerror("保存失败", str(e), parent=self.root)

    def post(self, event):
        self.events.put(event)

    def _drain_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                kind = event.get("type")
                if kind == "log":
                    self.log(event.get("text", ""))
                elif kind == "busy":
                    self._set_busy(bool(event.get("value")), event.get("task", ""))
                elif kind == "error":
                    self.log("错误：" + event.get("text", "未知错误"))
                    messagebox.showerror("操作失败", event.get("text", "未知错误"), parent=self.root)
                elif kind == "info":
                    self.log(event.get("text", ""))
                    if event.get("dialog"):
                        messagebox.showinfo(event.get("title", "提示"), event["dialog"], parent=self.root)
                elif kind == "warning":
                    self.log("警告：" + event.get("text", ""))
                    if event.get("dialog"):
                        messagebox.showwarning(event.get("title", "提示"), event["dialog"], parent=self.root)
                elif kind == "rows_loaded":
                    self._set_rows(event["rows"], event.get("source", "未知来源"))
                elif kind == "progress":
                    self._update_task_row(event)
                elif kind == "summary":
                    self.task_summary_var.set(event.get("text", ""))
                elif kind == "catalog_progress":
                    total = event.get("total")
                    done = event.get("done", 0)
                    if total:
                        self.status_var.set(f"正在下载 Catalog：{core.human_size(done)} / {core.human_size(total)}")
                    else:
                        self.status_var.set(f"正在下载 Catalog：{core.human_size(done)}")
                elif kind == "clear_task_rows":
                    for item in self.task_tree.get_children():
                        self.task_tree.delete(item)
                    self.download_tree_items.clear()
                elif kind == "refresh_rows":
                    self.apply_filter()
        except queue.Empty:
            pass
        self.root.after(100, self._drain_events)

    def _set_busy(self, value: bool, task=""):
        self.busy = value
        if value:
            self.current_task = task
            self.status_var.set(f"正在执行：{task}")
            self.global_progress.start(10)
            self.cancel_btn.configure(state="normal" if "下载" in task else "disabled")
            self.pause_btn.configure(state="normal" if task.startswith("下载资源") else "disabled", text="暂停下载")
        else:
            self.current_task = ""
            self.status_var.set("就绪")
            self.global_progress.stop()
            self.cancel_btn.configure(state="disabled")
            self.pause_btn.configure(state="disabled", text="暂停下载")
            self.resume_event.set()
            self.cancel_event.clear()
        for btn in (self.import_btn, self.scan_btn, self.catalog_btn, self.update_catalog_btn, self.preflight_btn, self.download_btn, self.download_export_btn, self.export_btn):
            btn.configure(state="disabled" if value else "normal")

    def launch_task(self, title, fn):
        if self.busy:
            messagebox.showinfo("任务运行中", "当前已有任务正在执行，请等待结束后再启动新任务。", parent=self.root)
            return False
        # Snapshot all Tk variables on the UI thread before background workers use settings.
        self.save_config(silent=True)
        self.cancel_event.clear()
        self.resume_event.set()
        for item in self.task_tree.get_children():
            self.task_tree.delete(item)
        self.download_tree_items.clear()
        self._set_busy(True, title)
        self.log(f"开始：{title}")

        def runner():
            try:
                fn()
                self.post({"type": "log", "text": f"完成：{title}"})
            except Exception as e:
                detail = f"{type(e).__name__}: {e}"
                self.post({"type": "error", "text": detail})
                self.post({"type": "log", "text": traceback.format_exc()})
            finally:
                self.post({"type": "busy", "value": False, "task": title})
        threading.Thread(target=runner, name="GUI-" + re.sub(r"\W+", "_", title), daemon=True).start()
        return True

    def auto_load_source(self):
        if self.busy:
            return
        try:
            remembered = str(self._config.get("resource_file", "") or "").strip()
            if remembered and Path(remembered).expanduser().is_file():
                self.load_file(Path(remembered).expanduser().resolve())
                return
            csvs = core.find_csv_candidates()
            if csvs:
                self.load_file(Path(csvs[0]))
                return
            catalogs = core.find_catalog_text_candidates()
            bundles = core.find_catalog_bundle_candidates()
            if catalogs:
                self.analyze_catalog_path(Path(catalogs[0]))
                return
            if bundles:
                self.analyze_catalog_path(Path(bundles[0]))
                return
            self.status_var.set("未找到本地资源表。请导入 CSV/XLSX/JSON，或分析 Catalog。")
            self.log("自动扫描未找到资源表或本地 Catalog；可导入资源表，或点击“在线更新 Catalog”。")
        except Exception as e:
            self.log(f"自动扫描失败：{type(e).__name__}: {e}")

    def choose_resource_file(self):
        path = filedialog.askopenfilename(parent=self.root, title="选择资源表", filetypes=[("资源表", "*.csv *.xlsx *.json *.txt"), ("CSV", "*.csv"), ("Excel", "*.xlsx"), ("JSON", "*.json"), ("全部文件", "*.*")])
        if path:
            self.load_file(Path(path))

    def load_file(self, path: Path):
        def work():
            rows = read_resource_file(path)
            self.post({"type": "rows_loaded", "rows": rows, "source": str(path)})
            self.post({"type": "log", "text": f"资源表读取成功：{path}；记录 {len(rows):,} 条。"})
        self.launch_task("加载资源表", work)

    def choose_catalog_file(self):
        path = filedialog.askopenfilename(parent=self.root, title="选择 Catalog 或 Catalog Bundle", filetypes=[("Catalog 文件", "*.json *.txt *.bundle"), ("全部文件", "*.*")])
        if path:
            self.analyze_catalog_path(Path(path))

    def analyze_catalog_path(self, path: Path):
        def work():
            self.post({"type": "log", "text": f"正在分析 Catalog：{path}"})
            p = path
            if p.suffix.casefold() == ".bundle":
                extracted = ROOT_DIR / "catalog.json"
                core.extract_catalog_from_bundle(p, extracted)
                p = extracted
            rows = core.analyze_catalog(p)
            csv_path, xlsx_path = core.write_mapping_files(ROOT_DIR, rows)
            self.post({"type": "rows_loaded", "rows": normalize_rows(rows), "source": f"Catalog / {path}"})
            self.post({"type": "log", "text": f"Catalog 分析完成：{len(rows):,} 条；生成资源表 {csv_path} 和 {xlsx_path}"})
        self.launch_task("分析 Catalog", work)

    def update_catalog(self):
        if not messagebox.askyesno("更新 Catalog", "将从设置中的 Catalog URL 下载最新 catalog.bundle，并重新生成 CSV/XLSX 资源表。继续？", parent=self.root):
            return
        def work():
            url = self._string_setting("catalog_bundle_url").strip() or core.CATALOG_BUNDLE_URL
            timeout, retries = self._numeric_setting("timeout", 60), self._numeric_setting("retries", 3)
            dest = ROOT_DIR / "catalog.bundle"
            part = dest.with_name(dest.name + ".part")
            last_error = None
            for attempt in range(1, retries + 1):
                try:
                    existing = part.stat().st_size if part.exists() else 0
                    headers = {"User-Agent": core.USER_AGENT, "Accept-Encoding": "identity"}
                    if existing:
                        headers["Range"] = f"bytes={existing}-"
                    req = urllib.request.Request(url, headers=headers)
                    with urllib.request.urlopen(req, timeout=timeout) as response:
                        status = getattr(response, "status", response.getcode())
                        clen = response.headers.get("Content-Length")
                        remain = int(clen) if clen and clen.isdigit() else None
                        if existing and status == 206:
                            mode, done, total = "ab", existing, (existing + remain if remain is not None else None)
                        else:
                            mode, done, total = "wb", 0, remain
                        with part.open(mode) as f:
                            while True:
                                if self.cancel_event.is_set():
                                    raise core.DownloadCancelled("用户取消 Catalog 下载")
                                chunk = response.read(512 * 1024)
                                if not chunk:
                                    break
                                f.write(chunk)
                                done += len(chunk)
                                self.post({"type": "catalog_progress", "done": done, "total": total})
                        os.replace(part, dest)
                    last_error = None
                    break
                except (urllib.error.URLError, urllib.error.HTTPError, TimeoutError, OSError, core.DownloadCancelled) as e:
                    last_error = e
                    if isinstance(e, core.DownloadCancelled):
                        raise
                    if attempt < retries:
                        time.sleep(min(2 ** attempt, 8))
            if last_error:
                raise RuntimeError(f"Catalog 下载失败：{last_error}")
            self.post({"type": "log", "text": "Catalog 下载完成，正在提取并分析。"})
            extracted = ROOT_DIR / "catalog.json"
            core.extract_catalog_from_bundle(dest, extracted)
            rows = core.analyze_catalog(extracted)
            csv_path, xlsx_path = core.write_mapping_files(ROOT_DIR, rows)
            self.post({"type": "rows_loaded", "rows": normalize_rows(rows), "source": f"Catalog / 已更新 {dest.name}"})
            self.post({"type": "log", "text": f"Catalog 更新成功：{len(rows):,} 条；资源表：{csv_path} / {xlsx_path}"})
        self.launch_task("下载并分析 Catalog", work)

    @staticmethod
    def _bundle_key(value):
        return str(value or "").replace("\\", "/").rsplit("/", 1)[-1].strip().casefold()

    def _set_rows(self, rows, source):
        remembered_keys = set()
        if bool(self._config.get("remember_ui_state", True)):
            remembered_keys.update(str(x).casefold() for x in self._config.get("checked_bundles", []) if str(x).strip())
        # Preserve current selections when a newly generated Catalog replaces
        # the in-memory table; bundle keys, not row indices, are stable IDs.
        for idx in getattr(self, "checked_indices", set()):
            if 0 <= idx < len(getattr(self, "rows", [])):
                key = self._bundle_key(self.rows[idx].get("bundle", ""))
                if key:
                    remembered_keys.add(key)
        self.rows = []
        self.checked_indices.clear()
        for i, row in enumerate(rows):
            item = {field: str(row.get(field, "") or "") for field in core.CSV_FIELDS}
            item["_gui_index"] = i
            self.rows.append(item)
            if self._bundle_key(item.get("bundle", "")) in remembered_keys:
                self.checked_indices.add(i)
        self.last_source = source
        self.source_var.set(f"数据源：{source}")
        source_path = Path(source).expanduser()
        if source_path.is_file() and source_path.suffix.casefold() in {".csv", ".xlsx", ".json", ".txt"}:
            self._config["resource_file"] = str(source_path.resolve())
        elif str(source).startswith("Catalog /"):
            generated_csv = ROOT_DIR / "bundle_id_mapping.csv"
            if generated_csv.is_file():
                self._config["resource_file"] = str(generated_csv.resolve())
        categories = sorted({r.get("category", "").strip() for r in self.rows if r.get("category", "").strip()}, key=str.casefold)
        category_values = ["全部分类", *categories]
        self.category_combo.configure(values=category_values)
        if self.category_var.get() not in category_values:
            self.category_var.set("全部分类")
        versions = sorted({r.get("version", "").strip() for r in self.rows if r.get("version", "").strip()}, key=str.casefold)
        version_values = ["全部版本", "(无版本)", *versions]
        self.version_combo.configure(values=version_values)
        if self.version_var.get() not in version_values:
            self.version_var.set("全部版本")
        self.save_config(silent=True)
        self.loaded_count = len(self.rows)
        self._refresh_local_status()
        self.apply_filter()
        self.log(f"当前资源库：{len(self.rows):,} 条；来源：{source}")
        self.status_var.set(f"已载入 {len(self.rows):,} 条资源")

    def _refresh_local_status(self):
        try:
            out_dir = self._path_setting("output_dir")
            logs_dir = self._path_setting("logs_dir")
            index = core.load_historical_bundle_index(out_dir, logs_dir)
            self.local_index = index
        except Exception:
            self.local_index = {}

    def _schedule_filter(self, *_):
        if self._filter_job is not None:
            try:
                self.root.after_cancel(self._filter_job)
            except tk.TclError:
                pass
        self._filter_job = self.root.after(180, self.apply_filter)

    def _row_quality(self, row):
        # Explicit catalog version wins over folder-name hints. This prevents a
        # stale/mixed asset_root path from overriding a LowQuality version tag.
        version_quality = core._quality_token(str(row.get("version", "") or ""), "")
        if version_quality:
            return version_quality
        root_quality = core._quality_token("", str(row.get("asset_root", "") or ""))
        if root_quality:
            return root_quality
        return core._quality_token("", str(row.get("bundle", "") or ""))

    def _resource_sort_key(self, idx, column):
        row = self.rows[idx]
        if column == "check":
            return (1 if idx in self.checked_indices else 0, idx)
        if column == "local":
            bundle = self._bundle_key(row.get("bundle", ""))
            return (1 if bundle in getattr(self, "local_index", {}) else 0, bundle)
        if column == "dependency":
            try:
                return (0, int(row.get("max_dependency_bundle_count", "") or len(core._parse_dependency_bundles(row))))
            except (TypeError, ValueError):
                return (0, len(core._parse_dependency_bundles(row)))
        value = row.get(column, "")
        if column in {"id", "asset_count"}:
            text_value = str(value or "").strip()
            if text_value.isdigit():
                return (0, int(text_value), text_value.casefold())
            if column == "id":
                numeric = re.search(r"\d+", text_value)
                if numeric:
                    return (0, int(numeric.group(0)), text_value.casefold())
            return (1, text_value.casefold(), text_value.casefold())
        return (str(value or "").casefold())

    def _update_sort_heading(self):
        if not hasattr(self, "resource_tree"):
            return
        for col, label in self.resource_heading_labels.items():
            arrow = ""
            if col == self.sort_column:
                arrow = " ↓" if self.sort_descending else " ↑"
            self.resource_tree.heading(col, text=label + arrow, command=lambda c=col: self.sort_resource_rows(c))

    def sort_resource_rows(self, column):
        if column == self.sort_column:
            self.sort_descending = not self.sort_descending
        else:
            self.sort_column = column
            self.sort_descending = False
        self._update_sort_heading()
        self.apply_filter()

    def apply_filter(self):
        self._filter_job = None
        query = self.search_var.get().strip()
        category = self.category_var.get()
        version = self.version_var.get()
        groups = core.parse_query_groups(query) if query else []
        keep = []
        for i, row in enumerate(self.rows):
            if category and category != "全部分类" and row.get("category", "") != category:
                continue
            row_version = str(row.get("version", "") or "").strip()
            if version == "(无版本)" and row_version:
                continue
            if version and version not in {"全部版本", "(无版本)"} and row_version != version:
                continue
            if groups and not any(all(core.token_matches_row(row, token) for token in group) for group in groups):
                continue
            keep.append(i)
        if self.sort_column not in self.resource_heading_labels:
            self.sort_column = "id"
        keep.sort(key=lambda idx: self._resource_sort_key(idx, self.sort_column), reverse=self.sort_descending)
        self.filtered_indices = keep
        children = self.resource_tree.get_children()
        if children:
            self.resource_tree.delete(*children)
        self.item_to_index.clear()
        for idx in keep:
            row = self.rows[idx]
            iid = f"r{idx}"
            self.item_to_index[iid] = idx
            checked = "☑" if idx in self.checked_indices else "☐"
            bundle = str(row.get("bundle", "")).replace("\\", "/").rsplit("/", 1)[-1]
            status = "已下载" if self._bundle_key(bundle) in getattr(self, "local_index", {}) else "未下载"
            try:
                dep_count = len(core._parse_dependency_bundles(row))
            except Exception:
                dep_count = row.get("max_dependency_bundle_count", "")
            vals = (checked, row.get("id", ""), row.get("category", ""), row.get("version", ""), bundle, row.get("asset_root", ""), row.get("asset_count", ""), dep_count, status)
            self.resource_tree.insert("", "end", iid=iid, values=vals)
        self.count_var.set(f"显示 {len(keep):,} / {len(self.rows):,} 条 · 已勾选 {len(self.checked_indices):,} 条")
        self._update_sort_heading()
        self.save_config(silent=True)

    def _resource_tree_click(self, event):
        region = self.resource_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        column = self.resource_tree.identify_column(event.x)
        item = self.resource_tree.identify_row(event.y)
        if column == "#1" and item in self.item_to_index:
            idx = self.item_to_index[item]
            if idx in self.checked_indices:
                self.checked_indices.remove(idx)
            else:
                self.checked_indices.add(idx)
            self.resource_tree.set(item, "check", "☑" if idx in self.checked_indices else "☐")
            self.count_var.set(f"显示 {len(self.filtered_indices):,} / {len(self.rows):,} 条 · 已勾选 {len(self.checked_indices):,} 条")
            self.save_config(silent=True)
            return "break"

    def _resource_tree_double_click(self, event):
        item = self.resource_tree.identify_row(event.y)
        if item in self.item_to_index:
            idx = self.item_to_index[item]
            row = self.rows[idx]
            messagebox.showinfo("资源详情", "\n".join(f"{k}: {row.get(k, '')}" for k in core.CSV_FIELDS), parent=self.root)

    def _visible_selection(self):
        return [self.item_to_index[iid] for iid in self.resource_tree.selection() if iid in self.item_to_index]

    def check_selected(self):
        self.checked_indices.update(self._visible_selection())
        self.apply_filter()

    def uncheck_selected(self):
        self.checked_indices.difference_update(self._visible_selection())
        self.apply_filter()

    def check_all_filtered(self):
        self.checked_indices.update(self.filtered_indices)
        self.apply_filter()

    def invert_filtered(self):
        for idx in self.filtered_indices:
            if idx in self.checked_indices:
                self.checked_indices.remove(idx)
            else:
                self.checked_indices.add(idx)
        self.apply_filter()

    def clear_checks(self):
        self.checked_indices.clear()
        self.apply_filter()

    def _chosen_rows(self):
        indices = sorted(self.checked_indices)
        if not indices:
            indices = self._visible_selection()
        if not indices:
            messagebox.showwarning("未选择资源", "请先勾选资源，或在资源列表中使用 Ctrl/Shift 多选。", parent=self.root)
            return []
        chosen = []
        seen = set()
        for i in indices:
            if 0 <= i < len(self.rows):
                row = dict(self.rows[i])
                key = str(row.get("bundle", "")).casefold()
                if key and key not in seen:
                    seen.add(key)
                    row["_selected_reason"] = "query"
                    chosen.append(row)
        return chosen

    def _skip_low_quality_enabled(self):
        var = getattr(self, "vars", {}).get("skip_low_quality")
        return bool(var.get()) if var is not None else bool(self._config.get("skip_low_quality", True))

    def _filter_download_quality(self, rows):
        if not self._skip_low_quality_enabled():
            return [dict(row) for row in rows], 0
        accepted, skipped = [], 0
        for row in rows:
            if self._row_quality(row) == "LowQuality":
                skipped += 1
            else:
                accepted.append(dict(row))
        return accepted, skipped

    def preflight_selected(self):
        selected = self._chosen_rows()
        if not selected:
            return
        direct, skipped_direct = self._filter_download_quality(selected)
        skip_low_quality = self._skip_low_quality_enabled()
        if skipped_direct:
            self.log(f"LowQuality 策略：预检查中排除 {skipped_direct} 个明确标记为 LowQuality 的主资源；无质量标记资源不会被排除。")
        if not direct:
            messagebox.showwarning("没有可检查的资源", "当前勾选内容均被“跳过 LowQuality”设置排除。未明确标记 LowQuality 的资源仍会正常处理。", parent=self.root)
            return
        def work():
            output_dir = self._path_setting("output_dir")
            logs_dir = self._path_setting("logs_dir")
            output_dir.mkdir(parents=True, exist_ok=True)
            logs_dir.mkdir(parents=True, exist_ok=True)
            all_rows = self.rows_snapshot()
            if skip_low_quality:
                all_rows = [row for row in all_rows if self._row_quality(row) != "LowQuality"]
            if self._bool_setting("expand_dependencies"):
                matches, unresolved = core.expand_with_dependencies(direct, all_rows)
            else:
                matches, unresolved = [dict(r) for r in direct], []
            names = core.build_local_names(matches)
            base_url = self._string_setting("base_url").strip() or core.DEFAULT_BUNDLE_BASE_URL
            if self._bool_setting("force"):
                existing, network = [], list(zip(matches, names))
            else:
                existing, network = core.preflight_existing(matches, names, output_dir, logs_dir, base_url)
            lines = [
                "预检查模式：只分析资源与本地文件，不启动网络下载。",
                f"直接选择：{len(direct):,} 个资源",
                f"依赖展开：{len(matches):,} 个 Bundle",
                f"本地已存在：{len(existing):,} 个",
                f"若开始下载，预计需下载：{len(network):,} 个",
                f"资源表缺失依赖：{len(unresolved):,} 个",
                "", "主资源示例：",
            ]
            lines.extend("  · " + str(r.get("bundle", "")) for r in direct[:25])
            if len(direct) > 25:
                lines.append(f"  ……另有 {len(direct)-25:,} 个主资源未显示")
            if unresolved:
                lines.extend(["", "未在资源表中找到的依赖示例："])
                lines.extend("  · " + x for x in unresolved[:20])
                if len(unresolved) > 20:
                    lines.append(f"  ……另有 {len(unresolved)-20:,} 个未显示")
            report = "\n".join(lines)
            self.post({"type": "info", "title": "预检查结果", "text": f"预检查完成：主资源 {len(direct)}，依赖展开 {len(matches)}，本地存在 {len(existing)}，待下载 {len(network)}。", "dialog": report})
        self.notebook.select(self.tasks_tab)
        self.launch_task("资源预检查（不下载）", work)

    def download_selected(self, auto_export=False):
        selected = self._chosen_rows()
        if not selected:
            return
        direct, skipped_direct = self._filter_download_quality(selected)
        skip_low_quality = self._skip_low_quality_enabled()
        if skipped_direct:
            self.log(f"LowQuality 策略：本次任务排除 {skipped_direct} 个明确标记为 LowQuality 的主资源；未标记质量的资源仍会下载。")
        if not direct:
            messagebox.showwarning("没有可下载的资源", "当前勾选内容均被“跳过 LowQuality”设置排除。取消该设置后可下载这些版本。", parent=self.root)
            return
        if auto_export:
            ok, problems = core.export_dependency_status()
            install_deps = False
            if not ok:
                ask = messagebox.askyesno("缺少导出依赖", "检测到导出依赖未就绪：\n" + "\n".join(problems) + "\n\n是否在本次任务中自动安装 UnityPy / Pillow？", parent=self.root)
                if not ask:
                    proceed = messagebox.askyesno("继续下载", "只下载、不导出，继续执行吗？", parent=self.root)
                    if not proceed:
                        return
                    auto_export = False
                else:
                    install_deps = True
            else:
                install_deps = False
        else:
            install_deps = False
        self.notebook.select(self.tasks_tab)
        def work():
            session, output_dir, export_root, logs_root = self._new_session()
            if install_deps:
                self._install_deps_worker(session)
            all_rows = self.rows_snapshot()
            if skip_low_quality:
                all_rows = [row for row in all_rows if self._row_quality(row) != "LowQuality"]
            if self._bool_setting("expand_dependencies"):
                matches, unresolved = core.expand_with_dependencies(direct, all_rows)
            else:
                matches, unresolved = [dict(r) for r in direct], []
            for row in matches:
                row.setdefault("_selected_reason", "query")
            names = core.build_local_names(matches)
            base_url = self._string_setting("base_url").strip() or core.DEFAULT_BUNDLE_BASE_URL
            force = self._bool_setting("force")
            workers = self._numeric_setting("workers", 20)
            timeout = self._numeric_setting("timeout", 60)
            retries = self._numeric_setting("retries", 3)
            if force:
                existing_records, need_network = [], list(zip([dict(r) for r in matches], names))
            else:
                existing_records, need_network = core.preflight_existing(matches, names, output_dir, logs_root, base_url)
            self.post({"type": "summary", "text": f"资源 {len(direct):,} · 含依赖 {len(matches):,} · 已存在 {len(existing_records):,} · 待下载 {len(need_network):,}"})
            self.post({"type": "log", "text": f"匹配 {len(direct):,} 个主资源，依赖展开后 {len(matches):,} 个 Bundle；待下载 {len(need_network):,} 个。"})
            if unresolved:
                self.post({"type": "warning", "text": f"资源表中找不到 {len(unresolved)} 个依赖 Bundle：" + ", ".join(unresolved[:20]) + (" …" if len(unresolved) > 20 else "" )})
            round_no = session.next_round()
            records = list(existing_records)
            if need_network:
                progress = GuiDownloadProgress(self, workers)
                with ThreadPoolExecutor(max_workers=max(1, workers), thread_name_prefix="DL") as pool:
                    future_map = {pool.submit(core.download_bundle_job, row, name, output_dir, base_url, progress, session, timeout=timeout, retries=retries, force=force): (row, name) for row, name in need_network}
                    for future in as_completed(future_map):
                        try:
                            records.append(future.result())
                        except Exception as e:
                            row, name = future_map[future]
                            session.log(f"download worker exception: {type(e).__name__}: {e}")
                            records.append({"id": row.get("id", ""), "category": row.get("category", ""), "version": row.get("version", ""), "selected_reason": row.get("_selected_reason", "query"), "original_bundle": row.get("bundle", ""), "local_file": name, "status": f"failed: {e}"})
                        self.post({"type": "summary", "text": f"下载处理 {len(records):,}/{len(matches):,}（新任务 {len(need_network):,}）"})
            records.sort(key=lambda r: str(r.get("local_file", "")).casefold())
            manifest = session.download_manifest(round_no)
            core.write_download_manifest(manifest, records)
            downloaded = sum(r.get("status") == "downloaded" for r in records)
            exists = sum(r.get("status") == "exists" for r in records)
            failed = sum(str(r.get("status", "")).startswith("failed") for r in records)
            cancelled = sum(r.get("status") == "cancelled" for r in records)
            core.cleanup_temp(output_dir=output_dir, work_dir=ROOT_DIR)
            self.post({"type": "log", "text": f"下载阶段结束：新下载 {downloaded}，本地跳过 {exists}，失败 {failed}，取消 {cancelled}；清单：{manifest}"})
            self.post({"type": "summary", "text": f"新下载 {downloaded} · 已存在 {exists} · 失败 {failed} · 取消 {cancelled}"})
            if self.cancel_event.is_set():
                self.post({"type": "warning", "text": "下载已取消。已完成文件保留，未完成的非空 .part 文件保留用于续传。"})
                return
            if auto_export:
                okdeps, missing, uncertain = core.verify_downloaded_dependencies(matches, records)
                if not okdeps:
                    self.post({"type": "warning", "text": f"仍有 {len(missing)} 个依赖 Bundle 未下载，继续尝试导出；详细情况见日志。"})
                if uncertain:
                    self.post({"type": "warning", "text": f"有 {uncertain} 条资源记录无法完整确认依赖数量。"})
                self._do_export(direct, all_rows, records, output_dir, export_root, session, round_no)
            self._refresh_local_status_threadsafe()
        self.launch_task("下载资源" + ("并导出" if auto_export else ""), work)

    def export_selected(self):
        direct = self._chosen_rows()
        if not direct:
            return
        ok, problems = core.export_dependency_status()
        install_deps = False
        if not ok:
            if messagebox.askyesno("缺少导出依赖", "检测到导出依赖未就绪：\n" + "\n".join(problems) + "\n\n是否自动安装？", parent=self.root):
                install_deps = True
            else:
                return
        self.notebook.select(self.tasks_tab)
        def work():
            session, output_dir, export_root, logs_root = self._new_session()
            if install_deps:
                self._install_deps_worker(session)
            all_rows = self.rows_snapshot()
            if self._bool_setting("expand_dependencies"):
                matches, unresolved = core.expand_with_dependencies(direct, all_rows)
            else:
                matches, unresolved = [dict(r) for r in direct], []
            names = core.build_local_names(matches)
            base_url = self._string_setting("base_url").strip() or core.DEFAULT_BUNDLE_BASE_URL
            existing, missing = core.preflight_existing(matches, names, output_dir, logs_root, base_url)
            if missing:
                missing_names = [str(r.get("bundle", "")) for r, _ in missing]
                self.post({"type": "warning", "text": f"有 {len(missing)} 个 Bundle 尚未下载，无法完整导出：" + ", ".join(missing_names[:20])})
                raise RuntimeError(f"导出前检查失败：有 {len(missing)} 个 Bundle 尚未下载。请先下载，或在下载后选择自动导出。")
            round_no = session.next_round()
            self._do_export(direct, all_rows, existing, output_dir, export_root, session, round_no)
        self.launch_task("仅导出选中资源", work)

    def _do_export(self, direct, all_rows, records, output_dir, export_root, session, round_no):
        layout = self._string_setting("export_layout")
        workers = self._numeric_setting("export_workers", 20)
        self.post({"type": "log", "text": f"开始 Export：主资源 {len(direct):,}，worker {workers}，布局 {layout}。"})
        ok, failed, manifest, fail_log = core.export_parallel(direct, all_rows, records, output_dir, export_root, layout, workers, session, round_no)
        core.cleanup_temp(output_dir=output_dir, work_dir=ROOT_DIR)
        self.post({"type": "summary", "text": f"Export 完成 · 成功项 {ok:,} · 失败项 / worker {failed:,}"})
        self.post({"type": "log", "text": f"Export 完成：成功项 {ok:,}，失败项 / worker {failed:,}；目录：{export_root}；清单：{manifest}"})
        if fail_log.exists():
            self.post({"type": "warning", "text": f"部分 Export 失败，详细信息：{fail_log}"})
        self._run_auto_spine_actions_after_export(export_root)

    def _new_session(self):
        output_dir = self._path_setting("output_dir")
        export_root = self._path_setting("export_dir")
        logs_root = self._path_setting("logs_dir")
        output_dir.mkdir(parents=True, exist_ok=True)
        export_root.mkdir(parents=True, exist_ok=True)
        logs_root.mkdir(parents=True, exist_ok=True)
        session = core.SessionLog(logs_root)
        session.log(f"GUI start V{core.VERSION}; python={sys.version}")
        try:
            core.migrate_legacy_logs(output_dir, export_root, session)
        except Exception as e:
            session.log(f"legacy log migration skipped: {e}")
        return session, output_dir, export_root, logs_root

    def _install_deps_worker(self, session):
        self.post({"type": "log", "text": "开始安装导出依赖：" + " ".join(core.EXPORT_PIP_REQUIREMENTS)})
        proc = subprocess.run([sys.executable, "-m", "pip", "install", *core.EXPORT_PIP_REQUIREMENTS], cwd=str(ROOT_DIR), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, errors="replace")
        for line in (proc.stdout or "").splitlines()[-100:]:
            self.post({"type": "log", "text": line})
        session.log("pip install exit=" + str(proc.returncode))
        if proc.returncode != 0:
            raise RuntimeError("安装 UnityPy / Pillow 失败。请检查 Python / pip 网络连接；pip 输出已写入界面日志。")
        import importlib
        importlib.invalidate_caches()
        ok, problems = core.export_dependency_status()
        if not ok:
            raise RuntimeError("安装后导出依赖仍不可用：" + ", ".join(problems))
        self.post({"type": "log", "text": "导出依赖安装完成。"})

    def repair_existing_export_classification(self):
        rows = self.rows_snapshot()
        if not rows:
            messagebox.showwarning("需要资源表", "请先载入资源表，再根据资源 ID / 版本信息修复历史导出分类。", parent=self.root)
            return
        self.save_config(silent=True)
        export_root = self._path_setting("export_dir")
        logs_root = self._path_setting("logs_dir")
        manifests = []
        if logs_root.exists():
            manifests.extend(p for p in logs_root.rglob("*export_manifest.csv") if p.is_file())
        # When repairing an older project's exported_assets from a new GUI
        # folder, also discover the old project's sibling logs directory.
        adjacent_logs = export_root.parent / "logs"
        try:
            same_log_root = adjacent_logs.resolve() == logs_root.resolve()
        except OSError:
            same_log_root = adjacent_logs == logs_root
        if not same_log_root and adjacent_logs.is_dir():
            manifests.extend(p for p in adjacent_logs.rglob("*export_manifest.csv") if p.is_file())
        legacy_manifest = export_root / "export_manifest.csv"
        if legacy_manifest.is_file():
            manifests.append(legacy_manifest)
        if not manifests:
            messagebox.showinfo("没有历史清单", "没有找到历史 Export Manifest。为避免误移动未登记文件，本操作不会猜测或删除未登记资源。", parent=self.root)
            return
        layout = self._string_setting("export_layout")
        proceed = messagebox.askyesno(
            "修复历史导出分类",
            f"将依据当前资源表和历史 Export Manifest，把已登记且仍存在的文件按当前布局（{layout}）重新归档，并清理相关空目录。\n\n"
            "未记录在 Manifest 中的文件会保留；遇到不同内容的同名文件会自动改名，不会直接覆盖。建议先备份导出目录。\n\n继续？",
            parent=self.root,
        )
        if not proceed:
            return

        def work():
            session, _output_dir, current_export_root, current_logs_root = self._new_session()
            result = core.repair_existing_exports(
                current_export_root, current_logs_root, rows, layout, session,
            )
            summary = (
                f"历史导出整理完成：移动 {result['moved']} 个文件；删除重复文件 {result['duplicate_removed']} 个；"
                f"无需移动 {result['unchanged']} 个；清理空目录 {result['pruned_dirs']} 个；"
                f"缺少资源表条目 {result['missing_metadata']} 个；歧义路径 {result['ambiguous']} 个；跳过 {result['skipped']} 个。"
            )
            self.post({"type": "summary", "text": summary})
            self.post({"type": "log", "text": summary})
            self.post({"type": "log", "text": f"整理日志：{session.app_log}"})

        self.launch_task("修复历史导出分类", work)

    def install_export_dependencies(self):
        ok, problems = core.export_dependency_status()
        if ok:
            messagebox.showinfo("依赖检查", "UnityPy / Pillow 均已可用。", parent=self.root)
            return
        if not messagebox.askyesno("安装导出依赖", "将通过当前 Python 的 pip 安装：\n" + "\n".join(core.EXPORT_PIP_REQUIREMENTS) + "\n\n继续？", parent=self.root):
            return
        def work():
            session, *_ = self._new_session()
            self._install_deps_worker(session)
        self.launch_task("安装导出依赖", work)

    def rows_snapshot(self):
        return [dict(r) for r in self.rows]

    def _refresh_local_status_threadsafe(self):
        try:
            self.local_index = core.load_historical_bundle_index(self._path_setting("output_dir"), self._path_setting("logs_dir"))
            self.post({"type": "refresh_rows"})
        except Exception:
            pass

    def _numeric_setting(self, key, default):
        try:
            return max(1, int(self._config.get(key, default)))
        except (KeyError, ValueError, TypeError):
            return default

    def _string_setting(self, key):
        return str(self._config.get(key, DEFAULT_CONFIG.get(key, "")))

    def _bool_setting(self, key):
        return bool(self._config.get(key, False))

    def _path_setting(self, key):
        raw = self._string_setting(key).strip() or DEFAULT_CONFIG[key]
        p = Path(raw).expanduser()
        if not p.is_absolute():
            p = ROOT_DIR / p
        return p.resolve()

    def browse_path(self, key, mode):
        current = self._string_setting(key)
        initial = Path(current).expanduser()
        if not initial.is_absolute():
            initial = ROOT_DIR / initial
        if mode == "dir":
            selected = filedialog.askdirectory(parent=self.root, initialdir=str(initial if initial.exists() else ROOT_DIR), title="选择目录")
        else:
            selected = filedialog.askopenfilename(parent=self.root, initialdir=str(initial.parent if initial.parent.exists() else ROOT_DIR), title="选择文件")
        if selected:
            self.vars[key].set(selected)

    def reset_settings(self):
        if not messagebox.askyesno("恢复默认", "将所有设置恢复默认值，确定？", parent=self.root):
            return
        for key, value in DEFAULT_CONFIG.items():
            var = self.vars.get(key)
            if var is not None:
                var.set(value if isinstance(var, tk.BooleanVar) else str(value))
        if hasattr(self, "search_var"):
            self.search_var.set("")
            self.category_var.set("全部分类")
            self.version_var.set("全部版本")
            self.checked_indices.clear()
            self.sort_column, self.sort_descending = "id", False
            self.apply_filter()
        if hasattr(self, "spine_action_tree"):
            for iid in self.spine_action_tree.get_children():
                self.spine_action_tree.delete(iid)
            for action in spine_engine.DEFAULT_ACTIONS:
                self._insert_spine_action(action)
        self.save_config(silent=True)
        self.log("已恢复默认设置。")

    def open_directory(self, key):
        self.save_config(silent=True)
        path = self._path_setting(key)
        path.mkdir(parents=True, exist_ok=True)
        try:
            if os.name == "nt":
                os.startfile(str(path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(path)])
            else:
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as e:
            messagebox.showerror("无法打开目录", f"{path}\n{e}", parent=self.root)

    def _update_task_row(self, event):
        name = str(event.get("name", "未知任务"))
        iid = name
        done = int(event.get("done", 0) or 0)
        total = event.get("total")
        speed = float(event.get("speed", 0) or 0)
        pct = f"{done * 100 / total:.1f}%" if total else "—"
        size = f"{core.human_size(done)} / {core.human_size(total)}" if total else core.human_size(done)
        vals = (name, event.get("status", "等待"), pct, size, f"{core.human_size(int(speed))}/s" if speed > 0 else "—")
        if self.task_tree.exists(iid):
            self.task_tree.item(iid, values=vals)
        else:
            try:
                self.task_tree.insert("", "end", iid=iid, values=vals)
            except tk.TclError:
                iid = f"task_{len(self.download_tree_items)}"
                self.task_tree.insert("", "end", iid=iid, values=vals)
        self.download_tree_items[name] = iid

    def toggle_pause(self):
        if not self.busy or "下载" not in self.current_task:
            return
        if self.resume_event.is_set():
            self.resume_event.clear()
            self.pause_btn.configure(text="继续下载")
            self.status_var.set("已暂停：下载线程将在当前分块边界停住")
            self.log("用户暂停下载。")
        else:
            self.resume_event.set()
            self.pause_btn.configure(text="暂停下载")
            self.status_var.set("已继续下载")
            self.log("用户继续下载。")

    def cancel_current(self):
        if self.busy and "下载" in self.current_task:
            if messagebox.askyesno("取消下载", "取消当前下载任务？已经完成的文件会保留，未完成的 .part 文件可用于续传。", parent=self.root):
                self.cancel_event.set()
                self.resume_event.set()
                self.status_var.set("正在取消；等待活动网络分块退出…")
                self.log("已请求取消下载。")

    def on_close(self):
        if self.busy:
            if not messagebox.askyesno("正在运行任务", "当前任务仍在执行。退出会请求取消下载并关闭界面，确定退出？", parent=self.root):
                return
            self.cancel_event.set()
            self.resume_event.set()
        self.save_config(silent=True)
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        app = SmartBundleApp(root)
    except Exception as e:
        messagebox.showerror("启动失败", f"程序初始化失败：{type(e).__name__}: {e}")
        raise
    root.mainloop()


if __name__ == "__main__":
    main()
