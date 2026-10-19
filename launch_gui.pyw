# -*- coding: utf-8 -*-
"""Windowed entry point that reports otherwise-silent startup/runtime errors."""
from __future__ import annotations

import runpy
import sys
import traceback
from pathlib import Path

ROOT = Path(__file__).resolve().parent
LOG_FILE = ROOT / "logs" / "gui_error.log"
_POPUP = {"shown": False}


def _write_log(title: str, details: str) -> None:
    try:
        LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
        with LOG_FILE.open("a", encoding="utf-8") as stream:
            stream.write("\n" + "=" * 72 + "\n")
            stream.write(title + "\n")
            stream.write(details.rstrip() + "\n")
    except Exception:
        pass


try:
    import tkinter
    from tkinter import messagebox

    _original_showerror = messagebox.showerror

    def _tracked_showerror(*args, **kwargs):
        _POPUP["shown"] = True
        return _original_showerror(*args, **kwargs)

    messagebox.showerror = _tracked_showerror

    def _report_callback_exception(self, exc, value, tb):
        details = "".join(traceback.format_exception(exc, value, tb))
        _write_log("Unhandled GUI callback exception", details)
        try:
            messagebox.showerror(
                "GUI 运行错误",
                "程序运行时发生未处理异常。详细信息已保存到：\n"
                + str(LOG_FILE)
                + "\n\n"
                + details[-3500:],
                parent=self,
            )
        except Exception:
            pass

    tkinter.Tk.report_callback_exception = _report_callback_exception
except Exception:
    messagebox = None


try:
    runpy.run_path(str(ROOT / "app.py"), run_name="__main__")
except Exception as exc:
    details = traceback.format_exc()
    _write_log("GUI startup failure", details)
    if not _POPUP["shown"] and messagebox is not None:
        try:
            messagebox.showerror(
                "Smart Bundle Downloader 启动失败",
                f"{type(exc).__name__}: {exc}\n\n"
                f"详细错误日志：\n{LOG_FILE}\n\n"
                "请把该日志发给开发者以便定位问题。",
            )
        except Exception:
            pass
    sys.exit(1)
