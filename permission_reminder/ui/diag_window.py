# -*- coding: utf-8 -*-
"""SMTP 连接诊断报告窗口。

独立的 Toplevel 弹窗，等宽字体展示诊断报告；若诊断给出了推荐 SMTP 配置，
额外提供一个按钮把推荐值一键套用到主界面的邮箱配置区。
"""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText
from typing import Callable, Optional, Tuple


def show_diag_report(parent: tk.Misc, report: str,
                     recommended: Optional[Tuple[str, int, str]] = None,
                     on_apply: Optional[Callable[[str, int, str], None]] = None) -> None:
    """弹出发信链路诊断报告窗口（仅主线程调用）。

    Args:
        parent: 父窗口。
        report: format_report() 生成的纯文本报告。
        recommended: 推荐的 (服务器, 端口, 加密方式)，无则 None。
        on_apply: "应用推荐配置"按钮的回调。
    """
    win = tk.Toplevel(parent)
    win.title("SMTP 连接诊断报告")
    win.geometry("720x560")
    win.transient(parent)

    frame = ttk.Frame(win, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text="发信链路诊断报告（只读，可复制）",
              font=("Microsoft YaHei", 10, "bold")).pack(anchor=tk.W, pady=(0, 6))

    text = ScrolledText(frame, wrap=tk.NONE, font=("Consolas", 9),
                        height=24, width=92)
    text.pack(fill=tk.BOTH, expand=True)
    text.insert("1.0", report)
    text.config(state=tk.DISABLED)

    btns = ttk.Frame(frame)
    btns.pack(fill=tk.X, pady=(8, 0))

    def do_copy() -> None:
        win.clipboard_clear()
        win.clipboard_append(report)
        copy_btn.config(text="已复制 ✓")
        win.after(1500, lambda: copy_btn.config(text="复制报告"))

    copy_btn = ttk.Button(btns, text="复制报告", command=do_copy)
    copy_btn.pack(side=tk.LEFT, padx=3)

    if recommended is not None and on_apply is not None:
        srv, prt, mode = recommended

        def do_apply() -> None:
            on_apply(srv, prt, mode)
            win.destroy()

        ttk.Button(btns,
                   text=f"应用推荐配置（{srv}:{prt} {mode}）",
                   command=do_apply).pack(side=tk.LEFT, padx=3)

    ttk.Button(btns, text="关闭", command=win.destroy).pack(side=tk.RIGHT, padx=3)
