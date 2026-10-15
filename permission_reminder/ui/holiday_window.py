# -*- coding: utf-8 -*-
"""休假计划展示窗口。

独立的 Toplevel 弹窗，用 Treeview 展示"前后一个月"窗口内的
法定节假日与调休补班安排。
"""
from __future__ import annotations

import tkinter as tk
from datetime import date
from tkinter import ttk
from typing import List

from ..core.holiday_sync import SpecialDay

_WEEKDAY_CN = ["一", "二", "三", "四", "五", "六", "日"]


def show_holiday_plan(parent: tk.Misc, days: List[SpecialDay],
                      title_suffix: str, note: str) -> None:
    """弹出休假计划窗口（仅主线程调用）。

    Args:
        parent: 父窗口。
        days: 区间内特殊日期列表（已升序）。
        title_suffix: 标题后缀，如 "2026-09 ~ 2026-11"。
        note: 底部说明（数据来源 / 同步状态）。
    """
    win = tk.Toplevel(parent)
    win.title(f"休假计划 - {title_suffix}")
    win.geometry("520x460")
    win.transient(parent)  # 跟随父窗口

    frame = ttk.Frame(win, padding=10)
    frame.pack(fill=tk.BOTH, expand=True)

    ttk.Label(frame, text=f"法定节假日 / 调休补班安排（{title_suffix}）",
              font=("Microsoft YaHei", 10, "bold")).pack(anchor=tk.W, pady=(0, 6))

    # 表格：日期 | 星期 | 名称 | 类型
    columns = ("date", "weekday", "name", "kind")
    tree = ttk.Treeview(frame, columns=columns, show="headings", height=14)
    tree.heading("date", text="日期")
    tree.heading("weekday", text="星期")
    tree.heading("name", text="节日名称")
    tree.heading("kind", text="类型")
    tree.column("date", width=120, anchor=tk.CENTER)
    tree.column("weekday", width=70, anchor=tk.CENTER)
    tree.column("name", width=180, anchor=tk.CENTER)
    tree.column("kind", width=80, anchor=tk.CENTER)
    tree.tag_configure("off", foreground="#1a7a34")     # 假日：绿色
    tree.tag_configure("work", foreground="#c0392b")    # 补班：红色
    scroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
    tree.configure(yscrollcommand=scroll.set)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scroll.pack(side=tk.RIGHT, fill=tk.Y)

    today = date.today()
    for sd in days:
        tags = ("off",) if sd.is_off else ("work",)
        marker = "（今天）" if sd.date == today else ""
        tree.insert("", tk.END, values=(
            sd.date.isoformat() + marker,
            f"周{_WEEKDAY_CN[sd.date.weekday()]}",
            sd.name,
            "休" if sd.is_off else "调休上班",
        ), tags=tags)

    if not days:
        tree.insert("", tk.END, values=("（该区间暂无节假日/补班数据）", "", "", ""))

    ttk.Label(win, text=note, foreground="gray", wraplength=490,
              justify=tk.LEFT).pack(anchor=tk.W, padx=12, pady=(0, 4))
    ttk.Button(win, text="关闭", command=win.destroy).pack(pady=(0, 10))
