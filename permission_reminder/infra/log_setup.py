# -*- coding: utf-8 -*-
"""日志模块。

职责：
1. 全局日志初始化：控制台 + 滚动文件（logs/app.log，按大小轮转）；
2. GuiLogHandler：把日志消息桥接到 GUI 日志窗口（通过回调注入，线程安全由
   UI 层的 queue + after 轮询保证——handler 只负责调用回调）。
"""
from __future__ import annotations

import logging
import threading
from logging.handlers import RotatingFileHandler
from typing import Callable, Optional

from .config_store import LOG_DIR

LOG_FILE = LOG_DIR / "app.log"

# 日志格式：时间 级别 [模块名] 消息
_FMT = "%(asctime)s %(levelname)-7s [%(name)s] %(message)s"
_DATEFMT = "%Y-%m-%d %H:%M:%S"


class GuiLogHandler(logging.Handler):
    """将日志记录转发到 GUI 日志窗口的回调。

    回调由 UI 层注入；本 handler 只做转发，不持有任何 tkinter 对象，
    线程安全性由 UI 侧的 queue.Queue + root.after() 轮询机制保证。
    """

    def __init__(self, callback: Callable[[str], None]) -> None:
        super().__init__()
        self._callback = callback
        self._lock_cb = threading.Lock()
        self.setFormatter(logging.Formatter(_FMT, _DATEFMT))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            msg = self.format(record)
            with self._lock_cb:
                self._callback(msg)
        except Exception:  # noqa: BLE001 —— 日志桥绝不向上抛异常
            self.handleError(record)


def setup_logging(gui_callback: Optional[Callable[[str], None]] = None) -> None:
    """初始化全局日志：文件（始终）+ 控制台（源码运行时）+ GUI 桥（可选）。

    Args:
        gui_callback: GUI 日志窗口的接收函数；打包版无控制台时仍写文件。
    """
    LOG_DIR.mkdir(parents=True, exist_ok=True)

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    # 防止重复初始化（例如测试中多次调用）
    for h in list(root.handlers):
        root.removeHandler(h)

    fmt = logging.Formatter(_FMT, _DATEFMT)

    file_handler = RotatingFileHandler(
        LOG_FILE, maxBytes=2 * 1024 * 1024, backupCount=5, encoding="utf-8"
    )
    file_handler.setFormatter(fmt)
    root.addHandler(file_handler)

    console = logging.StreamHandler()
    console.setFormatter(fmt)
    root.addHandler(console)

    if gui_callback is not None:
        root.addHandler(GuiLogHandler(gui_callback))
