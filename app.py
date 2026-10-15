# -*- coding: utf-8 -*-
"""权限到期提醒工具 —— 程序入口。

用法（源码运行）：
    python app.py
"""
from __future__ import annotations

import sys

# Windows 控制台默认 GBK，重定向输出时避免中文日志触发 UnicodeEncodeError
for _stream in (sys.stdout, sys.stderr):
    if _stream is not None and hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except (OSError, ValueError):
            pass

from permission_reminder.ui.main_window import MainWindow  # noqa: E402


def main() -> None:
    """启动主窗口。"""
    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
