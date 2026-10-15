# -*- coding: utf-8 -*-
"""定时调度模块。

职责：后台 daemon 线程，周期轮询系统时间，在每天配置的时:分 触发一次检查。

设计要点（已确认决策）：
- 纯 threading 实现（不用 schedule/apscheduler），30 秒轮询；
- 触发条件：调度开启 且 now >= 今日 HH:MM 且 今日尚未执行过定时检查；
- "今日已执行"状态仅存内存——程序重启后由发送记录（SentRecordStore）
  天然去重兜底，不会重复发送；
- 实际检查动作通过回调注入（run_check），回调内部自行处理异常，
  调度线程永不因业务异常退出。
"""
from __future__ import annotations

import logging
import threading
from datetime import date, datetime
from typing import Callable

logger = logging.getLogger(__name__)

POLL_INTERVAL_SECONDS = 30  # 轮询间隔


class SchedulerThread(threading.Thread):
    """每日定时调度线程。

    Args:
        get_config: 读取当前调度配置（返回 ScheduleConfig，允许运行中修改）。
        run_check: 触发时执行的检查函数（在工作线程中调用）。
    """

    def __init__(
        self,
        get_config: Callable,          # () -> ScheduleConfig
        run_check: Callable[[], None], # 无参检查回调
    ) -> None:
        super().__init__(daemon=True, name="scheduler")
        self._get_config = get_config
        self._run_check = run_check
        self._stop_event = threading.Event()
        self._last_run_date: date | None = None  # 今日是否已执行定时检查

    # -------------------------------------------------------------- 对外接口 --

    def stop(self) -> None:
        """请求线程退出（程序关闭时调用）。"""
        self._stop_event.set()

    # -------------------------------------------------------------- 线程主体 --

    def run(self) -> None:
        logger.info("调度线程已启动（每 %d 秒轮询一次）", POLL_INTERVAL_SECONDS)
        while not self._stop_event.wait(POLL_INTERVAL_SECONDS):
            try:
                self._tick()
            except Exception:  # noqa: BLE001 —— 调度线程兜底，绝不退出
                logger.exception("调度轮询出现未预期异常，已忽略并继续")
        logger.info("调度线程已退出")

    def _tick(self) -> None:
        """单次轮询逻辑。"""
        cfg = self._get_config()
        if not cfg.enabled:
            return
        now = datetime.now()
        # 今日已执行过定时检查（重启后 _last_run_date 为 None，但发送记录兜底）
        if self._last_run_date == now.date():
            return
        # 未到今日配置的触发时刻
        if (now.hour, now.minute) < (cfg.hour, cfg.minute):
            return
        logger.info("到达定时触发时刻 %02d:%02d，开始执行检查", cfg.hour, cfg.minute)
        self._last_run_date = now.date()
        self._run_check()  # 异常由回调内部处理，此处再兜一层
