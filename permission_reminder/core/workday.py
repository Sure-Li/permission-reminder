# -*- coding: utf-8 -*-
"""工作日计算模块。

职责：计算"到期日前推 N 个工作日"的触发日期。

策略（已确认）：
- 集成 chinese_calendar 排除周末 + 法定节假日；
- 超出 chinese_calendar 数据年份时优雅降级为仅排周末，并记录警告日志
  （预留 WorkdayCalendar 抽象接口，节假日策略可替换/扩展）。
"""
from __future__ import annotations

import logging
from abc import ABC, abstractmethod
from datetime import date, timedelta
from pathlib import Path
from typing import Dict, Set

logger = logging.getLogger(__name__)


class WorkdayCalendar(ABC):
    """工作日日历抽象接口（扩展点：可注入自定义节假日列表）。"""

    @abstractmethod
    def is_workday(self, d: date) -> bool:
        """判断某日是否为工作日。"""

    def minus_workdays(self, d: date, n: int) -> date:
        """从日期 d 往前推 n 个工作日，返回触发日期。

        例如 d=2026-10-21（周三），n=5 -> 依次回退 10-20、10-19、10-16、
        10-15、10-14，返回 2026-10-14。
        """
        result = d
        remaining = n
        while remaining > 0:
            result -= timedelta(days=1)
            if self.is_workday(result):
                remaining -= 1
        return result


class WeekendOnlyCalendar(WorkdayCalendar):
    """仅排除周末的工作日日历（降级方案 / 扩展基线）。"""

    def __init__(self, extra_holidays: Set[date] | None = None) -> None:
        self._extra_holidays: Set[date] = extra_holidays or set()

    def is_workday(self, d: date) -> bool:
        if d.weekday() >= 5:  # 周六=5 周日=6
            return False
        return d not in self._extra_holidays


class ChineseWorkdayCalendar(WorkdayCalendar):
    """基于 chinese_calendar 的中国工作日日历（周末 + 法定节假日）。

    数据优先级：
    1. overrides —— 本地同步的节假日数据（data/holidays.json，可随官方公告更新）；
    2. chinese_calendar 包内置数据；
    3. 仅排除周末（包无该年份数据时的降级方案）。
    """

    def __init__(self, overrides: Dict[date, bool] | None = None) -> None:
        try:
            import chinese_calendar as cc  # 延迟导入，缺失时可降级
            self._cc = cc
        except ImportError:  # pragma: no cover —— 极端情况：包被卸载
            logger.warning("chinese_calendar 未安装，降级为仅排除周末")
            self._cc = None
        self._overrides: Dict[date, bool] = dict(overrides or {})
        if self._overrides:
            logger.info("已加载 %d 条本地同步节假日数据（优先于包内置数据）",
                        len(self._overrides))
        self._fallback = WeekendOnlyCalendar()
        self._fallback_years: Set[int] = set()

    def _year_supported(self, year: int) -> bool:
        """检查某年份是否在 chinese_calendar 数据范围内（结果缓存）。"""
        if self._cc is None or year in self._fallback_years:
            return False
        probe = date(year, 1, 2)  # 探测日（避开元旦）
        try:
            self._cc.is_workday(probe)
            return True
        except NotImplementedError:
            logger.warning(
                "chinese_calendar 无 %d 年节假日数据，该年份降级为仅排除周末"
                "（升级 chinese-calendar 包可获得新年度数据）", year,
            )
            self._fallback_years.add(year)
            return False

    def is_workday(self, d: date) -> bool:
        # 1. 本地同步数据优先（含调休补班日、法定假日）
        if d in self._overrides:
            return self._overrides[d]
        # 2. chinese_calendar 包内置数据
        if self._year_supported(d.year):
            try:
                return bool(self._cc.is_workday(d))
            except NotImplementedError:  # 双保险：跨年边界
                return self._fallback.is_workday(d)
        # 3. 降级：仅排除周末
        return self._fallback.is_workday(d)


def create_calendar(store_path: Path | None = None) -> WorkdayCalendar:
    """工厂：返回中国工作日日历（加载本地同步节假日数据，自动降级）。

    Args:
        store_path: 节假日本地数据文件路径；None 时用默认 data/holidays.json。
    """
    # 延迟导入避免循环依赖
    from .holiday_sync import HolidayStore
    from ..infra.config_store import DATA_DIR

    path = store_path if store_path is not None else DATA_DIR / "holidays.json"
    try:
        overrides = HolidayStore(path).workday_overrides()
    except Exception as exc:  # noqa: BLE001 —— 数据异常不影响主流程
        logger.warning("加载本地节假日数据失败，改用包内置数据：%s", exc)
        overrides = {}
    return ChineseWorkdayCalendar(overrides=overrides)
