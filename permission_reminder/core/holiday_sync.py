# -*- coding: utf-8 -*-
"""节假日同步模块。

职责：
1. 从在线数据源拉取指定年份的法定节假日安排（休假日 + 调休补班日）；
2. 本地持久化 data/holidays.json，作为工作日计算的**覆盖数据源**
   （优先级：本地同步数据 > chinese_calendar 包内置 > 仅排除周末）；
3. 提供日期区间内的特殊日期查询（供休假计划窗口展示）。

同步入口有两个（都走 sync_window_years，共用一把锁）：
- 手动：「同步最新节假日」按钮；
- 自动：每次检查（手动 / 定时 / 启动补检）前默认先同步一次，
  失败仅记日志、不打断检查（受配置项 auto_sync_holidays 控制）。

数据源：
- 主源：timor.tech 节假日 API（需带浏览器 User-Agent，否则 403）；
- 备源：holiday-cn（GitHub Pages，部分网络环境不可达，失败自动跳过）。
"""
from __future__ import annotations

import json
import logging
import re
import threading
import urllib.request
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, List, Tuple

logger = logging.getLogger(__name__)

# 数据源（按顺序尝试，全部失败才算同步失败）
HOLIDAY_SOURCES = [
    "https://timor.tech/api/holiday/year/{year}",
    "https://natescarlet.github.io/holiday-cn/{year}.json",
]
_HTTP_TIMEOUT = 15
_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) PermReminder/1.1"

_DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")


class HolidaySyncError(Exception):
    """节假日同步失败（网络不可达 / 数据格式异常）。"""


@dataclass
class SpecialDay:
    """一个特殊日期（法定节假日或调休补班日）。"""
    date: date
    name: str       # 节日名称 / "调休补班"
    is_off: bool    # True=休假日，False=调休补班（上班）

    @property
    def kind_label(self) -> str:
        return "休" if self.is_off else "补班"


# ----------------------------------------------------------------- 日期区间 --

def month_window_around(today: date) -> Tuple[date, date]:
    """计算"前后一个月"窗口：上个月 1 号 ~ 下个月最后一天。"""
    if today.month == 1:
        start = date(today.year - 1, 12, 1)
    else:
        start = date(today.year, today.month - 1, 1)
    # 下个月 1 号
    if today.month == 12:
        nxt = date(today.year + 1, 1, 1)
    else:
        nxt = date(today.year, today.month + 1, 1)
    # 下下个月 1 号减一天 = 下月最后一天
    if nxt.month == 12:
        end = date(nxt.year + 1, 1, 1) - timedelta(days=1)
    else:
        end = date(nxt.year, nxt.month + 1, 1) - timedelta(days=1)
    return start, end


def years_for_window(start: date, end: date) -> List[int]:
    """返回区间 [start, end] 跨越的年份（升序）。"""
    return list(range(start.year, end.year + 1))


# 同步锁：手动同步按钮与"检查前自动同步"可能并发读写 holidays.json，
# 用同一把锁保证拉取与落盘串行化，避免文件互相截断。
_SYNC_LOCK = threading.Lock()


def sync_window_years(store: "HolidayStore", today: date) -> Tuple[List[int], List[str]]:
    """同步"前后一个月"窗口所跨年份的节假日数据（best-effort）。

    同步范围与休假计划窗口一致——检查只关心"今天附近"的工作日，
    而窗口必然覆盖今天前后各一个月，足够判定今日触发项。

    Args:
        store: 节假日本地存储（同步成功即落盘）。
        today: 基准日期。
    Returns:
        (成功年份列表, 失败说明列表)；单个年份失败不抛异常，
        失败年份不会写入 store，因此不会污染已有缓存。
    """
    start, end = month_window_around(today)
    ok_years: List[int] = []
    failed: List[str] = []
    with _SYNC_LOCK:
        for year in years_for_window(start, end):
            try:
                store.update_year(year, fetch_year(year))
                ok_years.append(year)
            except HolidaySyncError as exc:
                failed.append(f"{year} 年：{exc}")
                logger.warning("节假日同步失败：%s 年 -> %s", year, exc)
    return ok_years, failed


# ----------------------------------------------------------------- 数据拉取 --

def _http_get_json(url: str) -> dict:
    """带 UA 的 HTTP GET，返回 JSON dict。"""
    req = urllib.request.Request(url, headers={"User-Agent": _UA})
    with urllib.request.urlopen(req, timeout=_HTTP_TIMEOUT) as resp:
        return json.loads(resp.read().decode("utf-8"))


def _parse_timor(payload: dict, year: int) -> Dict[str, dict]:
    """解析 timor.tech 响应为 {iso日期: {name, is_off}}。"""
    if payload.get("code") != 0 or not isinstance(payload.get("holiday"), dict):
        raise HolidaySyncError(f"timor.tech 响应格式异常：{str(payload)[:120]}")
    result: Dict[str, dict] = {}
    for _, v in payload["holiday"].items():
        if not isinstance(v, dict):
            continue
        d = str(v.get("date", ""))
        if not _DATE_RE.match(d):
            continue
        result[d] = {
            "name": str(v.get("name", "")) or "节假日",
            "is_off": bool(v.get("holiday", False)),
        }
    if not result:
        raise HolidaySyncError(f"timor.tech {year} 年数据为空")
    return result


def _parse_holiday_cn(payload: dict, year: int) -> Dict[str, dict]:
    """解析 holiday-cn 响应为 {iso日期: {name, is_off}}。"""
    days = payload.get("days")
    if not isinstance(days, list) or not days:
        raise HolidaySyncError(f"holiday-cn {year} 年数据为空（可能尚未发布）")
    result: Dict[str, dict] = {}
    for v in days:
        d = str(v.get("date", ""))
        if not _DATE_RE.match(d):
            continue
        result[d] = {
            "name": str(v.get("name", "")) or "节假日",
            "is_off": bool(v.get("isOffDay", False)),
        }
    if not result:
        raise HolidaySyncError(f"holiday-cn {year} 年数据解析为空")
    return result


def fetch_year(year: int) -> Dict[str, dict]:
    """拉取某年节假日数据，主源失败自动尝试备源。

    Returns:
        {iso日期: {"name": str, "is_off": bool}}
    Raises:
        HolidaySyncError: 所有数据源均失败。
    """
    errors: List[str] = []
    for url_tpl in HOLIDAY_SOURCES:
        url = url_tpl.format(year=year)
        try:
            payload = _http_get_json(url)
            parsed = (_parse_timor if "timor.tech" in url else _parse_holiday_cn)(
                payload, year)
            logger.info("节假日数据同步成功：%s（%d 条）", url, len(parsed))
            return parsed
        except Exception as exc:  # noqa: BLE001 —— 逐源尝试并记录失败原因
            errors.append(f"{url} -> {exc}")
            logger.warning("节假日数据源失败：%s -> %s", url, exc)
    raise HolidaySyncError(
        f"{year} 年节假日同步失败（所有数据源均不可达）：\n" + "\n".join(errors)
    )


# ----------------------------------------------------------------- 本地存储 --

class HolidayStore:
    """本地节假日数据存储：data/holidays.json。

    文件结构：
    {
      "synced_at": "2026-10-14T16:40:00",
      "years": {"2026": {"2026-01-01": {"name": "元旦", "is_off": true}, ...}}
    }
    """

    def __init__(self, path: Path) -> None:
        self._path = path
        self._data: dict = {"synced_at": None, "years": {}}
        self._load()

    def _load(self) -> None:
        if not self._path.exists():
            return
        try:
            raw = json.loads(self._path.read_text(encoding="utf-8"))
            if isinstance(raw, dict) and isinstance(raw.get("years"), dict):
                self._data = raw
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning("节假日本地数据损坏，已忽略（%s）", exc)

    def _flush(self) -> None:
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._path.write_text(
            json.dumps(self._data, ensure_ascii=False, indent=1),
            encoding="utf-8",
        )

    # -------------------------------------------------------------- 对外接口 --

    @property
    def synced_at(self) -> str:
        return self._data.get("synced_at") or "从未同步"

    @property
    def years(self) -> set:
        """已缓存的年份集合。"""
        return {int(y) for y in self._data["years"] if str(y).isdigit()}

    def has_year(self, year: int) -> bool:
        return str(year) in self._data["years"]

    def update_year(self, year: int, days: Dict[str, dict]) -> None:
        """写入某年数据并落盘（覆盖该年旧数据）。"""
        self._data["years"][str(year)] = days
        self._data["synced_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._flush()

    def workday_overrides(self) -> Dict[date, bool]:
        """导出为工作日覆盖表 {日期: 是否工作日}（供 WorkdayCalendar 使用）。"""
        overrides: Dict[date, bool] = {}
        for _, days in self._data["years"].items():
            for iso, info in days.items():
                try:
                    d = date.fromisoformat(iso)
                except ValueError:
                    continue
                # 休假日 -> 非工作日；调休补班 -> 工作日
                overrides[d] = not bool(info.get("is_off"))
        return overrides

    def special_days_between(self, start: date, end: date) -> List[SpecialDay]:
        """查询本地数据中 [start, end] 内的特殊日期（升序）。"""
        result: List[SpecialDay] = []
        for _, days in self._data["years"].items():
            for iso, info in days.items():
                try:
                    d = date.fromisoformat(iso)
                except ValueError:
                    continue
                if start <= d <= end:
                    result.append(SpecialDay(
                        date=d,
                        name=str(info.get("name", "")) or "节假日",
                        is_off=bool(info.get("is_off")),
                    ))
        result.sort(key=lambda s: s.date)
        return result


def package_special_days(start: date, end: date,
                         covered_years: set | None = None) -> List[SpecialDay]:
    """用 chinese_calendar 包内置数据补充区间内的特殊日期。

    仅对"本地尚无数据"的年份做补充（避免与本地数据重复），
    包无该年份数据时静默跳过。

    Args:
        start / end: 查询区间（含）。
        covered_years: 已有本地数据的年份集合。
    """
    covered = covered_years or set()
    result: List[SpecialDay] = []
    try:
        import chinese_calendar as cc
    except ImportError:  # pragma: no cover
        return result
    d = start
    while d <= end:
        if d.year not in covered:
            try:
                is_wd = bool(cc.is_workday(d))
                is_hol = bool(cc.is_holiday(d))
            except NotImplementedError:
                d += timedelta(days=1)
                continue
            if d.weekday() >= 5 and is_wd:
                result.append(SpecialDay(d, "调休补班", False))
            elif d.weekday() < 5 and is_hol:
                result.append(SpecialDay(d, "法定节假日", True))
        d += timedelta(days=1)
    return result
