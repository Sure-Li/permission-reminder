# -*- coding: utf-8 -*-
"""触发判断与发送记录去重模块。

职责：
1. ReminderItem / find_due_items：筛选"今天应提醒"的客户权限项
   （今天 == 到期日前推 5 个工作日；补检窗口内错过的标注补发）；
2. SentRecordStore：本地发送记录（客户号+权限+触发日期 三元组），
   手动检查 / 定时触发 / 启动补检共享，防止同日重复发送。
"""
from __future__ import annotations

import json
import logging
import threading
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from typing import Dict, List, Set

from .excel_reader import ClientRecord
from .workday import WorkdayCalendar, create_calendar

logger = logging.getLogger(__name__)

# 触发提前量：到期日前推 5 个工作日
ADVANCE_WORKDAYS = 5
# 补检窗口：错过触发日后 N 天内仍可补发（含当天）
BACKFILL_DAYS = 3


@dataclass
class ReminderItem:
    """一条待提醒的客户权限项。"""
    customer_id: str
    customer_name: str
    perm: str            # 权限标识：专业投资者 / 私募基金 / 资管权限
    perm_label: str      # 邮件正文使用的权限全称
    expiry: date         # 到期日期
    trigger_date: date   # 触发日期（到期日前推 5 个工作日）
    is_backfill: bool    # 是否补发（触发日期已过但未发送）


# 权限标识 -> 邮件正文权限名称
PERM_LABELS = {
    "专业投资者": "专业投资者权限",
    "私募基金": "私募基金权限",
    "资管权限": "私募资管权限",
}


def find_due_items(
    records: List[ClientRecord],
    selected_perms: Set[str],
    today: date,
    calendar: WorkdayCalendar | None = None,
    backfill_days: int = BACKFILL_DAYS,
) -> List[ReminderItem]:
    """筛选今天应提醒的权限项。

    Args:
        records: 客户记录列表（来自 excel_reader）。
        selected_perms: 界面勾选的权限标识集合（只检查勾选项）。
        today: 今天。
        calendar: 工作日日历；None 时用默认中国日历。
        backfill_days: 补检窗口天数。
    Returns:
        待提醒项列表（未做去重过滤，去重由 SentRecordStore 负责）。
    """
    cal = calendar or create_calendar()
    items: List[ReminderItem] = []
    for rec in records:
        for perm, expiry in rec.expiry_dates.items():
            if perm not in selected_perms:
                continue
            if expiry is None:  # NaN / 空 / 解析失败的跳过
                continue
            if expiry <= today:  # 已到期/今天到期的不再是"前推"场景，跳过
                continue
            trigger = cal.minus_workdays(expiry, ADVANCE_WORKDAYS)
            if today < trigger:
                continue  # 还没到触发日
            if (today - trigger).days > backfill_days:
                continue  # 超出补检窗口，放弃（避免长期关机后邮件风暴）
            items.append(ReminderItem(
                customer_id=rec.customer_id,
                customer_name=rec.customer_name,
                perm=perm,
                perm_label=PERM_LABELS.get(perm, perm),
                expiry=expiry,
                trigger_date=trigger,
                is_backfill=today > trigger,
            ))
    logger.info("触发筛选完成：勾选权限 %s，命中 %d 项（今天=%s）",
                sorted(selected_perms), len(items), today.isoformat())
    return items


class SentRecordStore:
    """发送记录存储：data/sent_records.json，线程安全。

    记录键："{customer_id}|{perm}|{trigger_date}"，含发送时间戳便于审计。
    """

    def __init__(self, path: Path) -> None:
        self._path = path
        self._lock = threading.Lock()
        # {key: sent_at_iso}
        self._records: Dict[str, str] = {}
        self._load()

    # -------------------------------------------------------------- 内部方法 --

    def _load(self) -> None:
        if not self._path.exists():
            return
        try:
            raw = json.loads(self._path.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                self._records = {str(k): str(v) for k, v in raw.items()}
        except (json.JSONDecodeError, OSError) as exc:
            logger.warning("发送记录文件损坏，已忽略（%s）", exc)
            self._records = {}

    def _flush(self) -> None:
        """落盘（调用方需已持有锁）。"""
        self._path.parent.mkdir(parents=True, exist_ok=True)
        self._path.write_text(
            json.dumps(self._records, ensure_ascii=False, indent=1),
            encoding="utf-8",
        )

    @staticmethod
    def _key(customer_id: str, perm: str, trigger_date: date) -> str:
        return f"{customer_id}|{perm}|{trigger_date.isoformat()}"

    # -------------------------------------------------------------- 对外接口 --

    def is_sent(self, customer_id: str, perm: str, trigger_date: date) -> bool:
        """该客户该权限在该触发日期是否已发送过。"""
        with self._lock:
            return self._key(customer_id, perm, trigger_date) in self._records

    def mark_sent(self, customer_id: str, perm: str,
                  trigger_date: date, sent_at: date) -> None:
        """标记已发送并立即落盘。"""
        with self._lock:
            self._records[self._key(customer_id, perm, trigger_date)] = \
                sent_at.isoformat()
            self._flush()

    def filter_unsent(self, items: List[ReminderItem]) -> List[ReminderItem]:
        """过滤掉已发送的项。"""
        with self._lock:
            return [i for i in items
                    if self._key(i.customer_id, i.perm, i.trigger_date)
                    not in self._records]

    def prune_before(self, keep_from: date) -> int:
        """清理 keep_from 之前的记录（按触发日期），返回清理条数。"""
        removed = 0
        with self._lock:
            kept: Dict[str, str] = {}
            for k, v in self._records.items():
                try:
                    trig = date.fromisoformat(k.rsplit("|", 1)[1])
                except (ValueError, IndexError):
                    continue  # 脏记录直接丢弃
                if trig >= keep_from:
                    kept[k] = v
                else:
                    removed += 1
            if removed:
                self._records = kept
                self._flush()
        return removed
