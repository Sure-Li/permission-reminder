# -*- coding: utf-8 -*-
"""Excel 读取模块。

职责：用 pandas 读取客户私募资管权限 Excel，解析为 ClientRecord 列表。

真实数据特征（必须兼容）：
- 第 1 行是合并大标题，真正的列名在第 2 行（header=1）；
- 日期单元格是 int 型 YYYYMMDD（如 20261021），也可能是 str / datetime /
  Excel 序列号；
- 客户姓名可能是数字（脏数据）、客户号可能为空、存在整行空行。
"""
from __future__ import annotations

import logging
import math
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

import pandas as pd

logger = logging.getLogger(__name__)

# Excel 中"到期日期"列名 -> 内部权限标识
EXPIRY_COLUMN_MAP: Dict[str, str] = {
    "专业投资者到期日期": "专业投资者",
    "私募基金到期日期": "私募基金",
    "资管权限到期日期": "资管权限",
}

# 必须存在的列（客户信息 + 三个到期日期列）
REQUIRED_COLUMNS = ["客户号", "客户姓名", *EXPIRY_COLUMN_MAP.keys()]

# Excel 序列号纪元基准（1900 日期系统，含 Lotus 1-2-3 的闰年 bug 偏移）
_EXCEL_EPOCH_ORDINAL = date(1899, 12, 30).toordinal()
# 仅接受现代业务日期范围内的序列号（2000-01-01 ~ 2099-12-31），
# 防止 999、12345 等垃圾小整数被误解析成 1900 年代的日期
_SERIAL_MIN = 36526   # 2000-01-01
_SERIAL_MAX = 73050   # 2099-12-31


class ExcelFormatError(Exception):
    """Excel 格式不符合预期（文件不存在 / 关键列缺失 / 读取失败）。"""


@dataclass
class ClientRecord:
    """一条客户权限记录。"""
    customer_id: str                            # 客户号（字符串，保留长度）
    customer_name: str                          # 客户姓名（原样输出，容忍脏数据）
    expiry_dates: Dict[str, Optional[date]] = field(default_factory=dict)
    # {"专业投资者": date | None, "私募基金": ..., "资管权限": ...}


def _is_nan_like(value: Any) -> bool:
    """判断值是否为 NaN / None / 空字符串 / 纯空白。"""
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    if isinstance(value, str) and value.strip() == "":
        return True
    try:  # pandas 的 pd.NA / pd.NaT 等
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def parse_date_cell(value: Any) -> Optional[date]:
    """解析日期单元格，兼容 int/str 的 YYYYMMDD、datetime、Excel 序列号。

    无法解析时返回 None（由调用方记录日志并跳过该权限）。
    """
    if _is_nan_like(value):
        return None
    # datetime / pd.Timestamp
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    # int / float：可能是 YYYYMMDD（如 20261021）或 Excel 序列号（如 46296）
    if isinstance(value, (int, float)):
        n = int(value)
        if 19000101 <= n <= 29991231:      # YYYYMMDD 形式
            try:
                return datetime.strptime(f"{n:08d}", "%Y%m%d").date()
            except ValueError:
                logger.warning("日期数值无法按 YYYYMMDD 解析: %r", value)
                return None
        if _SERIAL_MIN <= n <= _SERIAL_MAX:    # Excel 序列号形式（2000~2099 年）
            try:
                return date.fromordinal(_EXCEL_EPOCH_ORDINAL + n)
            except (ValueError, OverflowError):
                logger.warning("Excel 序列号超出合法日期范围: %r", value)
                return None
        logger.warning("日期数值无法识别（既非 YYYYMMDD 也非序列号）: %r", value)
        return None
    # 字符串：兼容 "20261021" / "2026-10-21" / "2026/10/21" / "2026年10月21日"
    if isinstance(value, str):
        s = value.strip()
        if s.isdigit() and len(s) == 8:
            try:
                return datetime.strptime(s, "%Y%m%d").date()
            except ValueError:
                pass
        for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y年%m月%d日"):
            try:
                return datetime.strptime(s, fmt).date()
            except ValueError:
                continue
        logger.warning("日期字符串无法解析: %r", value)
        return None
    logger.warning("未知日期单元格类型 %s: %r", type(value).__name__, value)
    return None


def read_clients(path: str | Path) -> List[ClientRecord]:
    """读取 Excel（第一个 sheet）并解析为客户记录列表。

    Raises:
        ExcelFormatError: 文件不存在 / 读取失败 / 关键列缺失。
    """
    p = Path(path)
    if not p.exists():
        raise ExcelFormatError(f"Excel 文件不存在：{p}")
    if p.suffix.lower() not in (".xlsx", ".xls"):
        raise ExcelFormatError(f"不是有效的 Excel 文件：{p.name}")

    try:
        # 第 1 行是合并大标题，真列名在第 2 行 -> header=1
        df = pd.read_excel(p, header=1, dtype=object, engine="openpyxl")
    except Exception as exc:  # noqa: BLE001 —— 统一转成业务异常
        raise ExcelFormatError(f"Excel 读取失败：{exc}") from exc

    # 按列名匹配（非位置），检查关键列
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ExcelFormatError(
            f"Excel 缺少关键列：{'、'.join(missing)}。\n"
            f"实际列名：{'、'.join(str(c) for c in df.columns)}"
        )

    records: List[ClientRecord] = []
    skipped = 0
    for idx, row in df.iterrows():
        cid = row["客户号"]
        if _is_nan_like(cid):
            # 空行或无客户号的行：跳过（整行全空不计警告）
            if not all(_is_nan_like(row[c]) for c in df.columns):
                logger.warning("第 %d 行客户号为空，已跳过", idx + 3)  # +3: 标题行+表头行+1基
                skipped += 1
            continue
        name = row["客户姓名"]
        customer_name = "" if _is_nan_like(name) else str(name).strip()
        expiry: Dict[str, Optional[date]] = {}
        for col, perm in EXPIRY_COLUMN_MAP.items():
            expiry[perm] = parse_date_cell(row[col])
        records.append(ClientRecord(
            customer_id=str(cid).strip(),
            customer_name=customer_name,
            expiry_dates=expiry,
        ))

    logger.info("Excel 解析完成：有效客户 %d 条，跳过 %d 条（%s）",
                len(records), skipped, p.name)
    return records
