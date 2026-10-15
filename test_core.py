# -*- coding: utf-8 -*-
"""核心逻辑 CLI 自测脚本（不启动 GUI）。

运行：.venv/Scripts/python.exe test_core.py
"""
from __future__ import annotations

import sys
import tempfile
from datetime import date
from pathlib import Path

from permission_reminder.core.checker import (
    ADVANCE_WORKDAYS, SentRecordStore, find_due_items,
)
from permission_reminder.core.excel_reader import (
    ExcelFormatError, parse_date_cell, read_clients,
)
from permission_reminder.core.workday import (
    ChineseWorkdayCalendar, WeekendOnlyCalendar, create_calendar,
)

TODAY = date(2026, 10, 14)  # 自测基准日

# ---------------------------------------------------------------------------
# 样例台账：自测不依赖任何真实客户数据，运行时在临时目录生成一份结构等价的 Excel。
# 覆盖的边界与真实台账一致：第 1 行合并大标题、真列名在第 2 行、日期为 int 型
# YYYYMMDD、存在姓名脏数据（"3"）、存在无客户号的汇总行（应被跳过）。
# ---------------------------------------------------------------------------
_FIXTURE_DIR = tempfile.TemporaryDirectory(prefix="pr_fixture_")
EXCEL = str(Path(_FIXTURE_DIR.name) / "sample_clients.xlsx")

# (客户号, 姓名, 专业投资者到期, 私募基金到期, 资管权限到期)；None 表示留空
_SAMPLE_ROWS = [
    ("590000000101", "测试甲", 20261021, 20261021, 20261021),
    ("590000000102", "测试乙", 20271013, None, 20261021),
    ("590000000103", "测试丙", 20261021, 20261021, 20270115),
    ("590000000104", "3", None, None, 20261021),        # 姓名脏数据
    ("590000000105", "测试戊", None, 20261021, 20270909),
    ("590000000106", "测试己", 20261021, 20261022, 20261022),
    (None, "汇总行", None, None, None),                  # 无客户号，应跳过
]
_SAMPLE_HEADERS = ["客户号", "客户姓名", "专业投资者到期日期",
                   "私募基金到期日期", "资管权限到期日期"]


def _build_sample_excel(path: Path) -> None:
    """生成样例台账：第 1 行大标题，第 2 行列名，第 3 行起为数据。"""
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font

    wb = Workbook()
    ws = wb.active
    ws.cell(row=1, column=1, value="客户私募资管权限有效期（样例数据，非真实客户）")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    for col, name in enumerate(_SAMPLE_HEADERS, start=1):
        ws.cell(row=2, column=col, value=name).font = Font(bold=True)
    for r, row in enumerate(_SAMPLE_ROWS, start=3):
        for col, value in enumerate(row, start=1):
            if value is not None:
                ws.cell(row=r, column=col, value=value)
    # 保留一个空 sheet，验证只读第一个 sheet
    wb.create_sheet("说明")
    wb.save(path)
    load_workbook(path).close()   # 确认文件可被 openpyxl 正常读回


_build_sample_excel(Path(EXCEL))

passed = failed = 0


def check(name: str, cond: bool, detail: str = "") -> None:
    global passed, failed
    if cond:
        passed += 1
        print(f"  [PASS] {name}")
    else:
        failed += 1
        print(f"  [FAIL] {name}  {detail}")


def _raise_when(fn) -> bool:
    """fn 应抛 ExcelFormatError，抛了返回 True。"""
    try:
        fn()
        return False
    except ExcelFormatError:
        return True


print("== 1. parse_date_cell 日期解析 ==")
check("int YYYYMMDD", parse_date_cell(20261021) == date(2026, 10, 21))
check("str YYYYMMDD", parse_date_cell("20261021") == date(2026, 10, 21))
check("str 2026-10-21", parse_date_cell("2026-10-21") == date(2026, 10, 21))
check("str 2026年10月21日", parse_date_cell("2026年10月21日") == date(2026, 10, 21))
import datetime as _dt
check("datetime", parse_date_cell(_dt.datetime(2026, 10, 21)) == date(2026, 10, 21))
check("Excel 序列号 46316 -> 2026-10-21", parse_date_cell(46316) == date(2026, 10, 21))
check("NaN -> None", parse_date_cell(float("nan")) is None)
check("None -> None", parse_date_cell(None) is None)
check("空字符串 -> None", parse_date_cell("  ") is None)
check("垃圾小整数 999 -> None", parse_date_cell(999) is None)
check("垃圾整数 12345 -> None", parse_date_cell(12345) is None)

print("== 2. 工作日计算 ==")
cal = ChineseWorkdayCalendar()
trigger = cal.minus_workdays(date(2026, 10, 21), ADVANCE_WORKDAYS)
check(f"2026-10-21 前5工作日 = 2026-10-14 (实际 {trigger})", trigger == date(2026, 10, 14))
# 周末跨越验证：到期日 2026-10-26（周一）前 5 个工作日 = 10-19
trigger2 = cal.minus_workdays(date(2026, 10, 26), 5)
check(f"2026-10-26（周一）前5工作日 = 2026-10-19 (实际 {trigger2})", trigger2 == date(2026, 10, 19))
# 降级日历：仅排周末
wk = WeekendOnlyCalendar()
check("仅周末日历：2026-10-21 前5 = 2026-10-14",
      wk.minus_workdays(date(2026, 10, 21), 5) == date(2026, 10, 14))
# 未来年份降级探测
_ = cal.is_workday(date(2099, 1, 4))
check("2099 年降级探测不崩溃", True)

print("== 3. 样例台账读取 ==")
try:
    records = read_clients(EXCEL)
    ids = [r.customer_id for r in records]
    print(f"  有效客户 {len(records)} 条: {ids}")
    check("读到 6 条有效记录（第 7 条无客户号已跳过）", len(records) == 6, f"实际 {len(records)}")
    check("首个客户号", records[0].customer_id == "590000000101")
    check("首条专业投资者到期日", records[0].expiry_dates["专业投资者"] == date(2026, 10, 21))
    check("首条私募基金到期日 2026-10-21", records[0].expiry_dates["私募基金"] == date(2026, 10, 21))
    check("脏姓名客户姓名原样输出", records[3].customer_name == "3")
    check("空单元格解析为 None", records[3].expiry_dates["专业投资者"] is None)
    by_id = {r.customer_id: r for r in records}
    check("资管权限到期 2026-10-21", by_id["590000000102"].expiry_dates["资管权限"] == date(2026, 10, 21))
    check("未来到期日不误解析", by_id["590000000102"].expiry_dates["专业投资者"] == date(2027, 10, 13))
except ExcelFormatError as exc:
    check("Excel 读取", False, str(exc))

print("== 4. 缺列异常 ==")
with tempfile.TemporaryDirectory() as td:
    import pandas as pd
    bad = Path(td) / "bad.xlsx"
    pd.DataFrame({"错误列": [1]}).to_excel(bad, index=False)
    try:
        read_clients(bad)
        check("缺列应抛 ExcelFormatError", False)
    except ExcelFormatError:
        check("缺列抛 ExcelFormatError", True)
    check("文件不存在抛 ExcelFormatError", _raise_when(lambda: read_clients(Path(td) / "nope.xlsx")))

print("== 5. 触发筛选（基准日 2026-10-14）==")
items = find_due_items(records, {"专业投资者", "私募基金", "资管权限"}, TODAY, cal)
for it in items:
    print(f"  命中: {it.customer_id} {it.customer_name} {it.perm} "
          f"到期{it.expiry} 触发{it.trigger_date} 补发={it.is_backfill}")
# 手工核算（基于样例台账 _SAMPLE_ROWS，到期日 > 10-14 的项）：
# - 0101: 专业投资者 10-21、私募基金 10-21、资管权限 10-21（触发均为 10-14）
# - 0102: 资管权限 10-21；专业投资者 2027-10-13（未来）
# - 0103: 专业投资者 10-21、私募基金 10-21；资管权限 2027-01-15（未来）
# - 0104: 资管权限 10-21（姓名脏数据）
# - 0105: 私募基金 10-21；资管权限 2027-09-09（未来）
# - 0106: 专业投资者 10-21；私募基金/资管 10-22（触发 10-15，明天不算）
expected_keys = {
    ("590000000101", "专业投资者"), ("590000000101", "私募基金"), ("590000000101", "资管权限"),
    ("590000000102", "资管权限"),
    ("590000000103", "专业投资者"), ("590000000103", "私募基金"),
    ("590000000104", "资管权限"),
    ("590000000105", "私募基金"),
    ("590000000106", "专业投资者"),
}
actual_keys = {(i.customer_id, i.perm) for i in items}
check(f"命中 9 项（实际 {len(items)}）", actual_keys == expected_keys,
      f"差异: {actual_keys ^ expected_keys}")
check("无补发项（都是当天触发）", all(not i.is_backfill for i in items))

# 勾选子集
items_sub = find_due_items(records, {"私募基金"}, TODAY, cal)
check("只勾私募基金时命中 3 项", len(items_sub) == 3, f"实际 {len(items_sub)}")

print("== 6. 补检场景（基准日 2026-10-16，错过 10-14）==")
items_late = find_due_items(records, {"专业投资者", "私募基金", "资管权限"},
                            date(2026, 10, 16), cal)
check("10-16 时 10-21 的项仍在补检窗口内", all(
    i.trigger_date >= date(2026, 10, 13) for i in items_late))
check("标注补发", all(i.is_backfill for i in items_late if i.trigger_date < date(2026, 10, 16)))
# 7528 的 10-22 到期（触发 10-15）在 10-16 应命中且为当天/补发
check("7528 私募基金在 10-16 命中", ("590000000106", "私募基金") in {(i.customer_id, i.perm) for i in items_late})

print("== 7. 发送记录去重 ==")
with tempfile.TemporaryDirectory() as td:
    store = SentRecordStore(Path(td) / "sent.json")
    check("初始未发送", not store.is_sent("590000000101", "专业投资者", date(2026, 10, 14)))
    store.mark_sent("590000000101", "专业投资者", date(2026, 10, 14), date(2026, 10, 14))
    check("标记后已发送", store.is_sent("590000000101", "专业投资者", date(2026, 10, 14)))
    filtered = store.filter_unsent(items)
    check(f"去重后剩余 {len(items)-1} 项", len(filtered) == len(items) - 1)
    # 重启后（重新加载）仍去重
    store2 = SentRecordStore(Path(td) / "sent.json")
    check("重启后记录仍在", store2.is_sent("590000000101", "专业投资者", date(2026, 10, 14)))
    # 清理旧记录（保留 10-15 及之后，10-14 的应被清掉）
    removed = store2.prune_before(date(2026, 10, 15))
    check("清理触发日之前的记录",
          removed == 1 and not store2.is_sent("590000000101", "专业投资者", date(2026, 10, 14)))

print("== 8. 节假日同步与工作日覆盖 ==")
from permission_reminder.core.holiday_sync import (
    HolidayStore, SpecialDay, _parse_holiday_cn, _parse_timor,
    month_window_around, package_special_days, sync_window_years, years_for_window,
)
from permission_reminder.core.workday import ChineseWorkdayCalendar

# 8.1 前后一个月窗口计算
s, e = month_window_around(date(2026, 10, 14))
check(f"10月基准窗口 = 2026-09-01 ~ 2026-11-30 (实际 {s} ~ {e})",
      s == date(2026, 9, 1) and e == date(2026, 11, 30))
s2, e2 = month_window_around(date(2026, 12, 31))
check(f"12月跨界窗口 = 2026-11-01 ~ 2027-01-31 (实际 {s2} ~ {e2})",
      s2 == date(2026, 11, 1) and e2 == date(2027, 1, 31))
s3, e3 = month_window_around(date(2027, 1, 5))
check(f"1月跨界窗口 = 2026-12-01 ~ 2027-02-28 (实际 {s3} ~ {e3})",
      s3 == date(2026, 12, 1) and e3 == date(2027, 2, 28))

# 8.2 数据源解析（离线样例，验证两种格式）
timor_sample = {"code": 0, "holiday": {
    "10-01": {"holiday": True, "name": "国庆节", "date": "2026-10-01"},
    "10-10": {"holiday": False, "name": "国庆后补班", "date": "2026-10-10"},
}}
parsed = _parse_timor(timor_sample, 2026)
check("timor 解析：休假日", parsed["2026-10-01"] == {"name": "国庆节", "is_off": True})
check("timor 解析：补班日", parsed["2026-10-10"] == {"name": "国庆后补班", "is_off": False})
cn_sample = {"year": 2026, "days": [
    {"name": "元旦", "date": "2026-01-01", "isOffDay": True},
    {"name": "春节前补班", "date": "2026-02-14", "isOffDay": False},
]}
parsed2 = _parse_holiday_cn(cn_sample, 2026)
check("holiday-cn 解析：休假日/补班日",
      parsed2["2026-01-01"]["is_off"] is True and parsed2["2026-02-14"]["is_off"] is False)

# 8.3 本地存储读写入与 workday 覆盖
with tempfile.TemporaryDirectory() as td:
    sp = Path(td) / "holidays.json"
    store = HolidayStore(sp)
    check("初始无缓存年份", store.years == set())
    store.update_year(2026, parsed)
    check("写入后含 2026", store.has_year(2026) and 2026 in store.years)
    sp2 = HolidayStore(sp)  # 重新加载（模拟重启）
    check("重启后数据仍在", sp2.has_year(2026))
    days = sp2.special_days_between(date(2026, 9, 15), date(2026, 10, 15))
    check(f"区间查询 2 项（实际 {len(days)}）", len(days) == 2)
    check("查询结果按日期升序", days[0].date < days[1].date)
    check("kind_label 正确", days[0].kind_label == "休" and days[1].kind_label == "补班")
    ov = sp2.workday_overrides()
    check("覆盖表：假日->非工作日", ov[date(2026, 10, 1)] is False)
    check("覆盖表：补班->工作日", ov[date(2026, 10, 10)] is True)

    # 覆盖优先于 chinese_calendar 内置数据
    cal_ov = ChineseWorkdayCalendar(overrides={date(2026, 10, 3): True})  # 周六强行上班
    check("覆盖优先：周六被设为工作日", cal_ov.is_workday(date(2026, 10, 3)) is True)
    check("未覆盖日期仍走常规判断（周日休）",
          cal_ov.is_workday(date(2026, 10, 4)) is False)

# 8.4 create_calendar 从指定文件加载覆盖
from permission_reminder.core.workday import create_calendar as _cc_factory
with tempfile.TemporaryDirectory() as td:
    sp = Path(td) / "holidays.json"
    HolidayStore(sp).update_year(2026, {"2026-10-03": {"name": "测试补班", "is_off": False}})
    cal_f = _cc_factory(sp)
    check("create_calendar 加载本地覆盖（周六上班）",
          cal_f.is_workday(date(2026, 10, 3)) is True)

# 8.5 包内置数据补充（不覆盖已有年份）
sup = package_special_days(date(2026, 9, 1), date(2026, 11, 30), covered_years={2026})
check("已有本地数据的年份不再补充", sup == [])
sup2 = package_special_days(date(2025, 9, 1), date(2025, 11, 30), covered_years=set())
check(f"无本地数据时用包内置补充 (?{len(sup2)} 项)", all(isinstance(x, SpecialDay) for x in sup2))

# 8.6 检查前自动同步：sync_window_years（打桩数据源，不联网）
print("== 8.6 检查前自动同步 ==")
import permission_reminder.core.holiday_sync as _hs

check("years_for_window 单年", years_for_window(date(2026, 9, 1), date(2026, 11, 30)) == [2026])
check("years_for_window 跨年",
      years_for_window(date(2026, 11, 1), date(2027, 1, 31)) == [2026, 2027])

_orig_fetch = _hs.fetch_year
_calls: list = []


def _stub_fetch(year: int) -> dict:
    _calls.append(year)
    return {f"{year}-01-01": {"name": "元旦", "is_off": True}}


_hs.fetch_year = _stub_fetch  # 模块级替换，sync_window_years 内按全局名查找
try:
    with tempfile.TemporaryDirectory() as td:
        st = HolidayStore(Path(td) / "holidays.json")
        # 注意：变量名不能叫 failed —— 会覆盖全局失败计数（check() 依赖它）
        ok, failed_years = sync_window_years(st, date(2026, 10, 14))
        check(f"同步年份 = 窗口年份 {[2026]}（实际 {_calls}）", _calls == [2026])
        check("成功年份列表 / 无失败", ok == [2026] and failed_years == [])
        check("数据已落盘且带同步时间",
              st.has_year(2026) and st.synced_at != "从未同步")
        check("重新加载后仍在（重启可用）",
              HolidayStore(Path(td) / "holidays.json").has_year(2026))

        # 跨年窗口会请求两年
        _calls.clear()
        sync_window_years(st, date(2026, 12, 20))
        check(f"跨年窗口请求 2 个年份（实际 {_calls}）", _calls == [2026, 2027])

        # 全部数据源失败：不抛异常、不污染已有缓存
        _calls.clear()
        st2 = HolidayStore(Path(td) / "h2.json")
        st2.update_year(2026, {"2026-10-01": {"name": "国庆节", "is_off": True}})

        def _bad_fetch(year: int) -> dict:
            _calls.append(year)
            raise _hs.HolidaySyncError(f"{year} 年节假日同步失败（所有数据源均不可达）")

        _hs.fetch_year = _bad_fetch
        ok2, failed2 = sync_window_years(st2, date(2026, 10, 14))
        check("全失败：返回空成功列表且不抛异常", ok2 == [] and len(failed2) == 1)
        check("全失败：失败信息含年份", "2026" in failed2[0])
        check("全失败：已有缓存未被污染", st2.has_year(2026))

        # 部分成功：一年成功一年失败
        def _half_fetch(year: int) -> dict:
            if year == 2027:
                raise _hs.HolidaySyncError("timor.tech 2027 年数据为空")
            return {f"{year}-01-01": {"name": "元旦", "is_off": True}}

        _hs.fetch_year = _half_fetch
        st3 = HolidayStore(Path(td) / "h3.json")
        ok3, failed3 = sync_window_years(st3, date(2026, 12, 20))
        check("部分成功：成功年份仍落盘", ok3 == [2026] and st3.has_year(2026))
        check("部分成功：失败年份不落盘", not st3.has_year(2027) and len(failed3) == 1)
finally:
    _hs.fetch_year = _orig_fetch

# 9. 配置项 auto_sync_holidays（默认开启 + 持久化 + 旧配置兼容）
print("== 9. 检查前自动同步配置项 ==")
import permission_reminder.infra.config_store as _cs

check("AppConfig 默认开启自动同步", _cs.AppConfig().auto_sync_holidays is True)

with tempfile.TemporaryDirectory() as td:
    _orig_cfg_path = _cs.CONFIG_PATH
    try:
        _cs.CONFIG_PATH = Path(td) / "config.json"
        _cs.save_config(_cs.AppConfig(auto_sync_holidays=False))
        check("关闭状态可持久化并读回", _cs.load_config().auto_sync_holidays is False)
        _cs.save_config(_cs.AppConfig(auto_sync_holidays=True))
        check("开启状态可持久化并读回", _cs.load_config().auto_sync_holidays is True)
        # 旧版本配置文件没有该字段 -> 默认开启
        _cs.CONFIG_PATH.write_text('{"excel_path": "old.xlsx"}', encoding="utf-8")
        check("旧配置缺该字段时默认开启（兼容升级）",
              _cs.load_config().auto_sync_holidays is True)
    finally:
        _cs.CONFIG_PATH = _orig_cfg_path

# =============================================== 10. 收件人校验与发信链路诊断
print("\n== 10. 收件人校验与发信链路诊断 ==")

import socket as _socket  # noqa: E402
import struct  # noqa: E402

from permission_reminder.infra.config_store import SmtpConfig  # noqa: E402
from permission_reminder.mailer import diagnose as dg  # noqa: E402
from permission_reminder.mailer.sender import (  # noqa: E402
    _auth_hint, _describe_failure, find_invalid_recipients, is_valid_email,
)

# --- 10.1 邮箱格式校验 ---
for good in ("a@b.com", "user@example.com", "recipient@example.com",
             "a.b-c_d+e@sub.example.com.cn"):
    check(f"合法邮箱 {good}", is_valid_email(good))
for bad in ("recipient@examplecom", "a@.com", "a@b", "a@b.c", "a b@c.com",
            "@qq.com", "a@@b.com", "a@b.com.", "a@b..com"):
    check(f"拦截非法邮箱 {bad}", not is_valid_email(bad), f"is_valid_email({bad!r})=True")
check("find_invalid_recipients 只挑出错的",
      find_invalid_recipients(["a@b.com", "recipient@examplecom", "c@d.cn"])
      == ["recipient@examplecom"])

# --- 10.2 MX 查询报文构造与解析（离线） ---
q = dg._build_mx_query("example.com", 0x1234)
check("MX 报文：事务 ID 与问题段正确",
      q[:2] == b"\x12\x34" and q[-4:] == struct.pack(">HH", 15, 1)
      and b"\x07example\x03com\x00" in q, q.hex())


def _mx_response(records, extra_a=False) -> bytes:
    """手工构造 DNS 响应报文（用于离线测试解析器）。"""
    hdr = struct.pack(">HHHHHH", 0x1234, 0x8180, 1, len(records) + int(extra_a), 0, 0)
    qsec = b"\x07example\x03com\x00" + struct.pack(">HH", 15, 1)
    body = b""
    for pref, host in records:
        rdata = struct.pack(">H", pref) + b"".join(
            bytes([len(p)]) + p.encode() for p in host.split(".")) + b"\x00"
        body += b"\xc0\x0c" + struct.pack(">HHIH", 15, 1, 300, len(rdata)) + rdata
    if extra_a:  # 混入一条 A 记录，应被跳过
        body += b"\xc0\x0c" + struct.pack(">HHIH", 1, 1, 300, 4) + bytes([203, 0, 113, 10])
    return hdr + qsec + body


recs = dg._parse_mx_response(_mx_response([(10, "mailex1.example.com"),
                                           (20, "mailex2.example.com")]))
check("解析 MX 记录（含压缩指针）",
      recs == [(10, "mailex1.example.com"), (20, "mailex2.example.com")], str(recs))
check("解析时跳过非 MX 记录",
      len(dg._parse_mx_response(_mx_response([(10, "mx3.qq.com")], extra_a=True))) == 1)

# --- 10.3 服务商识别 ---
# 回归缺陷：曾经的实现把 "mailexN" 前缀当作腾讯企业邮箱的特征，导致本案例误推荐
# smtp.exmail.qq.com。实测 example.com 是自建 Coremail + CACTER 网关，发信服务器就是 MX 主机本身。
p = dg.provider_from_mx(["mailex1.example.com", "mailex2.example.com"], "example.com")
check("mailexN 主机判为自建邮局（不再误判腾讯企业邮）",
      p is not None and p.name == "自建/自有邮件服务器"
      and p.server == "mailex1.example.com" and p.port == 25, str(p))
p = dg.provider_from_mx(["mx1.exmail.qq.com"], "foo.com")
check("明确的 exmail.qq.com 才判为腾讯企业邮箱",
      p is not None and p.name == "腾讯企业邮箱"
      and (p.server, p.port, p.ssl_mode) == ("smtp.exmail.qq.com", 465, "SSL"), str(p))
p = dg.provider_from_mx(["mxbiz1.qq.com"], "foo.com")
check("mxbiz 前缀判为腾讯企业邮箱", p is not None and p.name == "腾讯企业邮箱", str(p))
p = dg.provider_from_mx(["mx3.qq.com"], "qq.com")
check("识别腾讯个人邮箱", p is not None and (p.server, p.port) == ("smtp.qq.com", 465))
p = dg.provider_from_mx(["163mx01.mxmail.netease.com"], "163.com")
check("识别网易邮箱", p is not None and p.server == "smtp.163.com")
p = dg.provider_from_mx(["mail.example.com"], "example.com")
check("识别自建邮件服务器",
      p is not None and p.name.startswith("自建") and p.server == "mail.example.com", str(p))
check("空 MX 返回 None", dg.provider_from_mx([], "x.com") is None)

# --- 10.4 端口误用提示 ---
check("443 被识别为非 SMTP 端口",
      443 in dg.PORT_HINTS and "465" in dg.PORT_HINTS[443])
check("993/995 提示为收信端口",
      "IMAP" in dg.PORT_HINTS[993] and "POP3" in dg.PORT_HINTS[995])

# --- 10.5 diagnose_smtp 全链路（打桩网络，离线可跑） ---
_orig = (dg._tcp_probe, dg.mx_lookup, dg._try_login, dg.socket.getaddrinfo,
         dg.smtp_probe, dg.probe_mx_endpoints)
try:
    dg.socket.getaddrinfo = lambda h, p, *a, **k: [
        (2, 1, 6, "", ("203.0.113.10", p))]
    dg.mx_lookup = lambda d, timeout=5.0: [(10, "mailex1.example.com"),
                                           (10, "mailex2.example.com")]
    # 服务号探测打桩：mail.example.com 是"连上却零响应"的前端；mailex1 才是真 SMTP
    dg.smtp_probe = lambda host, port, timeout=8.0, do_ehlo=True: (
        dg.SmtpProbe(host, port, "silent", elapsed=8.0)
        if host == "mail.example.com" else
        dg.SmtpProbe(host, port, "smtp", banner="220 ", caps=["auth", "starttls"],
                     auth_mechs=["LOGIN", "PLAIN"], elapsed=0.1))
    dg.probe_mx_endpoints = lambda hosts, ports=(25,), timeout=6.0, max_hosts=3: [
        dg.SmtpProbe("mailex1.example.com", 25, "smtp", banner="220 ",
                     caps=["auth", "starttls"], elapsed=0.1)]

    # 场景 A：复刻用户现场 —— mail.example.com:443 SSL 超时 + 收件人漏点
    def probe_user_case(host, port, timeout=6.0):
        if host != "mail.example.com":
            return ("open", 0.1, "可连接")            # 出站基线可达
        if port == 25:
            return ("open", 0.02, "可连接")           # 25 秒连但无 SMTP 服务
        return ("timeout", 6.0, "超时无响应（防火墙丢包、端口未开放或服务未监听）")

    dg._tcp_probe = probe_user_case
    dg._try_login = lambda cfg, timeout=20.0: ("skip", "未调用")
    cfg = SmtpConfig(server="mail.example.com", port=443, ssl_mode="SSL",
                     username="user@example.com", password="S3cret-DoNotLeak",
                     recipients=["recipient@examplecom"])
    d = dg.diagnose_smtp(cfg)
    rep = dg.format_report(d)
    check("结论指出与账号密码无关", "与账号密码无关" in d.conclusion, d.conclusion)
    check("推荐配置为实测可用的 MX 端点 mailex1:25 STARTTLS",
          d.recommended == ("mailex1.example.com", 25, "STARTTLS"), str(d.recommended))
    check("443 端口误用被标为警告",
          any(s.name == "端口合理性" and s.status == "warn" for s in d.steps))
    check("收件人格式问题被指出",
          any(s.name == "收件人格式" and s.status == "fail" for s in d.steps))
    check("报告含结论/建议/推荐配置",
          "结论：" in rep and "建议：" in rep and "推荐配置" in rep)
    check("报告不含密码明文",
          "password" not in rep and "S3cret-DoNotLeak" not in rep)

    # 场景 A2：端点连得上却零响应（假连接）——不得把 TCP 通当成正常
    cfg_a2 = SmtpConfig(server="mail.example.com", port=25, ssl_mode="STARTTLS",
                        username="user@example.com", password="x",
                        recipients=["recipient@example.com"])
    d2 = dg.diagnose_smtp(cfg_a2)
    check("零响应端点结论点明不是邮件服务器",
          "没有返回任何 SMTP 服务号" in d2.conclusion, d2.conclusion)
    check("零响应时不尝试登录",
          any(s.name == "登录验证" and "未尝试登录" in s.detail for s in d2.steps))

    # 场景 A3：连接阶段即被网关拒绝 —— 必须完全跳过登录，且结论指向 IP 封禁
    dg.smtp_probe = lambda host, port, timeout=8.0, do_ehlo=True: dg.SmtpProbe(
        host, port, "reject",
        banner="554 reject by CACTER gateway, IP<1.2.3.4> is Hangup, "
               "Smtp auth error limit exceeded")
    called = []
    dg._try_login = lambda cfg, timeout=20.0: (called.append(1), ("ok", "不该被调用"))[1]
    d3 = dg.diagnose_smtp(cfg_a2)
    check("网关封禁时结论指向 IP 被封", "封禁" in d3.conclusion, d3.conclusion)
    check("网关封禁时绝不提交账号密码（保护出口 IP）", called == [], str(called))
    check("封禁建议含「等待自动解除」与「联系管理员」",
          any("自动解除" in s for s in d3.suggestions))
    dg.smtp_probe = lambda host, port, timeout=8.0, do_ehlo=True: (
        dg.SmtpProbe(host, port, "silent", elapsed=8.0)
        if host == "mail.example.com" else
        dg.SmtpProbe(host, port, "smtp", banner="220 ", caps=["auth", "starttls"],
                     elapsed=0.1))

    # 场景 B：服务器地址已改正、认证被拒（535）
    dg._tcp_probe = lambda host, port, timeout=6.0: ("open", 0.1, "可连接")
    dg._try_login = lambda cfg, timeout=20.0: (
        "auth", "服务器返回 535：Error: authentication failed, system busy")
    cfg = SmtpConfig(server="mailex1.example.com", port=25, ssl_mode="STARTTLS",
                     username="user@example.com", password="x",
                     recipients=["recipient@example.com"])
    d = dg.diagnose_smtp(cfg)
    check("认证失败结论指向密码类型", "认证被服务器拒绝" in d.conclusion, d.conclusion)
    check("自建邮局给出客户端专用密码等排查步骤",
          any("客户端专用密码" in s for s in d.suggestions))
    check("端点已是实测可用端点时不乱推推荐配置",
          d.recommended is None, str(d.recommended))

    # 场景 C：全部正常
    dg._try_login = lambda cfg, timeout=20.0: ("ok", "连接、加密与账号认证全部成功")
    d = dg.diagnose_smtp(cfg)
    check("全通时 ok=True 且无失败环节",
          d.ok is True and not any(s.status == "fail" for s in d.steps),
          d.conclusion)

    # 场景 D：配置缺字段直接返回
    d = dg.diagnose_smtp(SmtpConfig(server="", port=0, ssl_mode="SSL",
                                    username="", password="", recipients=[]))
    check("配置不全时提前返回并给出缺失项",
          not d.ok and "缺少" in d.conclusion and len(d.steps) == 1, d.conclusion)
finally:
    (dg._tcp_probe, dg.mx_lookup, dg._try_login, dg.socket.getaddrinfo,
     dg.smtp_probe, dg.probe_mx_endpoints) = _orig

# --- 10.6 错误文案分型 ---
_hint = _auth_hint("smtp.exmail.qq.com")
check("腾讯企业邮认证提示含专用密码步骤",
      "客户端专用密码" in _hint and "邮箱绑定" in _hint)
check("QQ 邮箱认证提示含授权码", "授权码" in _auth_hint("smtp.qq.com"))
_cfg443 = SmtpConfig(server="mail.example.com", port=443, ssl_mode="SSL",
                     username="a@b.com", password="x")
_msg = _describe_failure(_cfg443, _socket.timeout("timed out"))
check("超时文案点明端口误用", "443 是 HTTPS 端口" in _msg, _msg[:80])
check("超时文案引导用连接诊断", "连接诊断" in _msg)
check("超时文案说明与密码无关", "与账号密码无关" in _msg)
_msg_gai = _describe_failure(_cfg443, _socket.gaierror("getaddrinfo failed"))
check("域名解析失败有独立文案", "无法解析" in _msg_gai)

print("\n== 11. 服务号探测与假连接识别（本地桩）==")

import socket as _sock  # noqa: E402
import threading as _th  # noqa: E402
import time as _time  # noqa: E402


class _SmtpStub:
    """最小 SMTP 桩服务器（仅 127.0.0.1），用于离线验证探测与发送路径。

    greeting 为空表示"接受连接但什么都不发"（模拟被安全设备接管/非邮件服务器的假连接）。
    """

    def __init__(self, greeting: bytes = b"220 stub ESMTP\r\n",
                 auth_ok: bool = True, advertise_starttls: bool = False,
                 auth_error: bytes = b"535 5.7.8 Authentication credentials invalid\r\n") -> None:
        self.greeting = greeting
        self.auth_ok = auth_ok
        self.auth_error = auth_error
        self.advertise_starttls = advertise_starttls
        self.connections = 0
        self._sock = _sock.socket()
        self._sock.setsockopt(_sock.SOL_SOCKET, _sock.SO_REUSEADDR, 1)
        self._sock.bind(("127.0.0.1", 0))
        self._sock.listen(5)
        self.port = self._sock.getsockname()[1]
        _th.Thread(target=self._serve, daemon=True).start()

    def _serve(self) -> None:
        while True:
            try:
                conn, _ = self._sock.accept()
            except OSError:
                return
            self.connections += 1
            _th.Thread(target=self._handle, args=(conn,), daemon=True).start()

    def _handle(self, conn) -> None:
        f = conn.makefile("rwb")
        try:
            if not self.greeting:
                _time.sleep(6)          # 接受连接但保持沉默
                return
            f.write(self.greeting)
            f.flush()
            if not self.greeting.startswith(b"220"):
                return                  # 拒绝类服务号（如 554 网关封禁）
            in_data = False
            while True:
                line = f.readline()
                if not line:
                    return
                if in_data:
                    if line.strip() == b".":
                        in_data = False
                        f.write(b"250 OK\r\n")
                        f.flush()
                    continue
                up = line.strip().upper()
                if up.startswith((b"EHLO", b"HELO")):
                    caps = [b"250-stub", b"250-PIPELINING", b"250-AUTH PLAIN"]
                    if self.advertise_starttls:
                        caps.append(b"250-STARTTLS")
                    caps.append(b"250 8BITMIME")
                    f.write(b"\r\n".join(caps) + b"\r\n")
                    f.flush()
                elif up.startswith(b"AUTH"):
                    f.write(b"235 2.7.0 ok\r\n" if self.auth_ok else self.auth_error)
                    f.flush()
                elif up.startswith(b"DATA"):
                    in_data = True
                    f.write(b"354 End data with <CR><LF>.<CR><LF>\r\n")
                    f.flush()
                elif up.startswith(b"QUIT"):
                    f.write(b"221 Bye\r\n")
                    f.flush()
                    return
                else:
                    f.write(b"250 OK\r\n")
                    f.flush()
        except Exception:  # noqa: BLE001 —— 桩服务器，异常无需处理
            pass
        finally:
            try:
                conn.close()
            except OSError:
                pass

    def close(self) -> None:
        try:
            self._sock.close()
        except OSError:
            pass


# 11.1 真实 SMTP 服务号 + 能力识别
_stub = _SmtpStub(advertise_starttls=True)
_p = dg.smtp_probe("127.0.0.1", _stub.port, timeout=3)
check("探测到真实 SMTP 服务号（220）", _p.usable and _p.status == "smtp", _p.summary)
check("从 EHLO 中识别出 STARTTLS 能力", _p.supports_starttls, str(_p.caps))
check("从 EHLO 中识别出 AUTH 机制", _p.auth_mechs == ["PLAIN"], str(_p.auth_mechs))
_stub.close()

# 11.2 连上却零响应（假连接）——绝不能当成"网络正常"
_silent = _SmtpStub(greeting=b"")
_p2 = dg.smtp_probe("127.0.0.1", _silent.port, timeout=2)
check("零响应连接被识别为 silent 而非可用",
      _p2.status == "silent" and not _p2.usable, _p2.summary)
check("零响应说明中提示'不是邮件服务器'或'安全设备'",
      "不是邮件服务器" in _p2.summary or "安全设备" in _p2.summary, _p2.summary)
_silent.close()

# 11.3 网关拒绝类服务号
_rej = _SmtpStub(greeting=b"554 reject by CACTER gateway, IP<1.2.3.4> is Hangup, "
                          b"Smtp auth error limit exceeded\r\n")
_p3 = dg.smtp_probe("127.0.0.1", _rej.port, timeout=3)
check("网关封禁服务号被识别为 reject", _p3.status == "reject", _p3.summary)
check("封禁原文被完整保留（便于用户上报）", "CACTER" in _p3.banner)
_rej.close()

# 11.4 最优端点选择
_probes = [dg.SmtpProbe("a", 465, "smtp", caps=["auth"]),
           dg.SmtpProbe("b", 25, "smtp", caps=["auth", "starttls"]),
           dg.SmtpProbe("c", 587, "smtp", caps=["auth"])]
_best = dg._pick_best_endpoint(_probes)
check("优先选 25 端口且支持 STARTTLS 的端点",
      _best is not None and _best.host == "b" and _best.port == 25, str(_best))
check("无可用端点时返回 None",
      dg._pick_best_endpoint([dg.SmtpProbe("x", 25, "silent")]) is None)

print("\n== 12. 发送路径：加密模式与反封禁（本地桩）==")

from permission_reminder.mailer.sender import (  # noqa: E402
    EmailSender, MailSendError, auth_block_remaining, is_gateway_rejection,
    reset_auth_block,
)

check("网关拒绝文本可被识别", is_gateway_rejection("554 reject by CACTER gateway"))
check("普通业务错误不被误判为封禁",
      not is_gateway_rejection("550 mailbox unavailable"))

# 12.1 不加密模式（自建邮局 25 端口的兜底方案）
reset_auth_block()
_st = _SmtpStub()
_cfg = SmtpConfig(server="127.0.0.1", port=_st.port, ssl_mode="PLAIN",
                  username="me@x.com", password="pw", recipients=["to@x.com"])
try:
    EmailSender(_cfg).send(["to@x.com"], "【权限到期提醒】", "<p>h</p>", "t")
    check("PLAIN 模式发送成功", True)
except MailSendError as _e:
    check("PLAIN 模式发送成功", False, str(_e)[:100])
check("发送过程只建立一次连接（无多余重试）",
      _st.connections == 1, f"connections={_st.connections}")
_st.close()

# 12.2 STARTTLS 而服务器不支持 → 明确报错且不重试
_st2 = _SmtpStub(advertise_starttls=False)
_cfg2 = SmtpConfig(server="127.0.0.1", port=_st2.port, ssl_mode="STARTTLS",
                   username="me@x.com", password="pw", recipients=["to@x.com"])
try:
    EmailSender(_cfg2).send(["to@x.com"], "s", "h", "p")
    check("服务器不支持 STARTTLS 时明确报错", False, "未报错")
except MailSendError as _e:
    check("服务器不支持 STARTTLS 时明确报错", "STARTTLS" in str(_e), str(_e)[:100])
    check("报错给出改用 SSL 或「不加密」的指引",
          "不加密" in str(_e) or "SSL" in str(_e), str(_e)[:100])
check("加密方式不配套时不重试", _st2.connections == 1, f"connections={_st2.connections}")
_st2.close()

# 12.3 认证失败 → 不重试 + 进入冷静期 + 冷静期内不再触碰服务器
reset_auth_block()
_st3 = _SmtpStub(auth_ok=False)
_cfg3 = SmtpConfig(server="127.0.0.1", port=_st3.port, ssl_mode="PLAIN",
                   username="me@x.com", password="bad", recipients=["to@x.com"])
try:
    EmailSender(_cfg3).send(["to@x.com"], "s", "h", "p")
    check("认证失败抛出 MailSendError", False, "未报错")
except MailSendError as _e:
    check("认证失败抛出 MailSendError", "认证失败" in str(_e), str(_e)[:100])
    check("认证失败提示含自建邮局排查步骤（客户端专用密码等）",
          "客户端专用密码" in str(_e), str(_e)[:120])
check("认证失败只尝试一次（不重试、不加封禁计数）",
      _st3.connections == 1, f"connections={_st3.connections}")
check("认证失败后进入发送冷静期", auth_block_remaining() > 0)
_n_before = _st3.connections
try:
    EmailSender(_cfg3).send(["to@x.com"], "s", "h", "p")
    check("冷静期内直接拦截、不再连服务器", False, "未报错")
except MailSendError as _e:
    check("冷静期内直接拦截、不再连服务器",
          "暂停发送" in str(_e) and _st3.connections == _n_before,
          f"connections={_st3.connections}")
_st3.close()
reset_auth_block()
check("reset_auth_block 可解除冷静期", auth_block_remaining() == 0)

# 12.4 网关连接阶段封禁（554）→ 不重试 + 提示封禁 + 冷静期
reset_auth_block()
_st4 = _SmtpStub(greeting=b"554 reject by CACTER gateway, IP<1.2.3.4> is Hangup, "
                          b"Smtp auth error limit exceeded\r\n")
_cfg4 = SmtpConfig(server="127.0.0.1", port=_st4.port, ssl_mode="PLAIN",
                   username="me@x.com", password="pw", recipients=["to@x.com"])
try:
    EmailSender(_cfg4).send(["to@x.com"], "s", "h", "p")
    check("网关封禁被识别并给出专项提示", False, "未报错")
except MailSendError as _e:
    check("网关封禁被识别并给出专项提示", "安全网关" in str(_e), str(_e)[:120])
    check("封禁提示说明需等待自动解除", "自动解除" in str(_e))
check("网关封禁不重试", _st4.connections == 1, f"connections={_st4.connections}")
check("网关封禁后同样进入冷静期", auth_block_remaining() > 0)
_st4.close()
reset_auth_block()

# 12.5 同样是 535，还必须区分「网关策略拒绝」与「凭据真的错」——处置方式完全不同
#      （example.com 实测返回：(535, b'reject by CACTER gateway, Error: Unknown error')）
from permission_reminder.mailer.sender import is_gateway_auth_rejection  # noqa: E402

check("识别 CACTER 网关回的 535",
      is_gateway_auth_rejection("reject by CACTER gateway, Error: Unknown error"))
check("普通凭据错误不被误判为网关拒绝",
      not is_gateway_auth_rejection("5.7.8 Authentication credentials invalid"))
check("服务器忙/密码错类 535 不误判为网关拒绝",
      not is_gateway_auth_rejection("Error: authentication failed, system busy"))

_st5 = _SmtpStub(auth_ok=False,
                 auth_error=b"535 reject by CACTER gateway, Error: Unknown error\r\n")
_cfg5 = SmtpConfig(server="127.0.0.1", port=_st5.port, ssl_mode="PLAIN",
                   username="user@example.com", password="pw", recipients=["to@x.com"])
try:
    EmailSender(_cfg5).send(["to@x.com"], "s", "h", "p")
    check("网关 535 抛出 MailSendError", False, "未报错")
except MailSendError as _e:
    _msg = str(_e)
    check("网关 535 明确点出「未必是账号密码写错」", "未必是账号密码写错" in _msg, _msg[:120])
    check("网关 535 给出四步自检清单（含客户端专用密码与 IP 白名单）",
          all(k in _msg for k in ("①", "②", "③", "④", "客户端专用密码", "出口 IP")),
          _msg[:160])
    check("网关 535 仍提醒勿连续重试", "勿连续重试" in _msg)
check("网关 535 只尝试一次（不加重封禁计数）",
      _st5.connections == 1, f"connections={_st5.connections}")
_st5.close()
reset_auth_block()

# 12.6 自检清单第 ④ 步必须"按本机实情说话"：
#      它依赖"手工用邮件客户端发一封测试邮件"，而这台机器可能压根没装 Foxmail、
#      mailto 也没有默认程序 → 一律写"用 Foxmail 试试"是无效建议。
from permission_reminder.mailer.sender import (  # noqa: E402
    _probe_local_mail_client, _selfhost_checklist,
)

_c_yes = _selfhost_checklist((True, "Foxmail"))
check("清单④-有客户端：点出具体客户端名", "Foxmail" in _c_yes, _c_yes[:160])
check("清单④-有客户端：保留判定方法与出口 IP",
      "能发出" in _c_yes and "出口 IP" in _c_yes and "同样发不出" in _c_yes)
check("清单④-有客户端：不出现「走不通」字样", "走不通" not in _c_yes)

_c_no = _selfhost_checklist((False, "未设置默认邮件客户端"))
check("清单④-无客户端：明确指出该步骤在本机走不通",
      "走不通" in _c_no and "出口 IP" in _c_no, _c_no[:200])
check("清单④-无客户端：指向本工具的兜底模式",
      "发送方式" in _c_no and "不需要授权码" in _c_no)
check("清单④-无客户端：给出装 Foxmail 的替代路径", "装上 Foxmail" in _c_no)

_c_unk = _selfhost_checklist((None, ""))
check("清单④-探测不可用：回退到通用说法",
      "Foxmail" in _c_unk and "出口 IP" in _c_unk and "走不通" not in _c_unk)

check("清单①②③在三种分支下都完整保留",
      all(k in s for s in (_c_yes, _c_no, _c_unk) for k in ("①", "②", "③", "④")))
check("清单仍提醒勿连续重试",
      all("勿连续重试" in s for s in (_c_yes, _c_no, _c_unk)))

_ok_p, _name_p = _probe_local_mail_client()
check("探测函数返回 (bool|None, str)",
      (_ok_p is None or isinstance(_ok_p, bool)) and isinstance(_name_p, str),
      f"{_ok_p!r} / {_name_p!r}")
print(f"     [本机实测] 默认邮件客户端可用={_ok_p}  描述={_name_p!r}")

# ============================================================================
print("\n== 13. 兜底通道：导出 .eml/.html + 唤起本机邮件客户端 ==")

import json as _json  # noqa: E402
import urllib.parse  # noqa: E402
from email import policy as _policy  # noqa: E402
from email.parser import BytesParser  # noqa: E402

import permission_reminder.mailer.outbox as ob  # noqa: E402

check("非法文件名字符被替换", ob._safe_name('a<b>c:d/e"f\\g|h?i*j') == "a_b_c_d_e_f_g_h_i_j",
      ob._safe_name('a<b>c:d/e"f\\g|h?i*j'))

# --- 13.1 导出：文件名、目录、内容往返一致 ---
with tempfile.TemporaryDirectory() as _td:
    _orig_data = ob.DATA_DIR
    ob.DATA_DIR = Path(_td)                      # 重定向导出目录，避免污染真实 data
    try:
        _recipients = ["boss@x.com", "leader@x.com"]
        _subject = "【权限到期提醒】"
        _html = "<html><body><b>测试</b></body></html>"
        _plain = "测试正文\n第二行"
        _eml, _htm = ob.export_mail(_recipients, _subject, _html, _plain,
                                    from_addr="user@example.com", tag="权限到期提醒")

        check("导出 .eml 文件存在", _eml.exists() and _eml.suffix == ".eml", str(_eml))
        check("导出 .html 文件存在", _htm.exists() and _htm.suffix == ".html", str(_htm))
        check("导出文件落到 outbox 目录", _eml.parent.name == "outbox", str(_eml.parent))
        check("文件名含中文标签且不含非法字符",
              "权限到期提醒" in _eml.name and not any(c in _eml.name for c in '<>:"/\\|?*'),
              _eml.name)

        _msg = BytesParser(policy=_policy.default).parsebytes(_eml.read_bytes())
        check(".eml 主题可正确还原（中文不乱码）", _msg["Subject"] == _subject, str(_msg["Subject"]))
        check(".eml 收件人完整", all(r in _msg["To"] for r in _recipients), str(_msg["To"]))
        check(".eml 发件人写入", _msg["From"] == "user@example.com", str(_msg["From"]))
        _parts = [_p.get_content_type() for _p in _msg.walk() if not _p.is_multipart()]
        check(".eml 含纯文本 + HTML 两个部分",
              _parts == ["text/plain", "text/html"], str(_parts))
        _body_html = [_p for _p in _msg.walk()
                      if _p.get_content_type() == "text/html"][0].get_content()
        check(".eml HTML 正文与源一致", "<b>测试</b>" in _body_html, _body_html[:80])
        check(".html 预览可被浏览器按 UTF-8 打开",
              'charset="utf-8"' in _htm.read_text(encoding="utf-8"))

        # 同一秒内连续导出也不应互相覆盖（时间戳相同 -> 自动追加序号）
        _eml2, _htm2 = ob.export_mail(_recipients, _subject, _html, _plain, tag="权限到期提醒")
        check("同一秒内连续导出不会互相覆盖",
              len({_eml.name, _eml2.name}) == 2, f"{_eml.name} / {_eml2.name}")
        check("成对的 .html 与 .eml 同主名（便于识别）",
              _eml.stem == _htm.stem, f"{_eml.stem} / {_htm.stem}")

        # --- 13.2 mailto 构造 ---
        _url = ob.build_mailto(_recipients, _subject, _plain)
        check("mailto 指向全部收件人", _url.startswith("mailto:boss@x.com,leader@x.com?"), _url[:60])
        check("mailto 主题/正文已编码（无裸中文）",
              "%E3%80%90" in _url and " " not in _url and "\n" not in _url)
        check("mailto 正文换行编码为 %0A", "%0A" in _url)

        _long = "x" * (ob.MAILTO_BODY_LIMIT + 500)
        _url2 = ob.build_mailto(_recipients, _subject, _long)
        check("超长正文被截断并提示看 .eml", "截断" in urllib.parse.unquote(_url2),
              urllib.parse.unquote(_url2)[-60:])
    finally:
        ob.DATA_DIR = _orig_data

# --- 13.3 唤起客户端：先判定系统有没有 mailto 默认程序 ---
_orig_query = ob._query_assoc
_orig_startfile = getattr(ob.os, "startfile", None)
_calls: list = []
try:
    if _orig_startfile is not None:
        ob.os.startfile = lambda u: _calls.append(u)  # type: ignore[attr-defined]

    # (a) 没有默认邮件客户端：Windows 会弹「选取应用」，必须**不要**打开
    ob._query_assoc = lambda proto, kind: (
        r"C:\Windows\system32\OpenWith.exe" if kind == 2 else "选取应用")
    _has = ob.has_mail_client()
    _ret_a = ob.open_in_mail_client(["a@b.com"], "主题", "正文")
    check("无默认邮件客户端时 has_mail_client 为 False", _has is False)
    check("无默认客户端时不唤起（避免弹「选取应用」）",
          _ret_a is False and _calls == [], f"ret={_ret_a} calls={_calls}")

    # (b) 有默认邮件客户端：正常走 mailto:
    ob._query_assoc = lambda proto, kind: (
        r"C:\Program Files\Foxmail\Foxmail.exe" if kind == 2 else "Foxmail")
    check("有默认客户端时 has_mail_client 为 True", ob.has_mail_client() is True)
    check("描述文本给出客户端名", ob.describe_mail_client() == "Foxmail",
          ob.describe_mail_client())
    _ret_b = ob.open_in_mail_client(["a@b.com"], "主题", "正文")
    check("有客户端时用 mailto: 唤起",
          _ret_b is True and len(_calls) == 1 and str(_calls[0]).startswith("mailto:a@b.com?"),
          f"ret={_ret_b} calls={_calls}")
finally:
    ob._query_assoc = _orig_query
    if _orig_startfile is not None:
        ob.os.startfile = _orig_startfile  # type: ignore[attr-defined]

# 真实环境自检：只要求返回合理类型，不锁定具体机器
check("真实环境 mailto 探测返回 bool", isinstance(ob.has_mail_client(), bool))
check("真实环境客户端描述返回字符串", isinstance(ob.describe_mail_client(), str))

# --- 13.4 配置项 send_mode：默认值、持久化、脏值回退 ---
with tempfile.TemporaryDirectory() as _td:
    _orig_cfg = _cs.CONFIG_PATH
    _cs.CONFIG_PATH = Path(_td) / "config.json"
    try:
        _cs.save_config(_cs.AppConfig(send_mode="client"))
        check("send_mode 可持久化并读回", _cs.load_config().send_mode == "client")
        _cs.CONFIG_PATH.write_text('{"excel_path": "old.xlsx"}', encoding="utf-8")
        check("旧配置无 send_mode 时默认 SMTP 直发（向后兼容）",
              _cs.load_config().send_mode == "smtp")
        _cs.CONFIG_PATH.write_text('{"send_mode": "telepathy"}', encoding="utf-8")
        check("send_mode 被改坏时回落到 smtp", _cs.load_config().send_mode == "smtp")
        check("send_mode 不在配置字段外泄",
              "SEND_MODES" not in _json.dumps(_cs.asdict(_cs.AppConfig())),
              "SEND_MODES 被当成数据字段序列化了")
    finally:
        _cs.CONFIG_PATH = _orig_cfg

# ============================================================================
print("\n== 14. Foxmail 集成：自动定位主程序 + 用 Foxmail 投递 ==")

import permission_reminder.mailer.foxmail as fx  # noqa: E402

# --- 14.1 从注册表 command 字符串里解析 exe 路径 ---
check("解析带引号的 shell\\open\\command",
      fx.parse_command_path('"C:\\Program Files (x86)\\Foxmail\\Foxmail.exe" "%1"')
      == r"C:\Program Files (x86)\Foxmail\Foxmail.exe",
      fx.parse_command_path('"C:\\Program Files (x86)\\Foxmail\\Foxmail.exe" "%1"'))
check("解析无引号的 command（后面带参数）",
      fx.parse_command_path(r"D:\Foxmail\Foxmail.exe %1") == r"D:\Foxmail\Foxmail.exe",
      fx.parse_command_path(r"D:\Foxmail\Foxmail.exe %1"))
check("解析带 ,0 资源索引的 command",
      fx.parse_command_path('"C:\\Foxmail\\Foxmail.exe",0') == r"C:\Foxmail\Foxmail.exe",
      fx.parse_command_path('"C:\\Foxmail\\Foxmail.exe",0'))
check("空 command 返回空串", fx.parse_command_path("") == "")
_expanded = fx.parse_command_path(r"%ProgramFiles%\Foxmail\Foxmail.exe")
check("command 中的环境变量被展开", "%" not in _expanded, _expanded)

# --- 14.2 路径有效性校验 ---
check("不存在的路径不算有效 Foxmail", fx.is_valid_foxmail_exe(r"C:\nope\Foxmail.exe") is False)
check("空路径不算有效 Foxmail", fx.is_valid_foxmail_exe("") is False)

with tempfile.TemporaryDirectory() as _td:
    _fake_dir = Path(_td) / "Foxmail"
    _fake_dir.mkdir()
    _fake_exe = _fake_dir / "Foxmail.exe"
    _fake_exe.write_bytes(b"MZ fake")
    _other_exe = _fake_dir / "Other.exe"
    _other_exe.write_bytes(b"MZ fake")

    check("存在且文件名正确的算有效", fx.is_valid_foxmail_exe(str(_fake_exe)) is True)
    check("文件名不是 Foxmail.exe 的算无效",
          fx.is_valid_foxmail_exe(str(_other_exe)) is False)

    # --- 14.3 探测与 resolve ---
    check("detect_foxmail 接受调用方给的候选路径并优先采用",
          fx.detect_foxmail(extra=[str(_fake_exe)]) == str(_fake_exe),
          fx.detect_foxmail(extra=[str(_fake_exe)]))
    check("resolve 尊重界面已填的有效路径",
          fx.resolve(str(_fake_exe))[0] == str(_fake_exe))
    _exe_bad, _note_bad = fx.resolve(str(_other_exe))
    check("resolve 遇到无效路径会回落探测并给出可读说明",
          isinstance(_note_bad, str) and len(_note_bad) > 10, _note_bad[:80])

    # --- 14.4 命令行构造（纯函数，不启动进程）---
    _eml = _fake_dir / "权限到期提醒_20261014-093000.eml"
    _eml.write_bytes(b"From: a\r\nTo: b\r\n\r\nhello")
    _rcpts = ["recipient@example.com"]

    _argv_eml = fx.build_launch_argv(str(_fake_exe), "eml", str(_eml),
                                     _rcpts, "【权限到期提醒】", "正文")
    check("eml 方式：argv = [exe, eml 路径]",
          _argv_eml == [str(_fake_exe), str(_eml)], str(_argv_eml))
    _argv_ma = fx.build_launch_argv(str(_fake_exe), "mailto", "", _rcpts,
                                    "【权限到期提醒】", "第一行\n第二行")
    check("mailto 方式：argv[1] 是 mailto 链接",
          len(_argv_ma) == 2 and _argv_ma[1].startswith("mailto:recipient@example.com?"),
          _argv_ma[1][:80])
    check("mailto 方式：主题已编码进链接（无裸中文）",
          "%E6%9D%83%E9%99%90" in _argv_ma[1], _argv_ma[1][:120])
    check("auto 方式：有 .eml 时选 eml（更可靠）",
          fx.build_launch_argv(str(_fake_exe), "auto", str(_eml), _rcpts, "s", "p")
          == [str(_fake_exe), str(_eml)])
    check("auto 方式：无 .eml 时退回 mailto",
          fx.build_launch_argv(str(_fake_exe), "auto", "", _rcpts, "s", "p")[1]
          .startswith("mailto:"))

    # 参数不满足所选方式时必须明确报错，不能静默失败
    for _mode, _path, _desc in (
        ("telepathy", str(_eml), "未知唤起方式"),
        ("eml", "", "eml 方式但没给文件"),
        ("eml", str(_fake_dir / "不存在.eml"), "eml 文件不存在"),
    ):
        try:
            fx.build_launch_argv(str(_fake_exe), _mode, _path, _rcpts, "s", "p")
            check(f"build_launch_argv 拦住：{_desc}", False, "竟然没报错")
        except fx.FoxmailError:
            check(f"build_launch_argv 拦住：{_desc}", True)
    try:
        fx.build_launch_argv("", "auto", str(_eml), _rcpts, "s", "p")
        check("build_launch_argv 拦住空 exe 路径", False, "竟然没报错")
    except fx.FoxmailError:
        check("build_launch_argv 拦住空 exe 路径", True)

    # --- 14.5 真正唤起：替换 subprocess.Popen，不启动任何进程 ---
    _spawned = []
    _orig_popen = fx.subprocess.Popen
    try:
        fx.subprocess.Popen = lambda argv, **kw: _spawned.append((argv, kw))  # type: ignore
        _ret_argv = fx.open_foxmail_compose(str(_fake_exe), "eml", str(_eml),
                                            _rcpts, "s", "p")
        check("open_foxmail_compose 通过 subprocess.Popen 启动",
              len(_spawned) == 1 and _spawned[0][0] == [str(_fake_exe), str(_eml)],
              str(_spawned))
        check("open_foxmail_compose 返回实际使用的 argv", _ret_argv == _spawned[0][0])
        check("唤起 Foxmail 传入的是参数列表而非字符串（防命令注入）",
              isinstance(_spawned[0][0], list), str(type(_spawned[0][0])))
        check("唤起 Foxmail 未使用 shell=True",
              _spawned[0][1].get("shell") is not True, str(_spawned[0][1]))

        # 启动失败（如路径里有非法字符/权限不足）要转成 FoxmailError
        def _boom(argv, **kw):
            raise OSError("模拟启动失败")
        fx.subprocess.Popen = _boom  # type: ignore
        try:
            fx.open_foxmail_compose(str(_fake_exe), "eml", str(_eml), _rcpts, "s", "p")
            check("启动报错转成 FoxmailError", False, "竟然没报错")
        except fx.FoxmailError as _e:
            check("启动报错转成 FoxmailError", "启动 Foxmail 失败" in str(_e))
    finally:
        fx.subprocess.Popen = _orig_popen  # type: ignore

    try:
        fx.open_foxmail_compose(str(_other_exe), "eml", str(_eml), _rcpts, "s", "p")
        check("open_foxmail_compose 拦住无效 exe", False, "竟然没报错")
    except fx.FoxmailError:
        check("open_foxmail_compose 拦住无效 exe", True)

    # --- 14.6 描述文案 ---
    check("summary_line 含路径与唤起方式",
          "已集成" in fx.summary_line(str(_fake_exe), "eml")
          and "eml" in fx.summary_line(str(_fake_exe), "eml"))
    check("summary_line 未集成时明说未集成",
          "未集成" in fx.summary_line("", "auto"), fx.summary_line("", "auto"))
    check("describe_foxmail 对有效路径给出肯定描述",
          "已集成" in fx.describe_foxmail(str(_fake_exe)))
    # detected="" 表示"调用方已扫过且没有"，此时应给出"未检测到 + 下一步"，
    # 且不得再返回任何"已检测到"字样（保证提示与实际状态一致）
    _notfound = fx.describe_foxmail(r"C:\nope\Foxmail.exe", detected="")
    check("describe_foxmail(detected='') 走未找到分支并给出下一步",
          "未检测到" in _notfound and "未找到主程序" not in _notfound
          and len(_notfound) > 15, _notfound[:90])
    check("describe_foxmail(detected='') 不会谎报已检测到",
          "已集成" not in _notfound and "已自动检测到" not in _notfound, _notfound[:90])
    # 不给 detected 时允许自行探测（结果随机器而定，只要求类型与长度合理）
    _self = fx.describe_foxmail("")
    check("describe_foxmail 未传检测结果时会自行探测",
          isinstance(_self, str) and len(_self) > 8, _self[:90])

check("FOXMAIL_MODES 含 auto/eml/mailto 三种",
      set(fx.FOXMAIL_MODES) == {"auto", "eml", "mailto"}, str(fx.FOXMAIL_MODES))

# 真实环境自检（结果随机器而定，只要求类型正确、不抛异常、且与探测结果自洽）
_detected = fx.detect_foxmail()
check("真实环境 detect_foxmail 返回字符串", isinstance(_detected, str))
if _detected:
    check("真实环境探测到的路径确实有效", fx.is_valid_foxmail_exe(_detected) is True, _detected)
    check("真实环境 describe_foxmail 与探测结果自洽",
          "已" in fx.describe_foxmail(_detected), fx.describe_foxmail(_detected)[:60])
    check("真实环境 resolve('') 也能拿到它",
          fx.resolve("")[0] == _detected, fx.resolve("")[0])
print(f"     [本机实测] Foxmail 主程序 = {_detected or '(未检测到)'}")

# --- 14.7 配置项 foxmail_path / foxmail_mode ---
with tempfile.TemporaryDirectory() as _td_fx:
    _orig_cfg_fx = _cs.CONFIG_PATH
    _cs.CONFIG_PATH = Path(_td_fx) / "config.json"
    try:
        _cs.save_config(_cs.AppConfig(foxmail_path=r"D:\Foxmail\Foxmail.exe",
                                      foxmail_mode="eml"))
        _cf = _cs.load_config()
        check("foxmail_path 可持久化并读回",
              _cf.foxmail_path == r"D:\Foxmail\Foxmail.exe", _cf.foxmail_path)
        check("foxmail_mode 可持久化并读回", _cf.foxmail_mode == "eml", _cf.foxmail_mode)

        _cs.CONFIG_PATH.write_text('{"excel_path": "old.xlsx"}', encoding="utf-8")
        _cf = _cs.load_config()
        check("旧配置无 foxmail 字段时默认空路径 + auto",
              _cf.foxmail_path == "" and _cf.foxmail_mode == "auto",
              f"{_cf.foxmail_path!r} / {_cf.foxmail_mode!r}")

        _cs.CONFIG_PATH.write_text('{"foxmail_mode": "smoke-signal"}', encoding="utf-8")
        check("foxmail_mode 被改坏时回落到 auto", _cs.load_config().foxmail_mode == "auto")

        _cs.CONFIG_PATH.write_text('{"foxmail_path": "  D:\\\\Foxmail\\\\Foxmail.exe  "}',
                                   encoding="utf-8")
        check("foxmail_path 读取时去除首尾空白",
              _cs.load_config().foxmail_path == r"D:\Foxmail\Foxmail.exe",
              repr(_cs.load_config().foxmail_path))

        check("FOXMAIL_MODES 不在配置字段外泄",
              "FOXMAIL_MODES" not in _json.dumps(_cs.asdict(_cs.AppConfig())),
              "FOXMAIL_MODES 被当成数据字段序列化了")
    finally:
        _cs.CONFIG_PATH = _orig_cfg_fx

print(f"\n==== 结果：{passed} 通过 / {failed} 失败 ====")
sys.exit(1 if failed else 0)
