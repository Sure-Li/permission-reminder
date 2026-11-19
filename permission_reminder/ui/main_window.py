# -*- coding: utf-8 -*-
"""主窗口模块。

职责：装配全部模块（配置、Excel、工作日、检查、邮件、调度），
提供 tkinter 图形界面。

线程安全设计（tkinter 铁律：只在主线程操作控件）：
- 所有业务动作（读 Excel、发邮件）在 daemon 工作线程中执行；
- 工作线程通过 queue.Queue 把"日志行"与"UI 动作"投递给主线程；
- 主线程用 root.after(200, ...) 轮询两个队列刷新日志窗口 / 弹窗 / 按钮状态。
"""
from __future__ import annotations

import logging
import queue
import re
import threading
import tkinter as tk
from datetime import date, datetime, timedelta
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Any, Dict, List, Optional

from ..core.checker import (
    PERM_LABELS, ReminderItem, SentRecordStore, find_due_items,
)
from ..core.excel_reader import ClientRecord, ExcelFormatError, read_clients
from ..core.holiday_sync import (
    HolidayStore, month_window_around, package_special_days, sync_window_years,
)
from ..core.scheduler import SchedulerThread
from ..core.workday import create_calendar
from ..infra.config_store import (
    AppConfig, ScheduleConfig, SmtpConfig, clear_config, load_config,
    save_config, DATA_DIR,
)
from ..infra.log_setup import setup_logging
from ..mailer.diagnose import diagnose_smtp, format_report
from ..mailer.foxmail import (
    FOXMAIL_MODES, FoxmailError, open_default_apps_settings,
    open_foxmail_compose, resolve as resolve_foxmail, summary_line as foxmail_summary,
)
from ..mailer.outbox import (
    describe_mail_client, export_mail, has_mail_client, open_folder,
    open_in_mail_client, outbox_dir,
)
from ..mailer.sender import (
    SMTP_PRESETS, Attachment, MailSendError, EmailSender, auth_block_remaining,
    build_excel_attachment, find_invalid_recipients, render_summary_email, reset_auth_block,
    send_test_email,
)
from .diag_window import show_diag_report
from .holiday_window import show_holiday_plan

logger = logging.getLogger(__name__)

# 发送记录保留天数（超过的清理）
SENT_RECORD_KEEP_DAYS = 60


class MainWindow(tk.Tk):
    """应用主窗口。"""

    def __init__(self) -> None:
        super().__init__()
        self.title("权限到期提醒工具")
        self.minsize(760, 640)

        # ---- 跨线程通信队列 ----
        self._log_queue: "queue.Queue[str]" = queue.Queue()
        self._ui_queue: "queue.Queue[Dict[str, Any]]" = queue.Queue()

        # ---- 初始化日志（GUI 桥 -> 日志队列）----
        setup_logging(gui_callback=self._log_queue.put)
        logger.info("程序启动")

        # ---- 业务依赖 ----
        self.config: AppConfig = load_config()
        self._sent_store = SentRecordStore(DATA_DIR / "sent_records.json")
        pruned = self._sent_store.prune_before(date.today() - timedelta(days=SENT_RECORD_KEEP_DAYS))
        if pruned:
            logger.info("已清理 %d 条过期发送记录", pruned)

        # 调度配置快照（供调度线程读取；tkinter 变量不能跨线程访问）
        self._schedule_state = ScheduleConfig(
            enabled=self.config.schedule.enabled,
            hour=self.config.schedule.hour,
            minute=self.config.schedule.minute,
        )
        self._scheduler = SchedulerThread(
            get_config=lambda: self._schedule_state,
            run_check=self._request_check,  # 线程安全入口，见下
        )
        # 检查流水线互斥标志（防手动/定时/补检并发导致重复发送）
        self._check_busy = False
        # 窗口关闭标志（停止队列轮询，避免操作已销毁控件）
        self._closing = False

        # ---- 界面 ----
        self._build_vars()
        self._build_ui()
        self._load_into_ui()

        # ---- 启动调度与轮询 ----
        self._scheduler.start()
        self.after(200, self._poll_queues)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # ---- 启动补检（仅当定时任务处于开启状态）----
        if self._schedule_state.enabled:
            logger.info("定时任务处于开启状态，启动时执行一次补检")
            self.after(600, self._request_check)

    # ================================================================ 界面构建

    def _build_vars(self) -> None:
        """创建 tkinter 变量（仅主线程访问）。"""
        self._var_server = tk.StringVar()
        self._var_port = tk.IntVar()
        self._var_ssl = tk.StringVar(value="SSL")
        self._var_username = tk.StringVar()
        self._var_password = tk.StringVar()
        self._var_recipients = tk.StringVar()
        self._var_excel_path = tk.StringVar()
        self._var_hour = tk.IntVar()
        self._var_minute = tk.IntVar()
        self._var_perm_checks = {p: tk.BooleanVar() for p in PERM_LABELS}
        self._var_preset = tk.StringVar()
        self._var_auto_sync = tk.BooleanVar(value=True)  # 检查前自动同步节假日
        self._var_send_mode = tk.StringVar(value="smtp")  # 发送方式：smtp / client
        # Foxmail 集成（兜底通道优先用它投递）
        self._var_foxmail_path = tk.StringVar()
        self._var_foxmail_mode = tk.StringVar(value="auto")

    def _build_ui(self) -> None:
        """构建界面布局。"""
        container = ttk.Frame(self, padding=10)
        container.pack(fill=tk.BOTH, expand=True)

        self._build_smtp_frame(container)
        self._build_excel_frame(container)
        self._build_perm_frame(container)
        self._build_holiday_frame(container)
        self._build_schedule_frame(container)
        self._build_action_frame(container)
        self._build_log_frame(container)

    def _build_smtp_frame(self, parent: ttk.Frame) -> None:
        """① 邮箱配置区。"""
        frm = ttk.LabelFrame(parent, text="邮箱配置与发送方式", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))

        # 第一行：预设 + 服务器 + 端口
        row1 = ttk.Frame(frm); row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="常用预设:").pack(side=tk.LEFT)
        preset_cb = ttk.Combobox(row1, textvariable=self._var_preset,
                                 values=list(SMTP_PRESETS.keys()),
                                 state="readonly", width=14)
        preset_cb.pack(side=tk.LEFT, padx=(4, 12))
        preset_cb.bind("<<ComboboxSelected>>", self._on_preset_selected)
        ttk.Label(row1, text="SMTP服务器:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self._var_server, width=22).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(row1, text="端口:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self._var_port, width=7).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(row1, text="加密:").pack(side=tk.LEFT)
        ttk.Radiobutton(row1, text="SSL", variable=self._var_ssl,
                        value="SSL").pack(side=tk.LEFT)
        ttk.Radiobutton(row1, text="STARTTLS", variable=self._var_ssl,
                        value="STARTTLS").pack(side=tk.LEFT)
        ttk.Radiobutton(row1, text="不加密", variable=self._var_ssl,
                        value="PLAIN").pack(side=tk.LEFT)

        # 第二行：发件人 + 授权码
        row2 = ttk.Frame(frm); row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="发件人邮箱:").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self._var_username, width=28).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Label(row2, text="授权码/密码:").pack(side=tk.LEFT)
        ttk.Entry(row2, textvariable=self._var_password, width=28, show="*").pack(side=tk.LEFT, padx=4)

        # 第三行：收件人 + 按钮
        row3 = ttk.Frame(frm); row3.pack(fill=tk.X, pady=2)
        ttk.Label(row3, text="目标邮箱:").pack(side=tk.LEFT)
        ttk.Entry(row3, textvariable=self._var_recipients, width=52).pack(side=tk.LEFT, padx=(4, 12))
        ttk.Button(row3, text="保存邮箱配置", command=self._on_save_smtp).pack(side=tk.LEFT, padx=3)
        ttk.Button(row3, text="清除邮箱配置", command=self._on_clear_smtp).pack(side=tk.LEFT, padx=3)
        self._btn_test = ttk.Button(row3, text="发送测试邮件", command=self._on_send_test)
        self._btn_test.pack(side=tk.LEFT, padx=3)
        self._btn_diag = ttk.Button(row3, text="连接诊断", command=self._on_diagnose)
        self._btn_diag.pack(side=tk.LEFT, padx=3)

        # 第四行：发送方式（SMTP 直发 / 兜底走本机邮件客户端）
        row4 = ttk.Frame(frm); row4.pack(fill=tk.X, pady=2)
        ttk.Label(row4, text="发送方式:").pack(side=tk.LEFT)
        ttk.Radiobutton(row4, text="SMTP 直发", variable=self._var_send_mode,
                        value="smtp", command=self._on_send_mode_changed).pack(side=tk.LEFT, padx=(4, 8))
        ttk.Radiobutton(row4, text="本机邮件客户端（导出后由 Foxmail 等发送）",
                        variable=self._var_send_mode, value="client",
                        command=self._on_send_mode_changed).pack(side=tk.LEFT, padx=(0, 12))
        ttk.Button(row4, text="导出邮件到本地", command=self._on_export_mail).pack(side=tk.LEFT, padx=3)
        ttk.Button(row4, text="打开导出目录", command=self._on_open_outbox).pack(side=tk.LEFT, padx=3)

        # 第五行：Foxmail 集成（路径 + 自动检测 + 浏览）
        row5 = ttk.Frame(frm); row5.pack(fill=tk.X, pady=2)
        ttk.Label(row5, text="Foxmail:").pack(side=tk.LEFT)
        ttk.Entry(row5, textvariable=self._var_foxmail_path, width=44).pack(
            side=tk.LEFT, padx=(4, 6), fill=tk.X, expand=True)
        ttk.Button(row5, text="自动检测", command=self._on_detect_foxmail).pack(side=tk.LEFT, padx=3)
        ttk.Button(row5, text="浏览…", command=self._on_browse_foxmail).pack(side=tk.LEFT, padx=3)
        ttk.Button(row5, text="测试唤起", command=self._on_test_foxmail).pack(side=tk.LEFT, padx=3)

        # 第六行：唤起方式 + 设为默认邮件程序 + 状态
        row6 = ttk.Frame(frm); row6.pack(fill=tk.X, pady=2)
        ttk.Label(row6, text="唤起方式:").pack(side=tk.LEFT)
        for value, label in (("auto", "自动（推荐）"),
                             ("eml", "打开导出邮件(.eml)"),
                             ("mailto", "新建写信窗口(mailto)")):
            ttk.Radiobutton(row6, text=label, variable=self._var_foxmail_mode, value=value,
                            command=self._on_foxmail_mode_changed).pack(side=tk.LEFT, padx=(4, 6))
        ttk.Button(row6, text="设为默认邮件程序",
                   command=self._on_set_default_mail_app).pack(side=tk.LEFT, padx=(10, 3))
        self._lbl_foxmail_status = ttk.Label(frm, text="", foreground="gray", wraplength=900,
                                             justify=tk.LEFT)
        self._lbl_foxmail_status.pack(anchor=tk.W, pady=(2, 0))

        ttk.Label(frm, text="多个目标邮箱请用英文逗号分隔，例如：a@xx.com,b@xx.com；"
                            "发送失败或连不上时点「连接诊断」定位原因",
                  foreground="gray").pack(anchor=tk.W, pady=(2, 0))
        self._lbl_send_hint = ttk.Label(frm, text="", foreground="gray", wraplength=900,
                                        justify=tk.LEFT)
        self._lbl_send_hint.pack(anchor=tk.W, pady=(0, 0))

    def _build_excel_frame(self, parent: ttk.Frame) -> None:
        """② Excel 文件区。"""
        frm = ttk.LabelFrame(parent, text="Excel 数据文件", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))
        row = ttk.Frame(frm); row.pack(fill=tk.X)
        ttk.Label(row, text="文件路径:").pack(side=tk.LEFT)
        ttk.Entry(row, textvariable=self._var_excel_path, width=62).pack(
            side=tk.LEFT, padx=4, fill=tk.X, expand=True)
        ttk.Button(row, text="选择 Excel 文件", command=self._on_choose_excel).pack(side=tk.LEFT, padx=3)
        ttk.Button(row, text="清除设置文件", command=self._on_clear_config).pack(side=tk.LEFT, padx=3)

    def _build_perm_frame(self, parent: ttk.Frame) -> None:
        """③ 检查范围区。"""
        frm = ttk.LabelFrame(parent, text="检查范围（勾选要检查的权限到期日期列）", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))
        for perm in PERM_LABELS:
            ttk.Checkbutton(frm, text=f"{perm}到期日期",
                            variable=self._var_perm_checks[perm]).pack(side=tk.LEFT, padx=14)

    def _build_holiday_frame(self, parent: ttk.Frame) -> None:
        """④ 节假日区：同步最新节假日安排并查看休假计划。"""
        frm = ttk.LabelFrame(parent, text="节假日（工作日计算依据）", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))
        row = ttk.Frame(frm); row.pack(fill=tk.X)
        self._btn_sync = ttk.Button(row, text="同步最新节假日",
                                    command=self._on_sync_holidays)
        self._btn_sync.pack(side=tk.LEFT, padx=3)
        ttk.Button(row, text="查看休假计划", command=self._on_view_holidays).pack(
            side=tk.LEFT, padx=3)
        self._chk_auto_sync = ttk.Checkbutton(
            row, text="检查前自动同步", variable=self._var_auto_sync,
            command=self._on_toggle_auto_sync)
        self._chk_auto_sync.pack(side=tk.LEFT, padx=(14, 3))
        self._lbl_holiday_status = ttk.Label(row, text="从未同步", foreground="gray")
        self._lbl_holiday_status.pack(side=tk.LEFT, padx=10)
        ttk.Label(frm, text="开启后，每次检查（立即检查 / 定时 / 启动补检）前都会先拉取最新"
                            "节假日安排；同步失败仅记日志，不影响本次检查，自动沿用本地缓存。",
                  foreground="gray").pack(anchor=tk.W, pady=(2, 0))

    def _build_schedule_frame(self, parent: ttk.Frame) -> None:
        """④ 定时任务区。"""
        frm = ttk.LabelFrame(parent, text="定时任务（每天固定时间检查）", padding=8)
        frm.pack(fill=tk.X, pady=(0, 6))
        row = ttk.Frame(frm); row.pack(fill=tk.X)
        ttk.Label(row, text="每天").pack(side=tk.LEFT)
        sp_hour = ttk.Spinbox(row, from_=0, to=23, width=4, textvariable=self._var_hour,
                              command=self._on_schedule_time_changed)
        sp_hour.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(row, text="时").pack(side=tk.LEFT, padx=(2, 6))
        sp_min = ttk.Spinbox(row, from_=0, to=59, width=4, textvariable=self._var_minute,
                             command=self._on_schedule_time_changed)
        sp_min.pack(side=tk.LEFT, padx=(4, 0))
        ttk.Label(row, text="分").pack(side=tk.LEFT, padx=(2, 12))
        sp_hour.bind("<KeyRelease>", lambda _e: self._on_schedule_time_changed())
        sp_min.bind("<KeyRelease>", lambda _e: self._on_schedule_time_changed())
        self._btn_schedule = ttk.Button(row, text="启动定时任务", command=self._on_toggle_schedule)
        self._btn_schedule.pack(side=tk.LEFT, padx=3)
        self._lbl_schedule_status = ttk.Label(row, text="未启动", foreground="gray")
        self._lbl_schedule_status.pack(side=tk.LEFT, padx=10)

    def _build_action_frame(self, parent: ttk.Frame) -> None:
        """⑤ 操作区。"""
        frm = ttk.Frame(parent)
        frm.pack(fill=tk.X, pady=(0, 6))
        self._btn_check = ttk.Button(frm, text="立即检查并发送提醒", command=self._on_check_now)
        self._btn_check.pack(side=tk.LEFT)

    def _build_log_frame(self, parent: ttk.Frame) -> None:
        """⑥ 日志窗口区。"""
        frm = ttk.LabelFrame(parent, text="运行日志", padding=4)
        frm.pack(fill=tk.BOTH, expand=True)
        self._txt_log = tk.Text(frm, height=12, state=tk.DISABLED,
                                wrap=tk.NONE, font=("Consolas", 9))
        scroll_y = ttk.Scrollbar(frm, orient=tk.VERTICAL, command=self._txt_log.yview)
        scroll_x = ttk.Scrollbar(frm, orient=tk.HORIZONTAL, command=self._txt_log.xview)
        self._txt_log.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self._txt_log.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        frm.rowconfigure(0, weight=1)
        frm.columnconfigure(0, weight=1)

    # ============================================================ 配置 <-> 界面

    def _load_into_ui(self) -> None:
        """把配置加载到界面控件。"""
        c = self.config
        self._var_server.set(c.smtp.server)
        self._var_port.set(c.smtp.port)
        self._var_ssl.set(c.smtp.ssl_mode)
        self._var_username.set(c.smtp.username)
        self._var_password.set(c.smtp.password)
        self._var_recipients.set(", ".join(c.smtp.recipients))
        self._var_excel_path.set(c.excel_path)
        self._var_hour.set(c.schedule.hour)
        self._var_minute.set(c.schedule.minute)
        for perm, var in self._var_perm_checks.items():
            var.set(bool(c.checked_perms.get(perm, True)))
        self._var_auto_sync.set(bool(c.auto_sync_holidays))
        self._var_send_mode.set(c.send_mode if c.send_mode in AppConfig.SEND_MODES else "smtp")
        self._var_foxmail_path.set(c.foxmail_path)
        self._var_foxmail_mode.set(c.foxmail_mode if c.foxmail_mode in FOXMAIL_MODES else "auto")
        self._refresh_schedule_ui()
        self._refresh_holiday_status()
        self._refresh_send_mode_ui()
        self._refresh_foxmail_status(auto_detect=True)

    def _gather_config(self) -> AppConfig:
        """从界面控件收集配置（仅主线程调用）。"""
        try:
            port = int(self._var_port.get())
        except (tk.TclError, ValueError):
            port = 465
        try:
            hour = max(0, min(23, int(self._var_hour.get())))
        except (tk.TclError, ValueError):
            hour = 9
        try:
            minute = max(0, min(59, int(self._var_minute.get())))
        except (tk.TclError, ValueError):
            minute = 0
        return AppConfig(
            smtp=SmtpConfig(
                server=self._var_server.get().strip(),
                port=port,
                ssl_mode=self._var_ssl.get(),
                username=self._var_username.get().strip(),
                password=self._var_password.get(),
                recipients=self._parse_recipients(self._var_recipients.get()),
            ),
            excel_path=self._var_excel_path.get().strip(),
            schedule=ScheduleConfig(
                enabled=self._schedule_state.enabled,
                hour=hour, minute=minute,
            ),
            checked_perms={p: v.get() for p, v in self._var_perm_checks.items()},
            auto_sync_holidays=bool(self._var_auto_sync.get()),
            send_mode=self._var_send_mode.get() or "smtp",
            foxmail_path=self._var_foxmail_path.get().strip(),
            foxmail_mode=(self._var_foxmail_mode.get()
                          if self._var_foxmail_mode.get() in FOXMAIL_MODES else "auto"),
        )

    @staticmethod
    def _parse_recipients(text: str) -> List[str]:
        """解析收件人输入：兼容中英文逗号 / 分号 / 空白分隔。"""
        parts = re.split(r"[,，;；\s]+", text.strip())
        return [p for p in parts if "@" in p]

    # ================================================================ 事件处理

    def _on_preset_selected(self, event=None) -> None:
        """选择 SMTP 预设。"""
        # 防御：仅当用户确实与下拉框交互（其持有焦点）时套用，
        # 过滤 Tk 偶发的虚拟事件误触发，避免覆盖用户已填的自定义服务器
        if event is not None and event.widget is not None:
            try:
                if event.widget.focus_get() is not event.widget:
                    return
            except tk.TclError:
                return
        preset = SMTP_PRESETS.get(self._var_preset.get())
        if preset:
            server, port, ssl_mode = preset
            self._var_server.set(server)
            self._var_port.set(port)
            self._var_ssl.set(ssl_mode)
            logger.info("已套用 SMTP 预设：%s（%s:%d %s）",
                        self._var_preset.get(), server, port, ssl_mode)

    def _on_save_smtp(self) -> None:
        """保存邮箱配置（含目标邮箱）到本地。"""
        cfg = self._gather_config()
        if not cfg.smtp.server or not cfg.smtp.username or not cfg.smtp.password:
            messagebox.showwarning("提示", "请完整填写 SMTP 服务器、发件人邮箱和授权码")
            return
        if not cfg.smtp.recipients:
            messagebox.showwarning("提示", "请至少填写一个目标邮箱（多个用英文逗号分隔）")
            return
        # 收件人格式预检：地址写错时邮件服务器必然拒收，提前提醒避免无效重试
        invalid = find_invalid_recipients(cfg.smtp.recipients)
        if invalid:
            messagebox.showwarning(
                "目标邮箱格式可能有误",
                "以下目标邮箱格式不正确，发送时会被邮件服务器拒绝：\n  · "
                + "\n  · ".join(invalid)
                + "\n\n常见原因：漏写 @、漏写域名后缀（如 .com 少了一个点）、"
                  "或把“，”当成了英文逗号。\n配置已保存，建议修正后重新保存。")
        self.config = cfg
        save_config(cfg)
        # 用户已修改并保存邮箱配置 → 解除认证失败后的发送冷静期，允许立即重试
        reset_auth_block()
        messagebox.showinfo("成功", f"邮箱配置已保存到本地。\n收件人：{', '.join(cfg.smtp.recipients)}")

    def _on_clear_smtp(self) -> None:
        """清除邮箱配置（仅 SMTP / 收件人，保留 Excel 路径、定时与勾选状态）。"""
        if not messagebox.askyesno(
            "确认", "确定要清除邮箱配置吗？\n"
            "（SMTP 服务器、发件人、授权码、目标邮箱将清空；\n"
            "Excel 路径、定时任务、检查范围保持不变）",
        ):
            return
        self.config.smtp = SmtpConfig()
        save_config(self.config)
        self._load_into_ui()
        logger.info("邮箱配置已清除（其他设置保持不变）")
        messagebox.showinfo("完成", "邮箱配置已清除")

    def _on_choose_excel(self) -> None:
        """选择 Excel 文件并持久化路径。"""
        path = filedialog.askopenfilename(
            title="选择客户权限 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx *.xls *.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self._var_excel_path.set(path)
        self.config = self._gather_config()
        save_config(self.config)
        logger.info("已选择并保存 Excel 文件路径：%s", path)

    def _on_clear_config(self) -> None:
        """清除本地设置文件（仅配置；日志与发送记录保留）。"""
        if not messagebox.askyesno(
            "确认", "确定要清除本地设置文件吗？\n（邮箱配置、Excel 路径、定时设置将清空；"
            "日志和发送记录保留）",
        ):
            return
        clear_config()
        self.config = AppConfig()
        self._schedule_state.enabled = False
        self._load_into_ui()
        logger.info("本地设置文件已清除，界面已重置为默认值")
        messagebox.showinfo("完成", "本地设置文件已清除")

    def _on_schedule_time_changed(self) -> None:
        """时 / 分控件变化：同步调度快照并保存。"""
        try:
            self._schedule_state.hour = max(0, min(23, int(self._var_hour.get())))
            self._schedule_state.minute = max(0, min(59, int(self._var_minute.get())))
        except (tk.TclError, ValueError):
            return
        self.config = self._gather_config()
        save_config(self.config)
        self._refresh_schedule_ui()

    def _on_toggle_schedule(self) -> None:
        """启动 / 停止定时任务。"""
        try:
            self._schedule_state.hour = max(0, min(23, int(self._var_hour.get())))
            self._schedule_state.minute = max(0, min(59, int(self._var_minute.get())))
        except (tk.TclError, ValueError):
            messagebox.showwarning("提示", "时 / 分输入无效")
            return
        self._schedule_state.enabled = not self._schedule_state.enabled
        self.config = self._gather_config()
        save_config(self.config)
        self._refresh_schedule_ui()
        if self._schedule_state.enabled:
            logger.info("定时任务已启动：每天 %02d:%02d 检查（重启程序后自动恢复）",
                        self._schedule_state.hour, self._schedule_state.minute)
        else:
            logger.info("定时任务已停止")

    def _refresh_schedule_ui(self) -> None:
        """刷新定时任务按钮与状态显示。"""
        if self._schedule_state.enabled:
            self._btn_schedule.config(text="停止定时任务")
            self._lbl_schedule_status.config(
                text=f"已启动：每天 {self._schedule_state.hour:02d}:"
                     f"{self._schedule_state.minute:02d} 检查",
                foreground="#1a7a34")
        else:
            self._btn_schedule.config(text="启动定时任务")
            self._lbl_schedule_status.config(text="未启动", foreground="gray")

    # ================================================================ 发送方式

    def _on_send_mode_changed(self) -> None:
        """切换发送方式：刷新提示并立即持久化（避免用户忘记保存）。"""
        self._refresh_send_mode_ui()
        cfg = self._gather_config()
        try:
            save_config(cfg)
        except OSError as exc:
            logger.warning("保存发送方式失败：%s", exc)
            return
        logger.info("发送方式已切换为：%s",
                    "本机邮件客户端（导出兜底）" if cfg.send_mode == "client" else "SMTP 直发")

    def _refresh_send_mode_ui(self) -> None:
        """按当前发送方式刷新提示文案。

        兜底模式不经过 SMTP 认证，因此**不再要求填写授权码/密码**——
        这正是它能在"网关拒绝第三方客户端"时救急的原因。
        提示里会说明**实际会用哪个客户端**投递（优先 Foxmail）。
        """
        if self._var_send_mode.get() == "client":
            self._lbl_send_hint.config(
                text=self._compose_client_hint(),
                foreground="#8a5300")
            self._btn_check.config(text="立即检查并导出提醒")
        else:
            self._lbl_send_hint.config(
                text="当前为「SMTP 直发」模式：由本工具直连 SMTP 服务器发送。"
                     "若提示 535 认证被拒，可切到「本机邮件客户端」模式应急。",
                foreground="gray")
            self._btn_check.config(text="立即检查并发送提醒")

    def _compose_client_hint(self) -> str:
        """拼装兜底模式的提示文案（按 Foxmail / 系统客户端 / 都没有 三种情况）。

        单独抽出来是为了让"提示说的话"和"检查时实际做的事"共用同一份判断，
        避免出现"提示说会唤起 Foxmail，实际什么都没弹"的不一致。
        """
        head = ("当前为「本机邮件客户端」模式：检查时会把提醒导出到 data\\outbox\\"
                "（.eml + .html），再交给邮件客户端投递。此模式不需要授权码/密码，"
                "只需填好「目标邮箱」。适用于单位邮件网关拒绝第三方客户端认证"
                "（如 535 reject by CACTER gateway）的情况。\n")
        tail = "提醒含客户号与姓名，请务必用单位邮箱发送。"

        foxmail_exe, _note = resolve_foxmail(self._var_foxmail_path.get().strip())
        if foxmail_exe:
            mode = self._var_foxmail_mode.get()
            if mode not in FOXMAIL_MODES:
                mode = "auto"
            return head + (f"将用 Foxmail 投递：{foxmail_exe}\n"
                           f"{foxmail_summary(foxmail_exe, mode)}\n"
                           "检查时会唤起 Foxmail（收件人/主题/正文已填好），"
                           "请在其中点「发送」完成投递。\n") + tail

        client = describe_mail_client() or "未知"
        if has_mail_client():
            return head + (f"未检测到 Foxmail，将改用系统默认邮件客户端：{client}\n"
                           "检查时会唤起它的新建邮件窗口（收件人/主题/正文摘要已预填），"
                           "请在其中点「发送」完成投递。\n"
                           "建议安装 Foxmail 并配好单位邮箱账号后点「自动检测」做集成。\n") + tail
        return head + (f"本机既未检测到 Foxmail，也没有可用的默认邮件客户端"
                       f"（当前：{client}）。\n"
                       "检查时【不会唤起任何窗口】，而是直接打开导出目录——"
                       "请先用 Foxmail 双击打开其中的 .eml 再点「发送」。\n"
                       "若尚未安装 Foxmail，请先安装并配好单位邮箱账号，"
                       "然后点上面的「自动检测」完成集成。\n") + tail

    # ============================================================ Foxmail 集成

    def _refresh_foxmail_status(self, auto_detect: bool = False) -> None:
        """刷新 Foxmail 状态行。

        Args:
            auto_detect: True 时若界面路径为空且未检测到，则把自动探测结果
                         回填到输入框，省去用户手工填写。
        """
        if not hasattr(self, "_lbl_foxmail_status"):
            return
        typed = self._var_foxmail_path.get().strip()
        exe, note = resolve_foxmail(typed)
        if exe and not typed:
            # 自动探测到了：回填到输入框，让用户看得见也能改
            self._var_foxmail_path.set(exe)
        color = "#1a7f37" if exe else "#8a5300"
        self._lbl_foxmail_status.config(text=note, foreground=color)
        if auto_detect and exe:
            logger.info("Foxmail 集成可用：%s", exe)

    def _on_detect_foxmail(self) -> None:
        """自动检测 Foxmail 安装位置并保存。"""
        # 先清空旧值再探测，避免"上次填错的值"影响检测结果
        self._var_foxmail_path.set("")
        exe, note = resolve_foxmail("")
        if exe:
            self._var_foxmail_path.set(exe)
            self._refresh_foxmail_status()
            self._persist_quietly()
            logger.info("自动检测到 Foxmail：%s", exe)
            messagebox.showinfo("已集成 Foxmail",
                                f"已找到 Foxmail 并完成集成：\n{exe}\n\n"
                                "检查时若使用「本机邮件客户端」模式，将直接唤起它投递。")
        else:
            self._refresh_foxmail_status()
            messagebox.showwarning("未检测到 Foxmail", note)

    def _on_browse_foxmail(self) -> None:
        """手动指定 Foxmail.exe 路径。"""
        initial = self._var_foxmail_path.get().strip()
        initial_dir = str(Path(initial).parent) if initial and Path(initial).exists() else ""
        path = filedialog.askopenfilename(
            title="选择 Foxmail 主程序（Foxmail.exe）",
            initialdir=initial_dir or None,
            filetypes=[("Foxmail 主程序", "Foxmail.exe"), ("可执行文件", "*.exe"),
                       ("所有文件", "*.*")],
        )
        if not path:
            return
        if Path(path).name.lower() != "foxmail.exe":
            if not messagebox.askyesno(
                    "确认路径",
                    f"你选择的是：\n{path}\n\n它的文件名不是 Foxmail.exe，"
                    "确定要用它作为 Foxmail 主程序吗？"):
                return
        self._var_foxmail_path.set(path)
        self._refresh_foxmail_status()
        self._persist_quietly()
        logger.info("已手动指定 Foxmail：%s", path)
        self._refresh_send_mode_ui()

    def _on_foxmail_mode_changed(self) -> None:
        """切换 Foxmail 唤起方式后立即持久化。"""
        self._refresh_send_mode_ui()
        self._persist_quietly()

    def _on_test_foxmail(self) -> None:
        """测试唤起 Foxmail：用一封临时邮件验证集成是否真的可用。

        刻意用真实的 .eml 做测试（而不是空参数），因为默认方式就是"打开导出的邮件"；
        测试邮件只写进 data/outbox/，**不会真的发出去**。
        """
        exe, note = resolve_foxmail(self._var_foxmail_path.get().strip())
        if not exe:
            messagebox.showwarning("未集成 Foxmail", note)
            return
        mode = self._var_foxmail_mode.get()
        recipients = self._parse_recipients(self._var_recipients.get())
        subject = "【权限到期提醒】Foxmail 集成测试"
        plain = ("这是一封用于验证「权限到期提醒工具」与 Foxmail 集成的测试邮件。"
                 "如果你在 Foxmail 里看到了这封邮件（或它的写信窗口），说明集成可用。"
                 "本测试不会自动发送，请直接关闭即可。")
        html = ("<html><body><p>这是一封用于验证「权限到期提醒工具」与 Foxmail "
                "集成的测试邮件，不会自动发送。</p>"
                "<p>在 Foxmail 中看到它（或它的写信窗口），即表示集成可用。</p></body></html>")
        try:
            eml, _html = export_mail(recipients, subject, html, plain,
                                     from_addr=self._var_username.get().strip(),
                                     tag="Foxmail集成测试")
        except OSError as exc:
            messagebox.showerror("导出失败", f"无法写入导出目录：{exc}")
            return
        try:
            argv = open_foxmail_compose(exe, mode, str(eml), recipients, subject, plain)
        except FoxmailError as exc:
            logger.error("测试唤起 Foxmail 失败：%s", exc)
            self._refresh_foxmail_status()
            messagebox.showerror(
                "唤起 Foxmail 失败",
                f"{exc}\n\n已导出的测试邮件在这里：\n{eml}\n\n"
                "可以双击它手工验证；若「唤起方式」选的是 mailto 而失败，"
                "请改用「自动」或「打开导出邮件(.eml)」。")
            try:
                open_folder(eml)
            except OSError:
                pass
            return
        logger.info("Foxmail 集成测试已唤起：%s", argv)
        messagebox.showinfo(
            "已唤起 Foxmail",
            "已调用 Foxmail，请看是否出现了这封邮件（或其写信窗口）。\n\n"
            f"  · 主程序：{exe}\n"
            f"  · 方式：{foxmail_summary(exe, mode)}\n"
            f"  · 测试邮件：{eml}\n\n"
            "若 Foxmail 没有反应，请把「唤起方式」换成另一种再试一次。")

    def _on_set_default_mail_app(self) -> None:
        """打开 Windows「默认应用」设置，引导把 Foxmail 设为默认邮件程序。"""
        if not open_default_apps_settings():
            messagebox.showinfo("请手动设置",
                                "请在「设置 → 应用 → 默认应用 → 按协议选择默认应用」中，"
                                "把 MAILTO 指定为 Foxmail。")
            return
        messagebox.showinfo(
            "已打开系统设置",
            "已在系统设置中打开「默认应用」页。\n\n"
            "把默认邮件客户端设为 Foxmail（可在搜索框输入 MAILTO 直接定位），"
            "之后本工具在无法唤起 Foxmail 时也能通过系统默认邮件程序投递。\n\n"
            "注：这一步由 Windows 完成，本工具不会去改注册表。")

    def _persist_quietly(self) -> None:
        """保存配置，失败只记日志（用于 Foxmail 相关的即改即存）。"""
        try:
            save_config(self._gather_config())
        except OSError as exc:
            logger.warning("保存 Foxmail 配置失败：%s", exc)

    def _on_open_outbox(self) -> None:
        """在文件管理器中打开导出目录。"""
        try:
            target = outbox_dir()
            open_folder(target)
            logger.info("已打开导出目录：%s", target)
        except OSError as exc:
            logger.error("打开导出目录失败：%s", exc)
            messagebox.showerror("打开失败", f"无法打开导出目录：{exc}")

    def _on_export_mail(self) -> None:
        """手动把当前提醒导出为本地邮件（只落盘，不发送、不标记已发）。"""
        self._run_check_threaded(export_only=True)

    def _export_and_handoff(self, cfg: AppConfig, items: List[ReminderItem],
                            subject: str, html: str, plain: str,
                            attachments: Optional[List[Attachment]] = None) -> None:
        """兜底通道：导出 .eml/.html 并唤起本机邮件客户端（工作线程中执行）。

        刻意**不在这里标记已发**：邮件最终是否发出由用户决定，
        由主线程弹窗确认后再标记，避免"其实没发出去却被当成已提醒"。
        """
        if attachments is None:
            attachments = [build_excel_attachment(items)]
        try:
            eml_path, html_path = export_mail(
                cfg.smtp.recipients, subject, html, plain,
                from_addr=cfg.smtp.username, tag="权限到期提醒",
                attachments=attachments,
            )
        except OSError as exc:
            logger.error("导出提醒邮件失败：%s", exc)
            self._ui_queue.put({
                "type": "error", "title": "导出失败",
                "message": f"无法写入导出目录：{exc}\n\n请检查 data 目录是否可写。",
            })
            return

        # 投递优先级：① 已集成的 Foxmail → ② 系统默认邮件客户端 → ③ 只打开目录
        # 之所以优先 Foxmail：它是单位要求用的客户端，账号凭据在它手里，
        # 而系统默认的可能是个人版 Outlook（用个人邮箱发客户信息不合规）。
        foxmail_exe, _note = resolve_foxmail(cfg.foxmail_path)
        opened = False
        step: Optional[str] = None
        if foxmail_exe:
            try:
                mode = "eml" if attachments else cfg.foxmail_mode
                argv = open_foxmail_compose(foxmail_exe, mode, str(eml_path),
                                            cfg.smtp.recipients, subject, plain)
                opened = True
                logger.info("已用 Foxmail 投递：%s", argv)
                step = (f"已唤起 Foxmail 打开完整邮件（{foxmail_summary(foxmail_exe, mode)}），"
                        "Excel 附件已包含在 .eml 中，请再次编辑或转发后确认附件并发送。")
            except FoxmailError as exc:
                logger.warning("唤起 Foxmail 失败，回退到系统默认客户端：%s", exc)
                step = None

        if not opened and step is None:
            client = describe_mail_client() or "未知"
            try:
                opened = open_in_mail_client(cfg.smtp.recipients, subject, plain,
                                             eml_path=eml_path if attachments else None)
            except OSError as exc:
                logger.warning("唤起本机邮件客户端失败：%s", exc)
                opened = False
            if opened:
                step = (f"未使用 Foxmail（未集成或唤起失败），"
                        f"已请求系统打开包含 Excel 附件的完整 .eml 邮件，"
                        "请确认使用单位邮件客户端，再次编辑或转发并发送。\n"
                        "提示：也可在界面「Foxmail」一栏点「自动检测」完成集成，"
                        "以后优先用 Foxmail 投递。")
            else:
                step = (f"未能自动唤起任何邮件客户端（Foxmail 未集成，"
                        f"系统默认：{client}），已为你打开导出目录，"
                        "请手动发送其中的 .eml 文件。")
                try:
                    open_folder(eml_path)
                except OSError as exc:
                    logger.warning("打开导出目录失败：%s", exc)

        self._ui_queue.put({
            "type": "confirm_export",
            "items": [(it.customer_id, it.perm, it.trigger_date) for it in items],
            "today": date.today(),
            "message": (
                f"已导出 {len(items)} 项到期提醒：\n"
                f"  · 邮件文件：{eml_path}\n"
                f"  · 正文预览：{html_path}\n\n"
                f"{step}\n\n"
                "⚠️ 提醒内容含客户号与姓名，请务必用单位邮箱发送，不要用个人邮箱。"
            ),
        })

    def _on_send_test(self) -> None:
        """发送测试邮件（工作线程）。"""
        cfg = self._gather_config()
        # 兜底模式不经 SMTP，只需目标邮箱即可"测试"（导出一封测试邮件唤起客户端）
        if cfg.send_mode == "client":
            self._on_export_test(cfg)
            return
        if not cfg.smtp.server or not cfg.smtp.username or not cfg.smtp.password:
            messagebox.showwarning("提示", "请先完整填写邮箱配置")
            return
        if not cfg.smtp.recipients:
            messagebox.showwarning("提示", "请先填写目标邮箱")
            return
        # 认证失败冷静期保护：邮件安全网关按 IP 统计认证失败并封禁，
        # 这里先拦一道，避免用户连续点"测试邮件"把出口 IP 送进黑名单
        remaining = auth_block_remaining()
        if remaining > 0:
            if not messagebox.askyesno(
                "冷静期提示",
                f"上次账号认证失败后进入了发送冷静期，还有约 "
                f"{remaining // 60} 分 {remaining % 60} 秒。\n\n"
                "邮件服务器（或其安全网关）会按来源 IP 统计认证失败次数，"
                "连续重试会导致出口 IP 被临时封禁。\n\n"
                "如果已经确认修正了密码或服务器地址，可点击「是」立即重试；\n"
                "否则建议先点「连接诊断」定位问题。",
            ):
                return
            reset_auth_block()
        self._btn_test.config(state=tk.DISABLED)

        def work() -> None:
            try:
                send_test_email(cfg.smtp)
                self._ui_queue.put({"type": "info",
                                    "title": "成功",
                                    "message": "测试邮件已发送，请到目标邮箱查收"})
                logger.info("测试邮件已发送至：%s", ", ".join(cfg.smtp.recipients))
            except MailSendError as exc:
                self._ui_queue.put({"type": "error", "title": "发送失败", "message": str(exc)})
                logger.error("测试邮件发送失败：%s", exc)
            finally:
                self._ui_queue.put({"type": "enable_test"})
        threading.Thread(target=work, daemon=True, name="test-mail").start()

    def _on_export_test(self, cfg: AppConfig) -> None:
        """兜底模式下的「测试」：导出一封测试邮件并唤起邮件客户端，全程不经 SMTP。"""
        if not cfg.smtp.recipients:
            messagebox.showwarning("提示", "请先填写目标邮箱")
            return
        invalid = find_invalid_recipients(cfg.smtp.recipients)
        if invalid:
            messagebox.showwarning("目标邮箱格式不正确",
                                   "以下地址格式不正确：\n  · " + "\n  · ".join(invalid))
            return
        self._btn_test.config(state=tk.DISABLED)

        def work() -> None:
            subject = "【权限到期提醒】测试邮件"
            html = ("<html><body><p>这是一封来自「权限到期提醒工具」的测试邮件，"
                    "用于验证<strong>本机邮件客户端</strong>这条兜底通道是否可用。</p>"
                    "<p>如果你能收到本邮件，说明用 Foxmail 发送是通的——"
                    "在单位邮件网关拒绝第三方客户端认证的情况下，"
                    "可以把「发送方式」切到「本机邮件客户端」继续使用本工具。</p>"
                    "</body></html>")
            plain = ("这是「权限到期提醒工具」的测试邮件，"
                     "用于验证本机邮件客户端这条兜底通道是否可用。")
            try:
                eml, html_path = export_mail(cfg.smtp.recipients, subject, html, plain,
                                             from_addr=cfg.smtp.username, tag="测试邮件")
                # 与正式检查保持一致：优先 Foxmail，其次是系统默认客户端
                foxmail_exe, _note = resolve_foxmail(cfg.foxmail_path)
                opened = False
                used = ""
                if foxmail_exe:
                    try:
                        open_foxmail_compose(foxmail_exe, cfg.foxmail_mode, str(eml),
                                             cfg.smtp.recipients, subject, plain)
                        opened = True
                        used = f"已唤起 Foxmail（{foxmail_summary(foxmail_exe, cfg.foxmail_mode)}）"
                    except FoxmailError as exc:
                        logger.warning("唤起 Foxmail 失败，回退系统默认客户端：%s", exc)
                if not opened:
                    client = describe_mail_client() or "未知"
                    try:
                        opened = open_in_mail_client(cfg.smtp.recipients, subject, plain)
                    except OSError as exc:
                        logger.warning("唤起本机邮件客户端失败：%s", exc)
                        opened = False
                    if opened:
                        used = f"已唤起系统默认客户端「{client}」的新建邮件窗口"
                    else:
                        used = (f"未能自动唤起（Foxmail 未集成，系统默认：{client}）")
                if not opened:
                    try:
                        open_folder(eml)
                    except OSError as exc:
                        logger.warning("打开导出目录失败：%s", exc)
                self._ui_queue.put({
                    "type": "info", "title": "已生成测试邮件",
                    "message": f"测试邮件已导出到本地：\n  · {eml}\n  · {html_path}\n\n"
                               f"{used}\n"
                               + ("请点「发送」完成投递，收到即说明这条通道可用。"
                                  if opened else
                                  "已打开导出目录，请用 Foxmail 双击该 .eml 后点「发送」。"),
                })
                logger.info("测试邮件已导出：%s", eml)
            except OSError as exc:
                self._ui_queue.put({"type": "error", "title": "导出失败", "message": str(exc)})
                logger.error("导出测试邮件失败：%s", exc)
            finally:
                self._ui_queue.put({"type": "enable_test"})
        threading.Thread(target=work, daemon=True, name="test-mail-export").start()

    def _on_check_now(self) -> None:
        """立即检查（工作线程）。"""
        self._run_check_threaded()

    # ================================================================ 连接诊断

    def _on_diagnose(self) -> None:
        """SMTP 发信链路诊断（工作线程，只读不发送邮件）。"""
        cfg = self._gather_config()
        if not cfg.smtp.server:
            messagebox.showwarning("提示", "请先填写 SMTP 服务器地址，再执行连接诊断")
            return
        self._btn_diag.config(state=tk.DISABLED, text="诊断中…")
        logger.info("开始发信链路诊断：%s:%s（%s）",
                    cfg.smtp.server, cfg.smtp.port, cfg.smtp.ssl_mode)

        def work() -> None:
            try:
                diag = diagnose_smtp(cfg.smtp)
                report = format_report(diag)
                for line in report.splitlines():
                    logger.info("诊断 | %s", line)
                self._ui_queue.put({"type": "diag_report", "report": report,
                                    "recommended": diag.recommended})
            except Exception:  # noqa: BLE001 —— 兜底，诊断自身异常也要有反馈
                logger.exception("发信链路诊断出现未预期异常")
                self._ui_queue.put({
                    "type": "error", "title": "诊断失败",
                    "message": "诊断过程出现未预期异常，详情请查看日志窗口和日志文件",
                })
            finally:
                self._ui_queue.put({"type": "enable_diag"})

        threading.Thread(target=work, daemon=True, name="diag").start()

    def _apply_recommended(self, server: str, port: int, ssl_mode: str) -> None:
        """把诊断推荐的 SMTP 配置套用到界面并保存（仅主线程调用）。"""
        self._var_server.set(server)
        self._var_port.set(port)
        self._var_ssl.set(ssl_mode)
        self.config = self._gather_config()
        save_config(self.config)
        reset_auth_block()          # 换端点后解除冷静期，允许立即验证
        logger.info("已应用诊断推荐的 SMTP 配置：%s:%d %s", server, port, ssl_mode)
        messagebox.showinfo(
            "已应用推荐配置",
            f"已把 SMTP 服务器改为：\n  {server}:{port}（{ssl_mode}）\n\n"
            "请再点「发送测试邮件」验证；\n"
            "若提示认证失败，按弹窗指引改用「客户端专用密码 / 授权码」后重试。")

    # ================================================================ 节假日功能

    def _refresh_holiday_status(self) -> None:
        """刷新节假日状态标签（显示上次同步时间与已缓存年份）。"""
        store = HolidayStore(DATA_DIR / "holidays.json")
        if store.years:
            years = "、".join(str(y) for y in sorted(store.years))
            self._lbl_holiday_status.config(
                text=f"已缓存 {years} 年数据（{store.synced_at} 同步）",
                foreground="#1a7a34")
        else:
            self._lbl_holiday_status.config(
                text="未同步（使用 chinese-calendar 内置数据）", foreground="gray")

    def _build_holiday_days(self, store: HolidayStore) -> tuple:
        """组装窗口区间的特殊日期列表（本地数据 + 包内置补充）。

        Returns:
            (days, start, end)
        """
        today = date.today()
        start, end = month_window_around(today)
        days = store.special_days_between(start, end)
        days += package_special_days(start, end, covered_years=store.years)
        days.sort(key=lambda s: s.date)
        return days, start, end

    @staticmethod
    def _window_suffix(start: date, end: date) -> str:
        return f"{start.year}年{start.month}月 ~ {end.year}年{end.month}月"

    def _on_view_holidays(self) -> None:
        """查看休假计划（只读本地缓存，不访问网络）。"""
        store = HolidayStore(DATA_DIR / "holidays.json")
        days, start, end = self._build_holiday_days(store)
        note = (f"数据来源：本地同步数据（上次同步：{store.synced_at}）+ "
                f"chinese-calendar 内置数据，共 {len(days)} 项。"
                f"如需最新数据请点击\"同步最新节假日\"。")
        show_holiday_plan(self, days, self._window_suffix(start, end), note)

    def _on_toggle_auto_sync(self) -> None:
        """切换"检查前自动同步"开关并立即持久化。"""
        self.config.auto_sync_holidays = bool(self._var_auto_sync.get())
        save_config(self.config)
        logger.info("检查前自动同步节假日：%s",
                    "已开启" if self.config.auto_sync_holidays else "已关闭")

    def _auto_sync_holidays(self) -> None:
        """检查前自动同步节假日（在检查工作线程中调用，best-effort）。

        绝不抛异常、绝不弹窗：网络不可达、官方尚未公布数据源等情况下
        仅写日志，本次检查继续使用本地缓存 / chinese_calendar 内置数据。
        """
        try:
            store = HolidayStore(DATA_DIR / "holidays.json")
            ok_years, failed = sync_window_years(store, date.today())
            if ok_years:
                logger.info("检查前自动同步完成：%s 年数据已更新（%s）",
                            "、".join(str(y) for y in ok_years), store.synced_at)
                # 让状态标签反映最新同步时间
                self._ui_queue.put({"type": "refresh_holiday_status"})
            if failed:
                logger.warning(
                    "检查前自动同步有 %d 个年份未获取到数据（多为官方尚未公布或"
                    "网络不可达），本次检查沿用本地缓存与内置数据", len(failed))
        except Exception:  # noqa: BLE001 —— 自动同步永不影响检查主流程
            logger.exception("检查前自动同步出现未预期异常（已忽略，继续检查）")

    def _on_sync_holidays(self) -> None:
        """同步最新节假日安排，成功后展示前后一个月的休假计划（工作线程）。"""
        self._btn_sync.config(state=tk.DISABLED)

        def work() -> None:
            try:
                today = date.today()
                store = HolidayStore(DATA_DIR / "holidays.json")
                ok_years, failed = sync_window_years(store, today)
                days, start, end = self._build_holiday_days(store)
                suffix = self._window_suffix(start, end)

                if not days:
                    # 同步失败且本地/内置都无数据：仅提示，不弹空窗口
                    self._ui_queue.put({
                        "type": "error", "title": "节假日同步失败",
                        "message": "\n\n".join(failed) or "未获取到任何节假日数据",
                    })
                    return

                if failed:
                    note = ("⚠ 部分年份同步失败：" + "；".join(
                        f.split('：')[0] for f in failed)
                        + "。以下为本地缓存 + 内置数据。")
                else:
                    note = (f"✔ 同步成功（{'、'.join(str(y) for y in ok_years)} 年），"
                            f"数据更新于 {store.synced_at}。")
                note += f" 共 {len(days)} 项，来源：本地同步数据 + chinese-calendar 内置数据。"

                logger.info("节假日同步完成：成功年份 %s，失败 %d 个，计划项 %d 条",
                            ok_years, len(failed), len(days))
                self._ui_queue.put({
                    "type": "holiday_window", "days": days,
                    "suffix": suffix, "note": note,
                })
                if failed:
                    self._ui_queue.put({
                        "type": "error", "title": "部分年份同步失败",
                        "message": "\n".join(failed),
                    })
            except Exception:  # noqa: BLE001 —— 兜底
                logger.exception("节假日同步出现未预期异常")
                self._ui_queue.put({
                    "type": "error", "title": "错误",
                    "message": "节假日同步出现未预期异常，详情请查看日志窗口",
                })
            finally:
                self._ui_queue.put({"type": "refresh_holiday_status"})
                self._ui_queue.put({"type": "enable_sync"})
        threading.Thread(target=work, daemon=True, name="holiday-sync").start()

    # ================================================================ 检查流水线

    def _request_check(self) -> None:
        """线程安全的检查请求入口：任何线程可调用。

        调度线程 / 启动补检通过本方法把请求投递到 UI 队列，由主线程
        的 _poll_queues 转交 _run_check_threaded 执行——绝不跨线程碰控件。
        """
        self._ui_queue.put({"type": "run_check"})

    def _run_check_threaded(self, export_only: bool = False) -> None:
        """在后台线程执行完整检查流水线（手动 / 定时 / 补检共用）。

        Args:
            export_only: True 时只把提醒导出成本地邮件，不发送、不标记已发
                         （「导出邮件到本地」按钮使用）。

        仅允许主线程调用（经 _poll_queues 或按钮事件）。
        """
        if self._check_busy:
            logger.info("上一次检查仍在进行中，跳过本次触发")
            return
        self._check_busy = True
        # 按钮置灰需在主线程完成
        self._btn_check.config(state=tk.DISABLED)
        cfg = self._gather_config()  # 主线程快照

        def work() -> None:
            try:
                self._do_check(cfg, export_only=export_only)
            except Exception:  # noqa: BLE001 —— 兜底，任何异常都不静默
                logger.exception("检查过程出现未预期异常")
                self._ui_queue.put({
                    "type": "error", "title": "错误",
                    "message": "检查过程出现未预期异常，详情请查看日志窗口和日志文件",
                })
            finally:
                self._ui_queue.put({"type": "enable_check"})
        threading.Thread(target=work, daemon=True, name="check").start()

    def _do_check(self, cfg: AppConfig, export_only: bool = False) -> None:
        """完整检查流水线（工作线程中执行）。

        步骤：校验配置 -> 读 Excel -> 同步最新节假日 -> 筛选触发项
              -> 去重 -> 投递（SMTP 直发 / 导出到本机客户端）-> 记录。
        """
        # 1. 前置校验（按发送方式区分要求：兜底通道不需要授权码/密码）
        if not cfg.excel_path:
            self._ui_queue.put({"type": "error", "title": "缺少文件",
                                "message": "请先选择 Excel 数据文件"})
            return
        need_smtp = (not export_only) and cfg.send_mode != "client"
        if need_smtp:
            if not cfg.smtp.username or not cfg.smtp.password or not cfg.smtp.recipients:
                self._ui_queue.put({
                    "type": "error", "title": "缺少邮箱配置",
                    "message": "请先完整填写并保存邮箱配置（发件人、授权码、目标邮箱）"})
                return
        elif not cfg.smtp.recipients:
            self._ui_queue.put({
                "type": "error", "title": "缺少目标邮箱",
                "message": "请先填写并保存「目标邮箱」（本机邮件客户端模式无需授权码/密码）"})
            return
        invalid = find_invalid_recipients(cfg.smtp.recipients)
        if invalid:
            self._ui_queue.put({
                "type": "error", "title": "目标邮箱格式不正确",
                "message": "以下目标邮箱格式不正确，无法发送：\n  · " + "\n  · ".join(invalid)
                           + "\n\n常见原因：漏写 @ 或域名后缀（如 .com）。\n"
                             "请到「邮箱配置」修正并保存后重试。",
            })
            return
        selected = {p for p, on in cfg.checked_perms.items() if on}
        if not selected:
            self._ui_queue.put({"type": "error", "title": "未勾选权限",
                                "message": "请至少勾选一个要检查的权限到期日期列"})
            return

        # 2. 读取 Excel
        try:
            records: List[ClientRecord] = read_clients(cfg.excel_path)
        except ExcelFormatError as exc:
            logger.error("Excel 读取失败：%s", exc)
            self._ui_queue.put({"type": "error", "title": "Excel 读取失败", "message": str(exc)})
            return

        # 3. 检查前先同步最新节假日（默认开启；失败自动降级，不打断检查）
        if cfg.auto_sync_holidays:
            self._auto_sync_holidays()

        # 4. 筛选 + 去重（手动"只导出"时不去重，便于随时重新生成邮件核对）
        today = date.today()
        items = find_due_items(records, selected, today, create_calendar())
        if not export_only:
            items = self._sent_store.filter_unsent(items)
        if not items:
            logger.info("检查完成：今日没有需要提醒的客户权限（%s）",
                        datetime.now().strftime("%Y-%m-%d %H:%M"))
            self._ui_queue.put({"type": "info", "title": "检查完成",
                                "message": "今日没有需要提醒的客户权限。"})
            return

        # 5. 渲染汇总邮件
        subject = "【权限到期提醒】"
        html, plain = render_summary_email(items, today)
        attachments = [build_excel_attachment(items)]

        # 5a. 只导出：落盘即结束，不发送、不标记已发
        if export_only:
            try:
                eml, html_path = export_mail(
                    cfg.smtp.recipients, subject, html, plain,
                    from_addr=cfg.smtp.username, tag="权限到期提醒",
                    attachments=attachments)
            except OSError as exc:
                logger.error("导出提醒邮件失败：%s", exc)
                self._ui_queue.put({"type": "error", "title": "导出失败",
                                    "message": f"无法写入导出目录：{exc}"})
                return
            logger.info("已按需导出提醒邮件：%d 项 -> %s", len(items), eml)
            self._ui_queue.put({
                "type": "info", "title": "已导出到本地",
                "message": f"已导出 {len(items)} 项到期提醒（未发送、未标记已发）：\n"
                           f"  · 邮件文件：{eml}\n  · 正文预览：{html_path}\n\n"
                           "可在 Foxmail 中双击打开 .eml 后点「发送」投递。",
            })
            return

        # 5b. 兜底通道：导出 + 唤起本机邮件客户端（由用户确认后再标记已发）
        if cfg.send_mode == "client":
            self._export_and_handoff(cfg, items, subject, html, plain,
                                      attachments=attachments)
            return

        # 5c. SMTP 直发
        try:
            EmailSender(cfg.smtp).send(cfg.smtp.recipients, subject, html, plain,
                                       attachments=attachments)
        except MailSendError as exc:
            logger.error("提醒邮件发送失败（%d 项未标记已发，下次将重试）：%s", len(items), exc)
            self._ui_queue.put({"type": "error", "title": "邮件发送失败", "message": str(exc)})
            return

        # 6. 发送成功 -> 标记已发（逐项落盘）
        for it in items:
            self._sent_store.mark_sent(it.customer_id, it.perm, it.trigger_date, today)
        logger.info("提醒邮件已发送：%d 项，收件人 %s",
                    len(items), ", ".join(cfg.smtp.recipients))
        summary = "、".join(
            f"{it.customer_id}({it.perm_label} {_fmt_cn(it.expiry)})" for it in items
        )
        self._ui_queue.put({
            "type": "info", "title": "提醒已发送",
            "message": f"已发送 {len(items)} 项到期提醒：\n{summary}",
        })

    # ================================================================ 队列轮询

    def _poll_queues(self) -> None:
        """主线程每 200ms 轮询：刷新日志窗口、处理 UI 动作。"""
        if getattr(self, "_closing", False):
            return  # 窗口正在关闭，停止轮询，避免操作已销毁的控件
        # 日志
        try:
            while True:
                line = self._log_queue.get_nowait()
                self._txt_log.config(state=tk.NORMAL)
                self._txt_log.insert(tk.END, line + "\n")
                self._txt_log.see(tk.END)
                self._txt_log.config(state=tk.DISABLED)
        except queue.Empty:
            pass
        except tk.TclError:
            return  # 控件已销毁（竞态兜底）
        # UI 动作（弹窗 / 按钮恢复 / 窗口打开）
        try:
            while True:
                action = self._ui_queue.get_nowait()
                kind = action.get("type")
                if kind in ("info", "error"):
                    (messagebox.showinfo if kind == "info" else messagebox.showerror)(
                        action.get("title", ""), action.get("message", ""))
                elif kind == "run_check":
                    # 调度线程 / 启动补检转回主线程触发检查
                    self._run_check_threaded()
                elif kind == "holiday_window":
                    show_holiday_plan(self, action["days"], action["suffix"],
                                      action["note"])
                elif kind == "diag_report":
                    show_diag_report(self, action["report"],
                                     action.get("recommended"),
                                     on_apply=self._apply_recommended)
                elif kind == "refresh_holiday_status":
                    self._refresh_holiday_status()
                elif kind == "confirm_export":
                    self._confirm_export_sent(action)
                elif kind == "enable_check":
                    self._btn_check.config(state=tk.NORMAL)
                    self._check_busy = False
                elif kind == "enable_test":
                    self._btn_test.config(state=tk.NORMAL)
                elif kind == "enable_sync":
                    self._btn_sync.config(state=tk.NORMAL)
                elif kind == "enable_diag":
                    self._btn_diag.config(state=tk.NORMAL, text="连接诊断")
        except queue.Empty:
            pass
        except tk.TclError:
            return  # 控件已销毁（竞态兜底）
        self.after(200, self._poll_queues)

    def _confirm_export_sent(self, action: dict) -> None:
        """兜底通道：由用户确认邮件是否真的发出去了，再决定是否标记已发。

        刻意不做"导出即视为已发"——提醒事项一旦被标记，当天就不会再出现，
        而兜底通道的最后一步（在 Foxmail 里点发送）是由用户完成的，
        工具无从得知结果，因此必须问一句。
        """
        pairs = action.get("items") or []
        today = action.get("today") or date.today()
        logger.info("提醒邮件已导出并在邮件客户端中打开，等待用户确认是否已发送")
        if messagebox.askyesno(
            "确认是否已发送",
            action.get("message", "") +
            "\n\n是否已在邮件客户端中成功发送？\n"
            "  · 是 → 标记为已发送，下次不再重复提醒；\n"
            "  · 否 → 不标记，下次检查会再次提醒。",
        ):
            for customer_id, perm, trigger_date in pairs:
                self._sent_store.mark_sent(customer_id, perm, trigger_date, today)
            logger.info("用户确认已发送，标记 %d 项为已发", len(pairs))
            messagebox.showinfo("已标记", f"已标记 {len(pairs)} 项为已发送。")
        else:
            logger.info("用户未确认发送，%d 项保留待下次提醒", len(pairs))

    def _on_close(self) -> None:
        """关闭窗口：停调度线程后退出。"""
        self._closing = True
        self._scheduler.stop()
        logger.info("程序退出")
        self.destroy()


def _fmt_cn(d) -> str:
    """日期格式化为'某年某月某日'。"""
    return f"{d.year}年{d.month}月{d.day}日"
