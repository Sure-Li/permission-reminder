# -*- coding: utf-8 -*-
"""配置持久化模块。

职责：
1. 定义应用配置数据结构（AppConfig 及子结构）；
2. 提供 JSON 配置文件的加载 / 保存 / 清除；
3. 便携式路径策略：源码运行取脚本根目录，PyInstaller 打包后取 exe 所在目录，
   config.json / logs / data 均放在该目录下。
"""
from __future__ import annotations

import json
import logging
import sys
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import List

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------- 路径策略 ---

def get_app_root() -> Path:
    """获取应用根目录（便携式：源码取项目根，打包后取 exe 所在目录）。"""
    if getattr(sys, "frozen", False):  # PyInstaller 打包后的运行环境
        return Path(sys.executable).parent
    # permission_reminder/infra/config_store.py -> 项目根是上两级
    return Path(__file__).resolve().parent.parent.parent


APP_ROOT: Path = get_app_root()
CONFIG_PATH: Path = APP_ROOT / "config.json"
DATA_DIR: Path = APP_ROOT / "data"
LOG_DIR: Path = APP_ROOT / "logs"


# ------------------------------------------------------------ 配置数据结构 ---

@dataclass
class SmtpConfig:
    """发件邮箱与 SMTP 服务器配置。"""
    server: str = "smtp.qq.com"
    port: int = 465
    ssl_mode: str = "SSL"          # "SSL" 或 "STARTTLS"
    username: str = ""             # 发件人邮箱
    password: str = ""             # 授权码 / 密码
    recipients: List[str] = field(default_factory=list)  # 目标邮箱列表


@dataclass
class ScheduleConfig:
    """定时任务配置。"""
    enabled: bool = False          # 定时开关（持久化，重启后自动恢复）
    hour: int = 9                  # 触发时
    minute: int = 0                # 触发分


@dataclass
class AppConfig:
    """应用总配置。"""

    # 发送方式：
    #   "smtp"   —— 由本工具直连 SMTP 服务器发送（默认）
    #   "client" —— 兜底通道：把邮件导出到 data/outbox/ 并唤起本机邮件客户端
    #               （Foxmail/Outlook）发送，用于 SMTP 认证被单位网关拒绝的场景
    SEND_MODES = ("smtp", "client")

    smtp: SmtpConfig = field(default_factory=SmtpConfig)
    excel_path: str = ""           # 上次选择的 Excel 文件路径
    schedule: ScheduleConfig = field(default_factory=ScheduleConfig)
    auto_sync_holidays: bool = True  # 每次检查前自动同步最新节假日（默认开启）
    send_mode: str = "smtp"        # 发送方式（见上）
    # Foxmail 集成：兜底通道优先用它投递（不需要 SMTP 授权码/密码）
    #   foxmail_path —— Foxmail.exe 路径（空 = 未指定，运行时自动探测）
    #   foxmail_mode —— 唤起方式："auto" / "eml" / "mailto"（见 mailer/foxmail.py）
    FOXMAIL_MODES = ("auto", "eml", "mailto")
    foxmail_path: str = ""
    foxmail_mode: str = "auto"
    checked_perms: dict = field(
        default_factory=lambda: {
            "专业投资者": True,    # 专业投资者到期日期
            "私募基金": True,      # 私募基金到期日期
            "资管权限": True,      # 资管权限到期日期
        }
    )


# ----------------------------------------------------------------- 读写逻辑 ---

def load_config() -> AppConfig:
    """加载配置文件；不存在或损坏时返回默认配置（损坏则备份原文件并记录日志）。"""
    if not CONFIG_PATH.exists():
        logger.info("配置文件不存在，使用默认配置: %s", CONFIG_PATH)
        return AppConfig()
    try:
        raw = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        cfg = AppConfig(
            smtp=SmtpConfig(**{k: v for k, v in raw.get("smtp", {}).items()
                               if k in SmtpConfig.__dataclass_fields__}),
            excel_path=str(raw.get("excel_path", "")),
            schedule=ScheduleConfig(**{k: v for k, v in raw.get("schedule", {}).items()
                                       if k in ScheduleConfig.__dataclass_fields__}),
            checked_perms={**AppConfig().checked_perms,
                           **{k: bool(v) for k, v in raw.get("checked_perms", {}).items()}},
            # 旧版本配置文件无此字段时默认开启（保持"默认先更新"的行为）
            auto_sync_holidays=bool(raw.get("auto_sync_holidays", True)),
            # 旧版本配置文件无此字段时默认 SMTP 直发（保持原有行为）
            send_mode=str(raw.get("send_mode", "smtp")),
            # Foxmail 集成（旧配置无此字段时留空 + 自动方式，运行时自动探测）
            foxmail_path=str(raw.get("foxmail_path", "")),
            foxmail_mode=str(raw.get("foxmail_mode", "auto")),
        )
        # 端口等数值字段防御
        cfg.smtp.port = int(cfg.smtp.port)
        cfg.schedule.hour = max(0, min(23, int(cfg.schedule.hour)))
        cfg.schedule.minute = max(0, min(59, int(cfg.schedule.minute)))
        if not isinstance(cfg.smtp.recipients, list):
            cfg.smtp.recipients = []
        if cfg.send_mode not in AppConfig.SEND_MODES:
            # 配置文件被手工改坏时回落到默认发送方式，避免整份配置不可用
            logger.warning("未知的 send_mode=%r，已回落为 smtp", cfg.send_mode)
            cfg.send_mode = "smtp"
        if cfg.foxmail_mode not in AppConfig.FOXMAIL_MODES:
            logger.warning("未知的 foxmail_mode=%r，已回落为 auto", cfg.foxmail_mode)
            cfg.foxmail_mode = "auto"
        cfg.foxmail_path = cfg.foxmail_path.strip()
        logger.info("已加载配置文件: %s", CONFIG_PATH)
        return cfg
    except (json.JSONDecodeError, TypeError, ValueError) as exc:
        backup = CONFIG_PATH.with_suffix(".json.bak")
        try:
            CONFIG_PATH.replace(backup)
            logger.warning("配置文件损坏（%s），已备份至 %s，使用默认配置", exc, backup)
        except OSError:
            logger.warning("配置文件损坏（%s）且备份失败，使用默认配置", exc)
        return AppConfig()


def save_config(cfg: AppConfig) -> None:
    """保存配置到 JSON 文件（UTF-8，缩进 2）。"""
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.write_text(
        json.dumps(asdict(cfg), ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    logger.info("配置已保存: %s", CONFIG_PATH)


def clear_config() -> None:
    """清除本地配置文件（仅清配置，日志与发送记录保留）。"""
    if CONFIG_PATH.exists():
        CONFIG_PATH.unlink()
        logger.info("配置文件已清除: %s", CONFIG_PATH)
