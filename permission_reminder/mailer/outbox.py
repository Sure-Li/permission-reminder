# -*- coding: utf-8 -*-
"""本地兜底投递通道（完全不依赖 SMTP）。

适用场景
--------
单位邮件系统启用了安全网关（如 Coremail + CACTER），第三方客户端认证被策略拒绝
（典型返回 `535 reject by CACTER gateway`），在拿到正确的客户端专用密码或管理员
放行之前，SMTP 直发这条路走不通。

此时本模块把提醒邮件**落成本地文件**，再借助用户本机已经配置好、能正常发信的
邮件客户端（Foxmail / Outlook / Windows 邮件等）发出：

- `xxx.eml`  —— 标准 MIME 邮件（UTF-8，纯文本 + HTML 双格式）。
  在 Foxmail 中双击即可打开，点「再次编辑 / 转发」后直接发送即完成投递。
  用「发送邮件»」之外，也可直接拖拽到客户端的写信窗口。
- `xxx.html` —— 正文预览，用浏览器直接打开核对内容。

另外提供 `open_in_mail_client()`：通过 `mailto:` 唤起系统默认邮件客户端的
**新建邮件窗口**（收件人 / 主题 / 正文摘要已预填），适合想一步到位的场景。

设计约定
--------
- 所有文件写入 `<应用根目录>/data/outbox/`，文件名带时间戳，**不覆盖历史文件**；
- 文件名中的时间戳用 `YYYYmmdd-HHMMSS`，便于人工按时间查找；
- 本模块不吞异常：调用方（UI 层）负责捕获并提示。
"""
from __future__ import annotations

import logging
import os
import subprocess
import sys
import urllib.parse
from datetime import datetime
from html import escape
from pathlib import Path
from typing import List, Optional, Tuple

from ..infra.config_store import DATA_DIR
from .sender import Attachment, build_mime_message

logger = logging.getLogger(__name__)

# 兜底通道的导出子目录
OUTBOX_DIRNAME = "outbox"

# mailto: 正文在不同客户端都有长度上限（Outlook/Foxmail 约为 2000 字符），
# 超长会被截断甚至让客户端拒绝打开，因此正文只放摘要，完整内容见 .eml / .html
MAILTO_BODY_LIMIT = 1500


def outbox_dir() -> Path:
    """返回导出的目录（不存在则创建）。"""
    d = DATA_DIR / OUTBOX_DIRNAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def _stamp() -> str:
    """文件名用时间戳（精确到秒）。"""
    return datetime.now().strftime("%Y%m%d-%H%M%S")


def _safe_name(prefix: str) -> str:
    """生成合法的文件名主体（去掉 Windows 不允许的字符）。"""
    bad = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in bad else ch for ch in prefix).strip()
    return cleaned or "邮件"


def _unique_path(folder: Path, stem: str, suffix: str) -> Path:
    """返回一个尚未被占用的文件路径。

    同一秒内连续导出两次时时间戳会相同，若直接写入会把上一封邮件**覆盖**掉。
    这里退化为追加 `-1` / `-2` 序号，保证每次导出都留有独立文件。
    """
    candidate = folder / f"{stem}{suffix}"
    if not candidate.exists():
        return candidate
    for i in range(1, 1000):
        candidate = folder / f"{stem}-{i}{suffix}"
        if not candidate.exists():
            return candidate
    # 极端情况（同秒导出上千次）：退回带毫秒的时间戳，仍不覆盖
    return folder / f"{stem}-{datetime.now().strftime('%f')}{suffix}"


def _html_document(subject: str, html_body: str) -> str:
    """给正文套一层最小外壳，保证浏览器以 UTF-8 正确渲染。"""
    return ("<!DOCTYPE html><html lang=\"zh-CN\"><head><meta charset=\"utf-8\">"
            f"<title>{escape(subject)}</title></head><body>{html_body}</body></html>")


def export_eml(recipients: List[str], subject: str, html_body: str,
               plain_body: str, from_addr: str = "",
               tag: str = "", attachments: Optional[List[Attachment]] = None) -> Path:
    """把邮件写成本地 .eml 文件，返回文件路径。

    Args:
        recipients: 收件人列表（写入 To 头）。
        subject:    邮件主题。
        html_body:  HTML 正文。
        plain_body: 纯文本正文（作为 alternative 的第一段）。
        from_addr:  发件人地址，仅用于写入 From 头（可留空）。
        tag:        文件名前缀，便于人工区分（如"权限到期提醒"）。
    """
    path = _unique_path(outbox_dir(), f"{_safe_name(tag or subject)}_{_stamp()}", ".eml")
    msg = build_mime_message(from_addr, recipients, subject, html_body, plain_body,
                             attachments=attachments)
    path.write_bytes(msg.as_bytes())
    logger.info("已导出邮件到本地: %s", path)
    return path


def export_html(subject: str, html_body: str, tag: str = "") -> Path:
    """把 HTML 正文单独导出，便于浏览器直接核对排版。"""
    path = _unique_path(outbox_dir(), f"{_safe_name(tag or subject)}_{_stamp()}", ".html")
    path.write_text(_html_document(subject, html_body), encoding="utf-8")
    logger.info("已导出 HTML 预览: %s", path)
    return path


def export_mail(recipients: List[str], subject: str, html_body: str,
                plain_body: str, from_addr: str = "",
                tag: str = "", attachments: Optional[List[Attachment]] = None) -> Tuple[Path, Path]:
    """一次性导出 .eml 与 .html，返回 (eml 路径, html 路径)。"""
    folder = outbox_dir()
    stem = f"{_safe_name(tag or subject)}_{_stamp()}"

    eml_path = _unique_path(folder, stem, ".eml")
    msg = build_mime_message(from_addr, recipients, subject, html_body, plain_body,
                             attachments=attachments)
    eml_path.write_bytes(msg.as_bytes())

    # html 与 eml 使用同一 stem；eml 已占位时给 html 也带上同样后缀，便于成对识别
    html_stem = eml_path.stem
    html_path = _unique_path(folder, html_stem, ".html")
    html_path.write_text(_html_document(subject, html_body), encoding="utf-8")

    logger.info("已导出邮件到本地: %s / %s", eml_path, html_path)
    return eml_path, html_path


def _query_assoc(protocol: str, assoc_str: int) -> str:
    """用 AssocQueryStringW 查询协议默认处理程序的某项信息。

    该 API 由 **shlwapi.dll** 导出（shell32 只按序号导出，找不到名字）。
    assoc_str: 1=COMMAND, 2=EXECUTABLE, 3=FRIENDLYDOCNAME, 4=FRIENDLYAPPNAME
    """
    if not sys.platform.startswith("win"):
        return ""
    try:
        import ctypes
        from ctypes import wintypes

        shlwapi = ctypes.windll.shlwapi  # type: ignore[attr-defined]
        shlwapi.AssocQueryStringW.argtypes = [
            wintypes.DWORD, wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR,
            wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD),
        ]
        shlwapi.AssocQueryStringW.restype = ctypes.c_long
        size = wintypes.DWORD(1024)
        buf = ctypes.create_unicode_buffer(1024)
        hr = shlwapi.AssocQueryStringW(0, assoc_str, protocol, None,
                                       buf, ctypes.byref(size))
        if hr == 0:
            return (buf.value or "").strip()
    except (OSError, AttributeError):
        pass
    return ""


# 没有默认处理程序时，Windows 会用 OpenWith.exe 弹「选取应用」
_NO_HANDLER_MARKERS = ("openwith", "选取应用", "pick an app", "how do you want to open")


def has_mail_client() -> bool:
    """系统是否配置了真正可用的 mailto 处理程序。

    若返回 False，直接调 `mailto:` 只会弹出 Windows 的「选取应用」对话框——
    对用户毫无帮助，因此调用方应改为"打开导出目录 + 提示手动发送"。
    """
    exe = _query_assoc("mailto", 2).lower()
    if not exe:
        return False
    return not any(m in exe for m in _NO_HANDLER_MARKERS)


def describe_mail_client() -> str:
    """描述当前默认邮件客户端，供界面提示使用（探测失败返回空串）。"""
    if not sys.platform.startswith("win"):
        return ""
    name = _query_assoc("mailto", 4)
    if has_mail_client():
        return name or "已配置的邮件客户端"
    return "未设置默认邮件客户端" + (f"（点击会弹出「{name}」选择框）" if name else "")


def detect_mail_client() -> str:
    """探测系统当前默认的邮件客户端，返回可读名称。

    兜底通道靠 `mailto:` 唤起系统默认邮件客户端，而不同机器上它可能是
    Foxmail、Outlook 或 Windows 自带邮件应用。提前告知用户"会弹出哪个程序"，
    避免出现"以为在发 Foxmail、实际弹了别的东西"的意外。

    实现上调用 Windows 的 `AssocQueryStringW(ASSOCSTR_FRIENDLYAPPNAME)`——
    这正是资源管理器"默认应用"里显示的那个名字，比翻注册表可靠。
    """
    if not sys.platform.startswith("win"):
        return ""
    if has_mail_client():
        return _query_assoc("mailto", 4) or "已配置的邮件客户端"

    # 退化：看有没有注册过经典邮件客户端（可能只是没被设为 mailto 默认程序）
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Clients\Mail") as key:
            names, i = [], 0
            while True:
                try:
                    names.append(winreg.EnumKey(key, i))
                except OSError:
                    break
                i += 1
            # Hotmail 是 Windows 自带的历史项，不代表"当前客户端"
            real = [n for n in names if n.lower() != "hotmail"]
            if real:
                return "、".join(real) + "（已安装，但未设为系统默认邮件程序）"
    except (ImportError, OSError):
        pass
    return "未找到已注册的邮件客户端"


def eml_viewer() -> str:
    """返回 .eml 文件默认关联的程序名（用于提示"双击 .eml 会用谁打开"）。"""
    return _query_assoc(".eml", 4)


def build_mailto(recipients: List[str], subject: str, plain_body: str) -> str:
    """构造 mailto: 链接（供唤起本机邮件客户端的新建邮件窗口）。"""
    body = plain_body or ""
    if len(body) > MAILTO_BODY_LIMIT:
        body = body[:MAILTO_BODY_LIMIT] + "\n……（内容过长已截断，完整内容请见导出的 .eml / .html 文件）"
    to = ",".join(recipients)
    # safe 保留换行与常见可读字符，其余交给标准库编码
    query = urllib.parse.urlencode({"subject": subject, "body": body},
                                   quote_via=urllib.parse.quote, safe="")
    return f"mailto:{to}?{query}"


def open_in_mail_client(recipients: List[str], subject: str,
                        plain_body: str, eml_path: Optional[Path] = None) -> bool:
    """唤起系统默认邮件客户端的新建邮件窗口（收件人/主题/正文摘要已预填）。

    Returns:
        True  —— 已把 `mailto:` 交给系统打开；
        False —— 系统没有配置默认邮件客户端（Windows 会弹「选取应用」选择框，
                 对用户没有任何帮助），因此**不打开**，由调用方引导手动发送。

    说明：
      - Windows 下用 `os.startfile`，macOS 用 `open`，其他平台用 `xdg-open`；
      - 只是"打开写信窗口"，最终仍需用户点发送——这正是让已配置好的
        Foxmail/Outlook 用它自己的凭据投递，从而绕开 SMTP 认证被拒的问题。
    """
    if eml_path is not None:
        # mailto 没有可移植的附件参数；附件提醒必须打开完整 MIME 文件。
        path = Path(eml_path).resolve(strict=True)
        if sys.platform.startswith("win"):
            handler = _query_assoc(".eml", 2).lower()
            if not handler or any(marker in handler for marker in _NO_HANDLER_MARKERS):
                logger.warning("系统没有可用的 .eml 关联，请手动发送导出的完整邮件")
                return False
            os.startfile(str(path))  # type: ignore[attr-defined]  # noqa: S606
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)  # noqa: S603,S607
        else:
            subprocess.run(["xdg-open", str(path)], check=False)  # noqa: S603,S607
        return True
    if sys.platform.startswith("win") and not has_mail_client():
        logger.warning("系统未设置默认邮件客户端（mailto 无关联），跳过自动唤起，"
                       "改为由用户手动发送导出的 .eml")
        return False
    url = build_mailto(recipients, subject, plain_body)
    if sys.platform.startswith("win"):
        os.startfile(url)  # type: ignore[attr-defined]  # noqa: S606 —— 仅打开邮件协议
    elif sys.platform == "darwin":
        subprocess.run(["open", url], check=False)  # noqa: S603,S607
    else:
        subprocess.run(["xdg-open", url], check=False)  # noqa: S603,S607
    return True


def open_folder(path: Path) -> None:
    """在系统文件管理器中打开目录（方便用户拖走导出的文件）。"""
    target = path if path.is_dir() else path.parent
    if sys.platform.startswith("win"):
        os.startfile(str(target))  # type: ignore[attr-defined]  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.run(["open", str(target)], check=False)  # noqa: S603,S607
    else:
        subprocess.run(["xdg-open", str(target)], check=False)  # noqa: S603,S607
