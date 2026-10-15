# -*- coding: utf-8 -*-
"""邮件发送与正文渲染模块。

职责：
1. EmailSender：SSL(465) / STARTTLS(587) 两种加密方式，超时 + 3 次指数退避重试；
2. render_summary_email：把当日待提醒项渲染成 HTML + 纯文本双格式汇总邮件；
3. send_test_email：验证 SMTP 配置的测试邮件。
"""
from __future__ import annotations

import logging
import re
import smtplib
import socket
import ssl
import time
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import List, Optional, Tuple

from ..core.checker import ReminderItem
from ..infra.config_store import SmtpConfig

logger = logging.getLogger(__name__)

# 主题固定为【权限到期提醒】
SUBJECT = "【权限到期提醒】"

# SMTP 超时与重试
SMTP_TIMEOUT_SECONDS = 15
MAX_RETRIES = 3
RETRY_DELAYS = (5, 10, 20)  # 指数退避（秒）

# 支持的加密方式：SSL（465 隐式 TLS）/ STARTTLS（587/25 显式升级）/ PLAIN（不加密，仅 25）
SSL_MODES = ("SSL", "STARTTLS", "PLAIN")

# ---- 反封禁保护 ----------------------------------------------------------
# 企业邮件安全网关（如 Coremail CACTER）会按来源 IP 统计 AUTH 失败次数，
# 超过阈值即封禁该 IP（返回 "554 ... Smtp auth error limit exceeded"）。
# 认证失败属"配置/策略"问题而非瞬时故障，反复重试只会加重封禁，
# 因此在认证失败后进入冷却期，冷却期内直接报错、不再向服务器发起认证。
AUTH_FAILURE_COOLDOWN_SECONDS = 600      # 10 分钟
_auth_block_until: float = 0.0           # 冷却截止时间戳（0 表示无冷却）


def note_auth_failure() -> None:
    """记录一次认证失败/网关封禁，开启冷却期。"""
    global _auth_block_until
    _auth_block_until = time.time() + AUTH_FAILURE_COOLDOWN_SECONDS


def reset_auth_block() -> None:
    """清除冷却期（用户修改并保存邮箱配置、或主动点测试邮件时调用）。"""
    global _auth_block_until
    _auth_block_until = 0.0


def auth_block_remaining() -> int:
    """剩余冷却秒数（0 表示未处于冷却期）。"""
    return max(0, int(_auth_block_until - time.time()))


# 常用 SMTP 预设（UI 快速填充用）
SMTP_PRESETS = {
    "QQ 邮箱": ("smtp.qq.com", 465, "SSL"),
    "网易 163": ("smtp.163.com", 465, "SSL"),
    "腾讯企业邮箱": ("smtp.exmail.qq.com", 465, "SSL"),
    "自建邮局（25 + STARTTLS）": ("mail.你的域名.com", 25, "STARTTLS"),
}


class MailSendError(Exception):
    """邮件发送失败（重试耗尽后抛出）。"""


def build_mime_message(from_addr: str, recipients: List[str], subject: str,
                       html_body: str, plain_body: str) -> MIMEMultipart:
    """构建 MIME 邮件（UTF-8，纯文本 + HTML 双格式）。

    SMTP 直发与「导出为 .eml」两条通道共用同一实现，保证两者内容完全一致。
    """
    msg = MIMEMultipart("alternative")
    msg["Subject"] = Header(subject, "utf-8")
    msg["From"] = from_addr
    msg["To"] = ", ".join(recipients)
    msg.attach(MIMEText(plain_body, "plain", "utf-8"))
    msg.attach(MIMEText(html_body, "html", "utf-8"))
    return msg


def is_gateway_rejection(text: str) -> bool:
    """判断服务器返回文本是否为安全网关的拒绝/封禁（而非普通业务错误）。

    典型：'554 reject by CACTER gateway, IP<x> is Hangup, Smtp auth error limit exceeded'
    """
    t = (text or "").lower()
    keys = ("reject by", "is hangup", "auth error limit", "blacklist",
            "blocked", "denied", "too many")
    return any(k in t for k in keys)


# ---------------------------------------------------------- 收件人格式校验 --

# 宽松但能拦住高频笔误（如漏写 .com 的点、缺少 @）的邮箱格式：
# 本地部分不含 @ 与分隔符；域名部分要求「标签.顶级域(≥2 个字母)」，故 a@b.c 与 a@b.com. 均判为非法
_EMAIL_RE = re.compile(r"^[^@\s,;<>()\[\]\\]+@"
                       r"[A-Za-z0-9]([A-Za-z0-9\-]*[A-Za-z0-9])?"
                       r"(\.[A-Za-z0-9]([A-Za-z0-9\-]*[A-Za-z0-9])?)*"
                       r"\.[A-Za-z]{2,}$")


def is_valid_email(addr: str) -> bool:
    """判断邮箱地址格式是否合法（含域名后缀，如 .com）。

    例：'recipient@examplecom'（漏点）→ False；'a@b.com' → True。
    """
    return bool(_EMAIL_RE.match((addr or "").strip()))


def find_invalid_recipients(recipients: List[str]) -> List[str]:
    """返回列表中格式不合法的收件人（保持原顺序）。"""
    return [r for r in (recipients or []) if not is_valid_email(r)]


# ------------------------------------------------------------ 错误处置提示 --

def _probe_local_mail_client() -> Tuple[Optional[bool], str]:
    """探测本机 `mailto:` 默认处理程序，供自检清单第 ④ 步"按本机实情说话"。

    Returns:
        (是否可用, 客户端描述)。第一个元素为 None 表示探测不可用
        （非 Windows 或探测失败），此时调用方应回退到通用说法。

    注：outbox 模块已导入 sender，因此这里只能**延迟导入**，不能放到模块顶层。
    """
    try:
        from .outbox import describe_mail_client, has_mail_client  # 延迟导入避免循环
    except Exception:  # noqa: BLE001 —— 探测失败不应影响报错本身
        return None, ""
    try:
        return bool(has_mail_client()), describe_mail_client()
    except Exception:  # noqa: BLE001
        return None, ""


def _selfhost_checklist(local_client: Optional[Tuple[Optional[bool], str]] = None) -> str:
    """自建邮局（Coremail/盈世、Exchange 等）认证被拒的四步自检清单。

    顺序刻意按"成本从低到高"排列：先确认密码，再换密码类型，再看地址，
    最后才找管理员——避免用户一上来就去联系 IT。

    第 ④ 步依赖"手工用邮件客户端发一封测试邮件"来判别病因，而这一步**并非
    所有机器都能做**（例如本机根本没装 Foxmail、`mailto:` 也没有默认程序）。
    因此这里先探测本机环境，再给出可执行的建议，而不是一律写"用 Foxmail 试试"。
    """
    if local_client is None:
        local_client = _probe_local_mail_client()
    has_client, client_name = local_client

    head = ("  ④ 客户端 / 来源 IP 是否被限制：部分单位启用了邮件安全网关"
            "（如 CACTER），只放行白名单客户端（Foxmail / Outlook）"
            "或仅允许内网 IP 认证。")
    if has_client:
        step4 = (head + f"本机默认邮件客户端是「{client_name}」，"
                        "可用它手工配同一个账号发一封测试邮件来判定：\n"
                        "      能发出 → 属于客户端白名单策略，请联系邮箱管理员"
                        "把本机出口 IP（或本工具）加入允许范围；\n"
                        "      同样发不出 → 回到 ① ②，问题在凭据本身。\n")
    elif has_client is False:
        step4 = (head + "但本机当前【没有默认邮件客户端】"
                        "（点 mailto 只会弹出「选取应用」），"
                        "所以「用客户端手工发一封」这一步在本机走不通。两条替代路径：\n"
                        "      · 装上 Foxmail 并配好该账号，手工发一封测试邮件："
                        "能发出 → 属于客户端 / IP 白名单策略，请联系邮箱管理员"
                        "把本机出口 IP（或本工具）加入允许范围；\n"
                        "      · 或把本工具的「发送方式」切到「本机邮件客户端」，"
                        "由本工具把提醒导出成 .eml 后手动投递"
                        "（该模式不需要授权码 / 密码，可以先用起来）。\n")
    else:
        step4 = (head + "可以用邮件客户端（Foxmail / Outlook）手工配同一个账号"
                        "发一封测试邮件来判定：\n"
                        "      能发出 → 属于客户端白名单策略，请联系邮箱管理员"
                        "把本机出口 IP（或本工具）加入允许范围；\n"
                        "      同样发不出 → 回到 ① ②，问题在凭据本身。\n")

    return ("\n\n若这是单位自建的邮件系统（Coremail / 盈世、Exchange 等），"
            "认证被拒请按下面顺序自查，成本从低到高：\n"
            "  ① 密码是否已过期/被改过：先用浏览器登录网页版邮箱，确认当前密码。"
            "本工具里保存的若是很早以前设置的初始密码，很可能已被强制更换；\n"
            "  ② 是否需要「客户端专用密码 / 授权码」：登录网页版邮箱 → 设置 → "
            "安全设置（有的版本叫「客户端设置」「账户安全」）→ 生成"
            "「客户端专用密码 / 授权码」，填到本工具的「授权码/密码」框，"
            "不要把网页登录密码填进去；\n"
            "  ③ 发件服务器地址是否写对：自建邮局的发信入口通常是 MX 记录指向的那台主机"
            "（如 mailex1.你的域名.com:25 + STARTTLS），而不是 mail.你的域名.com。"
            "点击「连接诊断」可自动探测；\n"
            + step4
            + "  提示：网关会统计认证失败次数并按来源 IP 封禁，请勿连续重试。")


def is_gateway_auth_rejection(text: str) -> bool:
    """判断 535 认证错误的具体文本是否由安全网关回出（而非密码真的错）。

    典型：(535, b'reject by CACTER gateway, Error: Unknown error')
    这类返回是网关的策略拒绝，很可能账号密码本身没问题，直接提示用户改密码会误导。
    """
    t = (text or "").lower()
    keys = ("reject by", "cacter", "gateway", "is hangup",
            "auth error limit", "blacklist", "blocked", "denied", "too many")
    return any(k in t for k in keys)


def _auth_hint(server: str) -> str:
    """认证失败时按服务商给出可执行的处置步骤。"""
    s = (server or "").lower()
    if s.endswith("exmail.qq.com"):
        return ("\n\n腾讯企业邮箱必须使用「客户端专用密码」，不能直接用网页登录密码：\n"
                "  ① 网页版登录邮箱 → 设置 → 收发信设置 → 勾选开启 IMAP/SMTP 服务并保存；\n"
                "  ② 设置 → 邮箱绑定 → 开启安全登录 → 生成客户端专用密码（16 位）；\n"
                "  ③ 把该 16 位专用密码填入本工具「授权码/密码」框后重试；\n"
                "  ④ 企业邮箱还需管理员在管理后台开放该成员的 IMAP/SMTP 服务范围。")
    if s.endswith("smtp.qq.com"):
        return ("\n\nQQ 邮箱需填写「授权码」而非 QQ 密码：\n"
                "  网页版 设置 → 账户 → 开启 POP3/SMTP 服务 → 生成授权码。")
    if "163.com" in s or "126.com" in s:
        return ("\n\n网易邮箱需在网页版 设置 → POP3/SMTP/IMAP 中开启服务，"
                "并使用生成的「授权码」而非登录密码。")
    # 自建/Coremail 邮局：这是国内券商、高校、政企最常见的自建邮件系统
    if s.endswith(".local") or "." in s:
        return _selfhost_checklist()
    return ("\n\n若当前填写的是网页登录密码，请改用服务商提供的"
            "「授权码 / 客户端专用密码 / 应用专用密码」，并确认发件邮箱与 SMTP 服务器属于同一服务商。")


_PORT_ADVICE = {
    443: "443 是 HTTPS 端口，不是 SMTP 端口，请改用 465(SSL) 或 587(STARTTLS)。",
    80: "80 是 HTTP 端口，不是 SMTP 端口。",
    993: "993 是 IMAP 收信端口，发信请用 465(SSL) 或 587(STARTTLS)。",
    995: "995 是 POP3 收信端口，发信请用 465(SSL) 或 587(STARTTLS)。",
    143: "143 是 IMAP 收信端口，不能用于发信（本工具只需 SMTP）。",
    110: "110 是 POP3 收信端口，不能用于发信。",
}


def _describe_failure(cfg: SmtpConfig, exc: Exception) -> str:
    """把底层异常翻译成"哪一环断了 + 怎么办"的中文说明。"""
    base = f"邮件发送失败：{type(exc).__name__}: {exc}"
    advice: List[str] = []
    detail = str(exc)

    if isinstance(exc, socket.gaierror):
        advice.append(f"发信服务器域名「{cfg.server}」无法解析，请检查服务器地址是否写错。")
    elif isinstance(exc, (socket.timeout, TimeoutError)):
        advice.append(f"与 {cfg.server}:{cfg.port} 建立连接超时，说明这个地址/端口无法访问"
                      "（与账号密码无关）。")
        if cfg.port in _PORT_ADVICE:
            advice.append(_PORT_ADVICE[cfg.port])
        else:
            advice.append("请确认发信服务器与端口：公有邮箱一般是 465(SSL) / 587(STARTTLS)；"
                          "单位自建邮局（Coremail / 盈世等）往往只开放 25(STARTTLS)，"
                          "且服务器应是 MX 记录指向的那台主机（如 mailex1.你的域名.com），"
                          "而不是 mail.你的域名.com。")
        advice.append("点击「连接诊断」可自动定位是哪一环断了并给出推荐配置。")
    elif isinstance(exc, ConnectionRefusedError):
        advice.append(f"{cfg.server}:{cfg.port} 拒绝连接（端口未监听），端口或服务器地址可能不对。")
    elif isinstance(exc, smtplib.SMTPNotSupportedError):
        advice.append(f"服务器未提供 STARTTLS 能力，当前「STARTTLS」加密方式不适用。"
                      f"请改用 SSL（通常配 465），或把加密方式改为「不加密」（自建邮局 25 端口）。")
    elif isinstance(exc, smtplib.SMTPServerDisconnected):
        if "unexpectedly closed" in detail.lower() or "connection reset" in detail.lower():
            advice.append(
                f"{cfg.server}:{cfg.port} 接受了 TCP 连接但立刻断开、且没有返回任何服务号 —— "
                "这通常意味着该地址并不是邮件服务器（例如只是域名解析到的一台网关/负载均衡），"
                "或者连接被本机网络中的安全设备接管了。")
            advice.append("请点击「连接诊断」，它会查询发件域名的 MX 记录并逐一探测哪台主机"
                          "真正提供 SMTP 服务，然后一键套用推荐配置。")
        else:
            advice.append("服务器中途断开连接，可能是加密方式与端口不配套"
                          "（465 需配 SSL，587/25 需配 STARTTLS），或服务器限制了来源 IP。")
    elif isinstance(exc, smtplib.SMTPRecipientsRefused):
        bad = "、".join(exc.recipients.keys()) if exc.recipients else ""
        advice.append(f"服务器拒绝了收件人地址：{bad}。请检查目标邮箱是否拼写正确、是否存在。")
    elif isinstance(exc, ssl.SSLError):
        advice.append("TLS 握手失败：请核对加密方式与端口是否配套（465 ↔ SSL，587/25 ↔ STARTTLS）。")

    return base + ("\n\n可能原因与处理：\n  · " + "\n  · ".join(advice) if advice else "")


class EmailSender:
    """邮件发送器。"""

    def __init__(self, cfg: SmtpConfig) -> None:
        self._cfg = cfg

    # -------------------------------------------------------------- 对外接口 --

    def send(self, recipients: List[str], subject: str,
             html_body: str, plain_body: str) -> None:
        """发送 HTML + 纯文本双格式邮件。

        Raises:
            MailSendError: 重试后仍失败。
        """
        if not self._cfg.username or not self._cfg.password:
            raise MailSendError("发件人邮箱或授权码未配置，请先在界面保存邮箱配置")
        if not recipients:
            raise MailSendError("收件人列表为空，请先在界面保存目标邮箱")

        # 0. 反封禁冷却：刚刚因认证失败/被网关封禁，冷却期内不再触碰服务器。
        #    企业邮件网关会按 IP 累计 AUTH 失败次数并封禁，连续重试会让封禁加重。
        remaining = auth_block_remaining()
        if remaining > 0:
            raise MailSendError(
                f"已暂停发送：{AUTH_FAILURE_COOLDOWN_SECONDS // 60} 分钟内的账号认证失败后进入冷静期，"
                f"还有约 {remaining // 60} 分 {remaining % 60} 秒。\n\n"
                "邮件服务器（或其安全网关）会统计认证失败次数并按来源 IP 封禁，"
                "连续重试只会加重封禁，因此这里先停下。\n"
                "请先按提示修正账号/密码/服务器配置，保存后本工具会自动解除冷静期。"
            )

        # 1. 发送前先校验收件人格式：地址写错时邮件服务器一定拒收，
        #    提前拦截可避免"重试 3 次 → 仍失败 → 提醒项未标记已发"的无效等待
        invalid = find_invalid_recipients(recipients)
        if invalid:
            raise MailSendError(
                "目标邮箱格式不正确，未发送邮件：\n  · " + "\n  · ".join(invalid)
                + "\n\n请检查是否漏写 @ 或域名后缀（例如 .com），"
                  "到「邮箱配置」修正并保存后重试。"
            )

        msg = self._build_message(recipients, subject, html_body, plain_body)
        last_exc: Exception | None = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                self._smtp_send(recipients, msg)
                logger.info("邮件发送成功（第 %d 次尝试，收件人 %d 个）",
                            attempt, len(recipients))
                return
            except smtplib.SMTPAuthenticationError as exc:
                # 认证失败：重试无意义，且会触发网关封禁计数 → 直接失败并进入冷静期
                note_auth_failure()
                text = (exc.smtp_error.decode("utf-8", "replace")
                        if isinstance(exc.smtp_error, bytes) else str(exc.smtp_error))
                # 535 有两种性质完全不同的成因，必须分开说：
                #   · 网关策略拒绝（如 "reject by CACTER gateway"）→ 密码未必是错的
                #   · 凭据本身不对 → 直接引导换密码类型
                if is_gateway_auth_rejection(text):
                    raise MailSendError(
                        f"SMTP 认证被安全网关拒绝：{text.strip()}\n\n"
                        "注意：这条 535 是「安全网关」回的策略拒绝，"
                        "未必是账号密码写错，因此不要只是反复改密码重试"
                        "（重试会加重网关对本机 IP 的封禁）。"
                        + _selfhost_checklist()
                    ) from exc
                raise MailSendError(
                    f"SMTP 认证失败（服务器拒绝账号或密码）：{exc}"
                    + _auth_hint(self._cfg.server)
                ) from exc
            except smtplib.SMTPRecipientsRefused as exc:
                raise MailSendError(_describe_failure(self._cfg, exc)) from exc
            except (smtplib.SMTPNotSupportedError, smtplib.SMTPServerDisconnected) as exc:
                # 加密方式不配套 / 连上即断：配置问题，重试必然同样失败
                raise MailSendError(_describe_failure(self._cfg, exc)) from exc
            except smtplib.SMTPResponseException as exc:
                # 5xx 策略类拒绝（网关封禁、来源 IP 不允许、中继被拒等）：重试无意义
                text = exc.smtp_error.decode("utf-8", "replace") if isinstance(
                    exc.smtp_error, bytes) else str(exc.smtp_error)
                if is_gateway_rejection(text):
                    note_auth_failure()
                    raise MailSendError(
                        f"邮件安全网关拒绝了本次连接：{text.strip()}\n\n"
                        "该网关会按来源 IP 统计认证失败次数，超限即临时封禁（需等待自动解除，"
                        "或联系邮箱管理员解封），请勿连续重试。\n"
                        "若返回值里含 auth error limit / Hangup，说明本机出口 IP 正在封禁中，"
                        "通常等待一段时间后自动恢复。" + _auth_hint(self._cfg.server)
                    ) from exc
                raise MailSendError(
                    f"服务器拒绝（{exc.smtp_code}）：{text.strip()}"
                    + _auth_hint(self._cfg.server)
                ) from exc
            except (ssl.SSLError, socket.gaierror, ConnectionRefusedError) as exc:
                # 配置类错误：重试同样必然失败，直接给出分型说明
                raise MailSendError(_describe_failure(self._cfg, exc)) from exc
            except (smtplib.SMTPException, OSError) as exc:
                last_exc = exc
                if attempt < MAX_RETRIES:
                    delay = RETRY_DELAYS[min(attempt - 1, len(RETRY_DELAYS) - 1)]
                    logger.warning("第 %d 次发送失败（%s），%d 秒后重试",
                                   attempt, exc, delay)
                    time.sleep(delay)
        # 兜底：把 last_exc 兜成 Exception，避免类型检查困惑
        raise MailSendError(_describe_failure(self._cfg, last_exc or MailSendError("未知错误")))

    # -------------------------------------------------------------- 内部方法 --

    def _build_message(self, recipients: List[str], subject: str,
                       html_body: str, plain_body: str) -> MIMEMultipart:
        """构建 MIME 邮件（UTF-8）。"""
        return build_mime_message(self._cfg.username, recipients,
                                  subject, html_body, plain_body)

    def _smtp_send(self, recipients: List[str],
                   msg: MIMEMultipart) -> None:
        """按配置的加密模式建立连接并发送。

        三种模式：
          SSL      —— 465 端口，连接即 TLS（最通用）
          STARTTLS —— 587/25 端口，先明文再升级为 TLS（自建邮局常用）
          PLAIN    —— 25 端口且服务器不提供 STARTTLS 时使用，全程明文
        """
        cfg = self._cfg
        ctx = ssl.create_default_context()
        mode = (cfg.ssl_mode or "SSL").upper()
        if mode == "STARTTLS":
            with smtplib.SMTP(cfg.server, cfg.port,
                              timeout=SMTP_TIMEOUT_SECONDS) as smtp:
                smtp.ehlo()
                if not smtp.has_extn("starttls"):
                    # 服务器不支持 STARTTLS：明确报错，避免 smtplib 抛出难懂的异常
                    raise smtplib.SMTPNotSupportedError(
                        f"{cfg.server}:{cfg.port} 未提供 STARTTLS 能力")
                smtp.starttls(context=ctx)
                smtp.ehlo()
                smtp.login(cfg.username, cfg.password)
                smtp.send_message(msg)
        elif mode == "PLAIN":
            with smtplib.SMTP(cfg.server, cfg.port,
                              timeout=SMTP_TIMEOUT_SECONDS) as smtp:
                smtp.ehlo()
                smtp.login(cfg.username, cfg.password)
                smtp.send_message(msg)
        else:  # 默认 SSL
            with smtplib.SMTP_SSL(cfg.server, cfg.port, context=ctx,
                                  timeout=SMTP_TIMEOUT_SECONDS) as smtp:
                smtp.login(cfg.username, cfg.password)
                smtp.send_message(msg)


def send_test_email(cfg: SmtpConfig) -> None:
    """发送测试邮件，用于验证 SMTP 配置是否可用。"""
    sender = EmailSender(cfg)
    sender.send(
        recipients=cfg.recipients,
        subject="【权限到期提醒】测试邮件",
        html_body=(
            "<p>这是一封测试邮件，用于验证邮箱配置。</p>"
            "<p>收到本邮件说明 SMTP 配置正确。</p>"
        ),
        plain_body="这是一封测试邮件，用于验证邮箱配置。\n收到本邮件说明 SMTP 配置正确。",
    )


# ----------------------------------------------------------------- 邮件渲染 --

def _fmt_cn(d) -> str:
    """日期格式化为'某年某月某日'。"""
    return f"{d.year}年{d.month}月{d.day}日"


def render_summary_email(items: List[ReminderItem],
                         today) -> Tuple[str, str]:
    """把当日待提醒项渲染为汇总邮件（HTML, 纯文本）。

    Args:
        items: 已过滤去重后的待提醒项。
        today: 今天的日期。
    Returns:
        (html_body, plain_body) 二元组。
    """
    has_backfill = any(i.is_backfill for i in items)
    header = f"以下 {len(items)} 项客户权限即将到期，请及时跟进办理续期："
    if has_backfill:
        header += "（标注【补发】的为错过往常提醒时点的补发提醒）"

    # ---- HTML 部分 ----
    rows = []
    for it in items:
        backfill_tag = '<span style="color:#c0392b;">【补发】</span>' if it.is_backfill else ""
        rows.append(
            "<tr>"
            f"<td>{_esc(it.customer_id)}</td>"
            f"<td>{_esc(it.customer_name) or '&mdash;'}</td>"
            f"<td>{_esc(it.perm_label)}</td>"
            f"<td>{_fmt_cn(it.expiry)}</td>"
            f"<td>{_fmt_cn(it.trigger_date)}</td>"
            f"<td>{backfill_tag or '&mdash;'}</td>"
            "</tr>"
        )
    html_body = f"""\
<div style="font-family:'Microsoft YaHei',SimSun,sans-serif;max-width:720px;">
  <h3 style="color:#1a5276;">{SUBJECT}</h3>
  <p>{_esc(header)}</p>
  <table border="1" cellspacing="0" cellpadding="6"
         style="border-collapse:collapse;font-size:14px;">
    <tr style="background:#d6eaf8;font-weight:bold;">
      <td>客户号</td><td>客户姓名</td><td>权限</td>
      <td>到期日期</td><td>原提醒时点</td><td>备注</td>
    </tr>
    {''.join(rows)}
  </table>
  <p style="color:#7f8c8d;font-size:12px;">
    本邮件由权限到期提醒工具于 {_fmt_cn(today)} 自动发送，请勿直接回复。
  </p>
</div>"""

    # ---- 纯文本部分 ----
    lines = [SUBJECT, "", header, ""]
    for it in items:
        tag = "【补发】" if it.is_backfill else ""
        lines.append(
            f"{tag}客户号：{it.customer_id}　客户姓名：{it.customer_name or '—'}　"
            f"{it.perm_label}　到期日期：{_fmt_cn(it.expiry)}"
        )
    lines += ["", f"本邮件由权限到期提醒工具于 {_fmt_cn(today)} 自动发送，请勿直接回复。"]
    plain_body = "\n".join(lines)

    return html_body, plain_body


def _esc(s: str) -> str:
    """HTML 转义（姓名等来自 Excel 的自由文本）。"""
    return (s.replace("&", "&amp;").replace("<", "&lt;")
             .replace(">", "&gt;").replace('"', "&quot;"))
