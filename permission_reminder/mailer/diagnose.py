# -*- coding: utf-8 -*-
"""SMTP 发信链路诊断模块。

用途：当"发送测试邮件""检查并发送提醒"报错时，用这一套检查把"到底哪一环断了"
定位到具体环节，而不是只给出一句 "timed out"。

诊断链条（逐环独立、失败不中断后续可执行环节）：
    [1] 配置完整性
    [2] 端口合理性（443/80/993/143/110 等常见误填）与加密方式是否配套
    [3] 发信服务器域名解析
    [4] 配置端口 TCP 连通性 + **服务号探测**（区分"真有邮件服务"与"连上却零响应"）
    [5] 本机出站基线对照（区分"目标主机不可达"与"本机被封 SMTP 出站"）
    [6] 发件域 MX 记录（纯 UDP DNS 查询，不依赖第三方库）
    [7] **MX 主机端点探测**（逐一实测哪台主机真的提供 SMTP，这类自建邮局的关键）
    [8] 服务商识别 + 基于实测的推荐发信端点
    [9] TLS 握手 + 账号登录验证（仅在服务号正常时执行，绝不会在已被封禁时提交密码）
    [10] 收件人地址格式

最后给出中文结论与可执行的修复建议，可选的"推荐 SMTP 配置"可直接被 UI 一键套用。
"""
from __future__ import annotations

import logging
import random
import socket
import smtplib
import ssl
import struct
import time
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Optional, Sequence, Tuple

from ..infra.config_store import SmtpConfig
from .sender import find_invalid_recipients

logger = logging.getLogger(__name__)

# ------------------------------------------------------------------ 常量表 --

# 公共 DNS（用于 MX 查询；按国内可达性排序，逐个尝试）
DNS_SERVERS: Tuple[str, ...] = ("223.5.5.5", "119.29.29.29", "114.114.114.114")

# SMTP 标准端口
STANDARD_SMTP_PORTS: Tuple[int, ...] = (465, 587, 25, 2525)

_PORT_ROLE = {
    465: "SSL（隐式 TLS，最通用）",
    587: "STARTTLS（邮件提交端口）",
    25: "明文/STARTTLS（多被运营商与云主机封禁）",
    2525: "STARTTLS（备用提交端口）",
}

# 常见端口误填提示
PORT_HINTS = {
    443: "443 是 HTTPS 端口，不是 SMTP 端口。除非服务商明确说明，否则请改用 465(SSL) 或 587(STARTTLS)。",
    80: "80 是 HTTP 端口，不是 SMTP 端口。请改用 465(SSL) 或 587(STARTTLS)。",
    993: "993 是 IMAP 收信端口。发信请用 465(SSL) 或 587(STARTTLS)。",
    995: "995 是 POP3 收信端口。发信请用 465(SSL) 或 587(STARTTLS)。",
    143: "143 是 IMAP 端口，不能用于发信。",
    110: "110 是 POP3 端口，不能用于发信。",
    8443: "8443 是 HTTPS 备用端口，不是 SMTP 端口。",
}

# 出站基线对照用的公共发信服务器
_BASELINE_TARGETS: Tuple[Tuple[str, int], ...] = (
    ("smtp.qq.com", 465),
    ("smtp.163.com", 465),
)


@dataclass
class DiagStep:
    """单个诊断环节的结果。"""
    name: str
    status: str          # ok / warn / fail / info / skip
    detail: str

    @property
    def mark(self) -> str:
        return {"ok": "OK  ", "warn": "警告", "fail": "失败",
                "info": "信息", "skip": "跳过"}.get(self.status, "信息")


@dataclass
class ProviderHint:
    """根据 MX 记录推断出的服务商与推荐发信配置。"""
    name: str
    server: str
    port: int
    ssl_mode: str
    reason: str


@dataclass
class SmtpDiagnosis:
    """一次完整诊断的结果。"""
    steps: List[DiagStep] = field(default_factory=list)
    ok: bool = False
    conclusion: str = ""
    suggestions: List[str] = field(default_factory=list)
    recommended: Optional[Tuple[str, int, str]] = None   # (server, port, ssl_mode)
    provider: Optional[ProviderHint] = None

    def add(self, name: str, status: str, detail: str) -> None:
        self.steps.append(DiagStep(name, status, detail))


# ------------------------------------------------------- MX 查询（std-lib） --

def _build_mx_query(domain: str, txid: int) -> bytes:
    """构造 DNS MX 查询报文（QTYPE=15）。"""
    header = struct.pack(">HHHHHH", txid, 0x0100, 1, 0, 0, 0)  # RD=1
    qname = b"".join(bytes([len(p)]) + p.encode("ascii")
                     for p in domain.split(".") if p) + b"\x00"
    return header + qname + struct.pack(">HH", 15, 1)


def _read_name(data: bytes, off: int) -> Tuple[str, int]:
    """解析（可能带压缩指针的）域名，返回 (域名, 新偏移)。"""
    labels: List[str] = []
    jumps = 0
    while off < len(data):
        ln = data[off]
        if ln == 0:
            off += 1
            break
        if ln & 0xC0 == 0xC0:                     # 压缩指针
            if off + 1 >= len(data):
                break
            ptr = struct.unpack(">H", data[off:off + 2])[0] & 0x3FFF
            jumps += 1
            if jumps > 8:                          # 防御环形指针
                break
            sub, _ = _read_name(data, ptr)
            labels.append(sub)
            off += 2
            break
        labels.append(data[off + 1:off + 1 + ln].decode("ascii", "replace"))
        off += 1 + ln
    return ".".join(p for p in labels if p), off


def _parse_mx_response(data: bytes) -> List[Tuple[int, str]]:
    """从 DNS 响应报文中提取 MX 记录，返回 [(优先级, 主机名)]（按优先级升序）。

    Raises:
        ValueError: 报文过短/格式非法。
    """
    if len(data) < 12:
        raise ValueError("DNS 响应过短")
    _txid, flags, qdcount, ancount, _ns, _ar = struct.unpack(">HHHHHH", data[:12])
    if flags & 0x000F:                            # RCODE != 0
        raise ValueError(f"DNS 返回错误码 {flags & 0x000F}")
    off = 12
    for _ in range(qdcount):                      # 跳过 question 段
        _, off = _read_name(data, off)
        off += 4
    out: List[Tuple[int, str]] = []
    for _ in range(ancount):
        if off + 10 > len(data):
            break
        _, off = _read_name(data, off)
        rtype, _rclass, _ttl, rdlen = struct.unpack(">HHIH", data[off:off + 10])
        off += 10
        end = min(off + rdlen, len(data))
        if rtype == 15 and rdlen >= 3:
            pref = struct.unpack(">H", data[off:off + 2])[0]
            name, _ = _read_name(data, off + 2)
            if name:
                out.append((pref, name))
        off = end
    return sorted(out)


def mx_lookup(domain: str, timeout: float = 5.0) -> List[Tuple[int, str]]:
    """查询域名 MX 记录（纯 stdlib UDP DNS，失败返回空列表，不抛异常）。

    使用原始 DNS 报文而非 DNS-over-HTTPS：部分内网环境会拦截/替换 HTTPS 证书链，
    而 UDP 53 通常可用；且无需引入 dnspython 等第三方依赖。
    """
    if not domain:
        return []
    txid = random.randint(0, 0xFFFF)
    packet = _build_mx_query(domain, txid)
    for server in DNS_SERVERS:
        sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        sock.settimeout(timeout)
        try:
            sock.sendto(packet, (server, 53))
            data, _ = sock.recvfrom(4096)
            if data[:2] != packet[:2]:            # 事务 ID 不匹配，丢弃
                continue
            return _parse_mx_response(data)
        except Exception as exc:                  # noqa: BLE001 —— 逐个 DNS 降级
            logger.debug("MX 查询经 %s 失败：%s", server, exc)
        finally:
            sock.close()
    logger.warning("MX 查询失败（所有 DNS 均不可达），跳过服务商识别")
    return []


def provider_from_mx(mx_hosts: Sequence[str], domain: str) -> Optional[ProviderHint]:
    """根据 MX 记录推断服务商（顺序敏感，先精确后宽泛）。

    注意：**不要**仅凭主机名前缀做判断。例如 "mailex1.<域名>" 曾经被误认为是
    腾讯企业邮箱的特征，但实际上国内大量自建 Coremail 邮局也用 mailexN 命名。
    因此这里只认明确的托管商域名后缀；落在自有域名下的 MX 一律按"自建邮局"处理，
    最终以 probe_mx_endpoints() 的实测结果为准。
    """
    hosts = [h.lower().rstrip(".") for h in mx_hosts if h]
    if not hosts:
        return None

    for h in hosts:
        if h.endswith("exmail.qq.com"):
            return ProviderHint("腾讯企业邮箱", "smtp.exmail.qq.com", 465, "SSL",
                                f"MX 记录 {h} 指向腾讯企业邮箱")
    for h in hosts:
        if h.endswith("mxhichina.com") or h.endswith("qiye.aliyun.com"):
            return ProviderHint("阿里云企业邮箱", "smtp.qiye.aliyun.com", 465, "SSL",
                                f"MX 记录 {h} 指向阿里云企业邮箱")
    for h in hosts:
        if h.endswith("mxmail.netease.com"):
            return ProviderHint("网易(企业)邮箱", "smtp.163.com", 465, "SSL",
                                f"MX 记录 {h} 指向网易企业邮箱")
    for h in hosts:
        if h.endswith("larksuite.com") or h.endswith("feishu.cn"):
            return ProviderHint("飞书企业邮箱", "smtp.feishu.cn", 465, "SSL",
                                f"MX 记录 {h} 指向飞书邮箱")
    for h in hosts:
        if h.endswith("protection.outlook.com") or h.endswith("office365.com"):
            return ProviderHint("Microsoft 365", "smtp.office365.com", 587, "STARTTLS",
                                f"MX 记录 {h} 指向 Microsoft 365")
    for h in hosts:
        if h.endswith("google.com") or h.endswith("googlemail.com"):
            return ProviderHint("Google Workspace", "smtp.gmail.com", 465, "SSL",
                                f"MX 记录 {h} 指向 Google Workspace")
    for h in hosts:
        # 腾讯企业邮箱的公开 MX 形如 mxbiz1.qq.com；个人 QQ 邮箱为 mx1.qq.com
        if h.startswith("mxbiz") and h.endswith("qq.com"):
            return ProviderHint("腾讯企业邮箱", "smtp.exmail.qq.com", 465, "SSL",
                                f"MX 记录 {h} 指向腾讯企业邮箱")
    for h in hosts:
        if h.endswith("qq.com"):
            return ProviderHint("腾讯个人邮箱", "smtp.qq.com", 465, "SSL",
                                f"MX 记录 {h} 指向腾讯邮箱")

    # 自建/自有邮件服务器：MX 落在自己域名下（如 mailex1.example.com）。
    # 这类系统（Coremail、Exchange、Postfix 等）的发信服务器就是 MX 主机本身，
    # 端口与加密方式必须实测，故这里给出"候选"而非定论。
    for h in hosts:
        if domain and h.endswith(domain):
            return ProviderHint("自建/自有邮件服务器", h, 25, "STARTTLS",
                                f"MX 记录 {h} 在自有域名下（国内自建邮局常见 Coremail + 安全网关），"
                                "发信服务器通常就是这台 MX 主机")
    return None


# --------------------------------------------------------------- 网络探测 --

def _tcp_probe(host: str, port: int, timeout: float = 6.0) -> Tuple[str, float, str]:
    """TCP 连通性探测。

    Returns:
        (状态, 耗时秒, 说明)；状态 ∈ {'open','timeout','refused','dns','error'}
    """
    t0 = time.time()
    sock = None
    try:
        sock = socket.create_connection((host, port), timeout=timeout)
        return "open", time.time() - t0, "可连接"
    except socket.gaierror as exc:
        return "dns", time.time() - t0, f"域名无法解析：{exc}"
    except socket.timeout:
        return "timeout", time.time() - t0, "超时无响应（防火墙丢包、端口未开放或服务未监听）"
    except ConnectionRefusedError:
        return "refused", time.time() - t0, "连接被拒绝（端口未监听）"
    except Exception as exc:                      # noqa: BLE001
        return "error", time.time() - t0, f"{type(exc).__name__}: {exc}"
    finally:
        if sock is not None:
            sock.close()


@dataclass
class SmtpProbe:
    """对某个 (主机, 端口) 的 SMTP 服务探测结果。"""
    host: str
    port: int
    status: str          # smtp / reject / silent / timeout / refused / dns / error
    banner: str = ""     # 服务号原文（已 strip）
    caps: List[str] = field(default_factory=list)   # EHLO 能力关键字（统一小写）
    auth_mechs: List[str] = field(default_factory=list)  # AUTH 支持的机制
    elapsed: float = 0.0

    @property
    def label(self) -> str:
        return {
            "smtp": "有 SMTP 服务",
            "reject": "拒绝/封禁",
            "silent": "连上但无响应",
            "timeout": "超时",
            "refused": "拒绝连接",
            "dns": "域名无法解析",
            "error": "异常",
        }.get(self.status, self.status)

    @property
    def usable(self) -> bool:
        """该端点是否为可用的 SMTP 服务（返回 220 服务号，且不是网关拒绝）。"""
        return self.status == "smtp"

    @property
    def supports_starttls(self) -> bool:
        return "starttls" in self.caps

    @property
    def summary(self) -> str:
        if self.status == "smtp":
            cap = "、".join(c.upper() for c in self.caps) if self.caps else "（未取到能力列表）"
            return (f"{self.host}:{self.port} 有 SMTP 服务"
                    f"（服务号 {self.banner[:50]}；能力：{cap}）")
        if self.status == "reject":
            return f"{self.host}:{self.port} 返回拒绝信息：{self.banner[:120]}"
        if self.status == "silent":
            return (f"{self.host}:{self.port} TCP 连上了，但 {self.elapsed:.0f} 秒内未返回任何服务号"
                    "（该地址很可能不是邮件服务器，或连接被安全设备接管）")
        return f"{self.host}:{self.port} {self.label}（{self.elapsed:.0f}s）"


def smtp_probe(host: str, port: int, timeout: float = 8.0,
               do_ehlo: bool = True) -> SmtpProbe:
    """探测某主机的某端口是否真的提供 SMTP 服务。

    这一步不能只用 TCP 连通性替代：企业网络中常见"TCP 能连上、但对方一个字节都不返回"
    的情况（安全设备接管连接、或该域名只是负载均衡前端），仅凭 TCP 会误判为"网络正常"。
    因此这里会主动读取服务号，并可选发送 EHLO 取得能力列表。

    Returns:
        SmtpProbe；状态 ∈ {'smtp','reject','silent','timeout','refused','dns','error'}
    """
    t0 = time.time()
    try:
        sock = socket.create_connection((host, port), timeout=timeout)
    except socket.gaierror:
        return SmtpProbe(host, port, "dns", elapsed=time.time() - t0)
    except socket.timeout:
        return SmtpProbe(host, port, "timeout", elapsed=time.time() - t0)
    except ConnectionRefusedError:
        return SmtpProbe(host, port, "refused", elapsed=time.time() - t0)
    except Exception:                              # noqa: BLE001
        return SmtpProbe(host, port, "error", elapsed=time.time() - t0)

    probe = SmtpProbe(host, port, "error", elapsed=time.time() - t0)
    try:
        sock.settimeout(timeout)
        try:
            raw = sock.recv(512)
        except (socket.timeout, TimeoutError):
            probe.status = "silent"
            probe.elapsed = time.time() - t0
            return probe
        except (ConnectionResetError, OSError):
            probe.status = "silent"
            probe.elapsed = time.time() - t0
            return probe

        probe.banner = raw.decode("utf-8", "replace").strip()
        code = probe.banner[:3]
        if code.startswith("2"):
            probe.status = "smtp"
        elif code.isdigit() and code[0] in "45":
            probe.status = "reject"       # 如 554 reject by CACTER gateway ...
        else:
            probe.status = "silent"       # 不是 SMTP 协议（例如 HTTP/IMAP 服务号）

        if probe.status == "smtp" and do_ehlo:
            try:
                sock.sendall(b"EHLO diag.probe\r\n")
                data = b""
                while True:               # 多行应答以 "250 " 结束
                    chunk = sock.recv(1024)
                    if not chunk:
                        break
                    data += chunk
                    if b"\r\n250 " in data or b"\n250 " in data:
                        break
                text = data.decode("utf-8", "replace")
                for raw_line in text.splitlines():
                    line = raw_line.strip()
                    if len(line) < 4:
                        continue
                    body = line[4:].strip()          # 去掉 "250-"/"250 " 前缀
                    if not body:
                        continue
                    kw = body.split()[0].lower()
                    if kw not in probe.caps:
                        probe.caps.append(kw)
                    if kw == "auth":
                        probe.auth_mechs = [m.upper() for m in body.split()[1:]]
            except Exception:          # noqa: BLE001 —— 能力列表取不到不影响结论
                pass
    finally:
        try:
            sock.close()
        except OSError:
            pass
    probe.elapsed = time.time() - t0
    return probe


# MX 主机上值得一试的端口（自建邮局常见 25；托管商常见 465/587）
MX_PROBE_PORTS: Tuple[int, ...] = (25, 465, 587)


def probe_mx_endpoints(mx_hosts: Sequence[str], ports: Sequence[int] = MX_PROBE_PORTS,
                       timeout: float = 6.0, max_hosts: int = 3) -> List[SmtpProbe]:
    """对 MX 主机逐个探测，找出真正提供 SMTP 服务、可用于发信的端点。

    只对每个主机做一轮探测，且不做任何认证尝试 —— 不会增加安全网关的失败计数。
    """
    results: List[SmtpProbe] = []
    for host in list(mx_hosts)[:max_hosts]:
        h = (host or "").strip().rstrip(".")
        if not h:
            continue
        for port in ports:
            probe = smtp_probe(h, port, timeout=timeout, do_ehlo=True)
            results.append(probe)
            if probe.status == "smtp" and port == 25:
                break              # 25 已可用则不必再扫该主机的其他端口
    return results


def _try_login(cfg: SmtpConfig, timeout: float = 20.0) -> Tuple[str, str]:
    """尝试建立 SMTP 连接并登录（不发送任何邮件）。

    Returns:
        (状态, 说明)；状态 ∈ {'ok','auth','tls','net','smtp','skip'}
    """
    ctx = ssl.create_default_context()
    mode = (cfg.ssl_mode or "SSL").upper()
    try:
        if mode == "SSL":
            with smtplib.SMTP_SSL(cfg.server, cfg.port, context=ctx,
                                  timeout=timeout) as smtp:
                smtp.login(cfg.username, cfg.password)
        else:
            with smtplib.SMTP(cfg.server, cfg.port, timeout=timeout) as smtp:
                smtp.ehlo()
                if mode == "STARTTLS":
                    if not smtp.has_extn("starttls"):
                        return ("tls", "服务器未提供 STARTTLS 能力，"
                                       "该端口应改用 SSL(465) 或把加密方式改为「不加密」")
                    smtp.starttls(context=ctx)
                    smtp.ehlo()
                smtp.login(cfg.username, cfg.password)
        return "ok", "连接、加密与账号认证全部成功"
    except smtplib.SMTPAuthenticationError as exc:
        code = exc.smtp_code
        err = exc.smtp_error
        err = err.decode("utf-8", "replace") if isinstance(err, bytes) else str(err)
        return "auth", f"服务器返回 {code}：{err.strip()}"
    except smtplib.SMTPResponseException as exc:
        err = exc.smtp_error
        err = err.decode("utf-8", "replace") if isinstance(err, bytes) else str(exc.smtp_error)
        return "reject", f"服务器返回 {exc.smtp_code}：{err.strip()}"
    except ssl.SSLError as exc:
        return "tls", f"TLS 握手失败：{exc}"
    except (socket.timeout, TimeoutError):
        return "net", "连接/登录阶段超时"
    except (smtplib.SMTPException, OSError) as exc:
        return "smtp", f"{type(exc).__name__}: {exc}"


def _auth_advice(server: str, provider: Optional[ProviderHint]) -> List[str]:
    """按服务商给出认证失败的具体处置步骤。"""
    s = (server or "").lower()
    if provider is not None and provider.name == "自建/自有邮件服务器":
        return [_SELFHOSTED_AUTH_ADVICE]
    if provider is not None and provider.name in ("腾讯企业邮箱", "腾讯个人邮箱"):
        step4 = ("\n  ④ 若为企业邮箱，需管理员在管理后台开放该成员的 IMAP/SMTP 服务范围。"
                 if provider.name == "腾讯企业邮箱" else "")
        return ["腾讯邮箱必须使用「客户端专用密码 / 授权码」，不能直接填网页登录密码：\n"
                "  ① 网页版登录邮箱 → 设置 → 收发信设置 → 勾选开启 IMAP/SMTP 服务并保存；\n"
                "  ② 设置 → 邮箱绑定 → 开启安全登录 → 生成客户端专用密码（16 位）；\n"
                "  ③ 把该 16 位专用密码填入本工具「授权码/密码」框，重新保存后测试。"
                + step4]
    if s.endswith("exmail.qq.com"):
        return ["腾讯企业邮箱需使用「客户端专用密码」：\n"
                "  ① 网页版 设置 → 收发信设置 → 开启 IMAP/SMTP 服务；\n"
                "  ② 设置 → 邮箱绑定 → 开启安全登录 → 生成客户端专用密码（16 位）；\n"
                "  ③ 填入本工具密码框后重新保存。"]
    if s.endswith("smtp.qq.com"):
        return [
            "QQ 邮箱需使用「授权码」：网页版 设置→账户→POP3/SMTP 服务 开启后生成授权码，",
            "填入本工具密码框（不是 QQ 密码）。",
        ]
    if "163.com" in s or "126.com" in s:
        return [
            "网易邮箱需在 网页版 设置→POP3/SMTP/IMAP 中开启服务，并使用生成的「授权码」而非登录密码。",
        ]
    return [_SELFHOSTED_AUTH_ADVICE]


# 自建邮局（Coremail / Exchange / Postfix 等）认证被拒的统一处置说明
_SELFHOSTED_AUTH_ADVICE = (
    "这是单位自建的邮件系统，认证被拒通常是以下三类原因（按可能性排序）：\n"
    "  ① 密码类型不对：管理员可能要求第三方客户端使用「客户端专用密码 / 授权码」，"
    "而非网页登录密码。请登录网页版邮箱 → 设置 → 安全设置/客户端设置，"
    "查看是否有「客户端专用密码 / 授权码」并生成一个填入本工具；\n"
    "  ② 发件服务器地址不对：自建邮局的发信服务器通常是 MX 记录指向的那台主机"
    "（本次诊断的「MX 端点」环节已实测），而不是 mail.<你的域名> 这种前端地址；\n"
    "  ③ 客户端访问被限制：部分单位启用了邮件安全网关（如 CACTER），"
    "只允许白名单客户端或仅允许内网 IP 认证，需联系邮箱管理员把你的账号/出口 IP 加入允许范围。\n"
    "  注意：这类网关会按来源 IP 统计认证失败次数并临时封禁，请勿连续重试。"
)


# ------------------------------------------------------------------ 主流程 --

def diagnose_smtp(cfg: SmtpConfig, timeout: float = 6.0,
                  skip_login: bool = False) -> SmtpDiagnosis:
    """对 SMTP 配置执行完整链路诊断（只读，不发送邮件）。"""
    diag = SmtpDiagnosis()
    server = (cfg.server or "").strip()
    port = int(cfg.port or 0)
    domain = cfg.username.split("@")[-1].strip().lower() if "@" in (cfg.username or "") else ""

    # ---------------- [1] 配置完整性 ----------------
    missing = []
    if not server:
        missing.append("SMTP 服务器")
    if not port:
        missing.append("端口")
    if not cfg.username:
        missing.append("发件人邮箱")
    if not cfg.password:
        missing.append("授权码/密码")
    if missing:
        diag.add("配置完整性", "fail", "缺少：" + "、".join(missing))
        diag.conclusion = "配置不完整，缺少：" + "、".join(missing) + "。请先在界面补齐后再诊断。"
        return diag
    mode = cfg.ssl_mode or "SSL"
    diag.add("配置完整性", "ok",
             f"{server}:{port}（{mode}）　发件人 {cfg.username}　收件人 {len(cfg.recipients)} 个")

    # ---------------- [2] 端口合理性 ----------------
    if port in PORT_HINTS:
        diag.add("端口合理性", "warn", PORT_HINTS[port])
    elif port in _PORT_ROLE:
        diag.add("端口合理性", "ok", f"{port} 是标准 SMTP 端口：{_PORT_ROLE[port]}")
    else:
        diag.add("端口合理性", "warn",
                 f"{port} 不是标准 SMTP 端口（标准为 465/587/25/2525），请确认服务商是否特别说明")
    # 加密方式与端口是否般配（服务商可能不支持切换）
    if port == 465 and mode != "SSL":
        diag.add("加密匹配", "warn", f"465 端口通常必须用 SSL；当前选的是 {mode}，建议改为 SSL")
    elif port == 587 and mode == "SSL":
        diag.add("加密匹配", "warn", "587 端口通常必须用 STARTTLS；当前选的是 SSL，建议改为 STARTTLS")
    elif port == 25 and mode == "SSL":
        diag.add("加密匹配", "warn",
                 "25 端口不支持 SSL（隐式 TLS，连接即加密），请改为 STARTTLS；"
                 "若服务器不支持 STARTTLS，则改用「不加密」")

    # ---------------- [3] 域名解析 ----------------
    try:
        ips = sorted({info[4][0] for info in socket.getaddrinfo(server, port)})
        diag.add("域名解析", "ok", f"{server} → {', '.join(ips)}")
    except Exception as exc:                      # noqa: BLE001
        diag.add("域名解析", "fail", f"{server} 无法解析：{exc}")
        diag.conclusion = f"发信服务器域名 {server} 无法解析，请检查服务器地址是否写错。"
        return diag

    # ---------------- [4] 端口连通性 + 服务号探测（连接真实性）----------------
    status, elapsed, note = _tcp_probe(server, port, timeout)
    probe: Optional[SmtpProbe] = None
    if status == "open":
        diag.add("端口连通性", "ok", f"{server}:{port} 可连接（{elapsed:.2f}s），TCP 层正常")
        # TCP 通 ≠ 这里有邮件服务：企业网络中常见"连得上但一个字节都不返回"
        probe = smtp_probe(server, port, timeout=max(timeout, 8.0))
        if probe.status == "smtp":
            diag.add("服务号探测", "ok", probe.summary)
        elif probe.status == "reject":
            diag.add("服务号探测", "fail", probe.summary)
        else:
            diag.add("服务号探测", "warn", probe.summary)
        # 顺带看这台主机其他标准端口，帮助判断"只是端口填错"还是"整台主机没有 SMTP"
        others = []
        for alt in STANDARD_SMTP_PORTS:
            if alt == port:
                continue
            st2, _, _ = _tcp_probe(server, alt, min(timeout, 4.0))
            others.append(f"{alt}={'可连' if st2 == 'open' else '不通'}")
        diag.add("同主机其他端口", "info", f"{server}：{'、'.join(others)}")
    else:
        diag.add("端口连通性", "fail", f"{server}:{port} {note}（{elapsed:.2f}s）")
        others = []
        for alt in STANDARD_SMTP_PORTS:
            if alt == port:
                continue
            st2, _, _ = _tcp_probe(server, alt, min(timeout, 4.0))
            others.append(f"{alt}={'可连' if st2 == 'open' else '不通'}")
        diag.add("同主机其他端口", "info", f"{server}：{'、'.join(others)}")

    # ---------------- [5] 本机出站基线对照 ----------------
    # 目标不通、或"连得上却无服务号"时都值得对照：后者常常是网络层安全设备造成的
    baseline_ok: List[str] = []
    if status != "open" or (probe is not None and not probe.usable):
        for host, bport in _BASELINE_TARGETS:
            st3, _, _ = _tcp_probe(host, bport, min(timeout, 5.0))
            if st3 == "open":
                baseline_ok.append(f"{host}:{bport}")
        # 基线用公共服务器做「服务号」级别的验证，比 TCP 更能说明问题
        baseline_probe = smtp_probe("smtp.qq.com", 25, timeout=min(timeout, 8.0))
        if baseline_ok:
            extra = ("；且 smtp.qq.com:25 能取到真实 SMTP 服务号，说明本机网络本身允许 SMTP"
                     if baseline_probe.usable else "")
            diag.add("本机出站基线", "ok",
                     f"{'、'.join(baseline_ok)} 均可连接{extra} → 问题出在目标服务器地址/端口，"
                     "而不是本机网络封禁")
        else:
            diag.add("本机出站基线", "fail",
                     "公共 SMTP 服务器（smtp.qq.com:465 等）同样连不上 → 可能是本机防火墙、"
                     "上网行为管理或代理策略拦截了对外 SMTP 端口")

    # ---------------- [6] MX 记录 ----------------
    mx_hosts: List[str] = []
    if domain:
        mxs = mx_lookup(domain)
        if mxs:
            mx_hosts = [h for _, h in mxs]
            diag.add("MX 记录", "ok", f"{domain} 的 MX：" + "、".join(mx_hosts))
        else:
            diag.add("MX 记录", "skip", f"无法查询 {domain} 的 MX（DNS 不可达），跳过端点探测")

    # ---------------- [7] MX 主机端点探测（找出真正的发信服务器）----------------
    mx_probes: List[SmtpProbe] = []
    if mx_hosts:
        mx_probes = probe_mx_endpoints(mx_hosts, timeout=min(timeout, 6.0))
        for p in mx_probes:
            diag.add(f"MX 端点 {p.host}:{p.port}",
                     "ok" if p.usable else "info", p.summary)

    # ---------------- [8] 服务商识别与推荐配置 ----------------
    prov = provider_from_mx(mx_hosts, domain) if mx_hosts else None
    if prov is not None:
        diag.provider = prov
        diag.add("服务商识别", "info", f"{prov.name}（依据：{prov.reason}）")
    elif mx_hosts:
        diag.add("服务商识别", "info", "MX 未匹配到已知托管服务商")

    best = _pick_best_endpoint(mx_probes)
    if best is not None:
        best_mode = ("STARTTLS" if best.supports_starttls
                     else "PLAIN" if best.port == 25 else "SSL")
        same = (best.host.lower() == server.lower() and best.port == port)
        diag.add("推荐发信端点", "ok" if same else "warn",
                 f"实测可用：{best.host}:{best.port}（{best_mode}）"
                 + ("，与当前配置一致" if same else "，与当前配置不同"))
        if not same:
            diag.recommended = (best.host, best.port, best_mode)

    # ---------------- [9] TLS + 登录验证 ----------------
    login_status = "skip"
    login_detail = "端口不通，未尝试登录"
    if skip_login:
        login_detail = "已按设置跳过登录验证"
    elif status != "open":
        pass
    elif probe is not None and probe.status == "reject":
        # 连接阶段就被网关拒绝：再提交账号密码只会加重封禁，绝不尝试
        login_status = "ban"
        login_detail = f"连接阶段即被拒绝，未尝试登录：{probe.banner[:140]}"
    elif probe is not None and not probe.usable:
        login_detail = "该端点未返回 SMTP 服务号，未尝试登录（避免无意义的等待与账号锁定）"
    else:
        login_status, login_detail = _try_login(cfg)
    diag.add("登录验证", {"ok": "ok", "auth": "fail", "tls": "fail", "net": "fail",
                          "smtp": "fail", "ban": "fail"}.get(login_status, "skip"),
             login_detail)

    # ---------------- [10] 收件人格式 ----------------
    invalid = find_invalid_recipients(cfg.recipients)
    if not cfg.recipients:
        diag.add("收件人格式", "fail", "未配置目标邮箱")
    elif invalid:
        diag.add("收件人格式", "fail", "格式不正确：" + "、".join(invalid)
                 + "（缺少 @ 或域名后缀，如 .com）")
    else:
        diag.add("收件人格式", "ok", f"{len(cfg.recipients)} 个收件人格式正确")

    # ---------------- 结论与建议 ----------------
    _compose_verdict(diag, cfg, status, note, login_status, login_detail, invalid,
                     baseline_ok, probe, mx_probes)
    return diag


def _pick_best_endpoint(probes: Sequence[SmtpProbe]) -> Optional[SmtpProbe]:
    """从探测结果中挑出最适合发信的端点。

    优先级：25+STARTTLS（自建邮局最标准）> 465 > 其他可用端口。
    """
    good = [p for p in probes if p.usable]
    if not good:
        return None

    def rank(p: SmtpProbe) -> Tuple[int, int]:
        if p.port == 25 and p.supports_starttls:
            return (0, 0)
        if p.port == 465:
            return (1, 0)
        if p.port == 25:
            return (2, 0)
        return (3, p.port)

    return sorted(good, key=rank)[0]


def _compose_verdict(diag: SmtpDiagnosis, cfg: SmtpConfig, tcp_status: str,
                     tcp_note: str, login_status: str, login_detail: str,
                     invalid: List[str], baseline_ok: List[str],
                     probe: Optional[SmtpProbe] = None,
                     mx_probes: Optional[List[SmtpProbe]] = None) -> None:
    """根据各环节结果生成结论与建议。"""
    server, port, mode = cfg.server, int(cfg.port or 0), cfg.ssl_mode
    mx_probes = mx_probes or []
    rec = diag.recommended

    if tcp_status != "open":
        diag.conclusion = (
            f"连接失败（timed out/refused）与账号密码无关：{server}:{port} 这个发信端点本身不可达。"
        )
        if port in PORT_HINTS:
            diag.suggestions.append(PORT_HINTS[port])
        if rec is not None:
            diag.suggestions.append(
                f"改用实测可用的发信端点 {rec[0]}:{rec[1]}（{rec[2]}）—— 诊断时该端点返回了"
                "真实的 SMTP 服务号，可直接点「应用推荐配置」。")
        elif diag.provider is not None:
            p = diag.provider
            diag.suggestions.append(
                f"候选发信服务器：{p.server}:{p.port}（{p.ssl_mode}）—— {p.reason}。")
        else:
            diag.suggestions.append(
                "向邮箱管理员确认正确的发信服务器地址与端口；标准做法是 smtp.<你的域名> "
                "或服务商给出的 smtp 主机名，端口 465(SSL)/587(STARTTLS)、自建邮局常用 25(STARTTLS)。")
        if baseline_ok:
            diag.suggestions.append(
                f"本机到 {'、'.join(baseline_ok)} 的连接正常，说明不是本机网络封禁 SMTP 出站。")
        else:
            diag.suggestions.append(
                "本机到公共 SMTP 端口也不通，请检查防火墙/代理策略是否允许对外 465、587 端口。")
    elif login_status == "ban":
        diag.conclusion = (
            "本机出口 IP 已被对方邮件安全网关封禁（认证失败次数超限），此状态下任何发送都会失败。"
        )
        diag.suggestions.append(
            "这是 IP 级临时封禁，通常等待一段时间会自动解除；若长时间未恢复，"
            "请联系邮箱管理员解封并确认客户端认证策略。")
        diag.suggestions.append("封禁解除前请勿反复点击发送/测试（每次都会加重计数）。")
        diag.suggestions.extend(
            _auth_advice(rec[0] if rec else server, diag.provider))
    elif probe is not None and not probe.usable and login_status == "skip":
        if probe.status == "silent":
            diag.conclusion = (
                f"{server}:{port} 接受了 TCP 连接、但没有返回任何 SMTP 服务号 —— "
                "这不是一个可用的发信端点（与账号密码无关）。"
            )
            diag.suggestions.append(
                "这类「连得上却零响应」通常是两种情况：① 该域名只是负载均衡/网关前端，"
                "真正的邮件服务器在 MX 记录里；② 连接被本机网络中的安全设备接管了。")
            if rec is not None:
                diag.suggestions.append(
                    f"诊断已通过 MX 探测找到可用端点：{rec[0]}:{rec[1]}（{rec[2]}），"
                    "点「应用推荐配置」即可切换。")
            else:
                diag.suggestions.append(
                    "请向邮箱管理员索取正确的 SMTP 发信主机名与端口。")
        else:
            diag.conclusion = f"{server}:{port} 未提供可用的 SMTP 服务：{probe.summary}"
            diag.suggestions.append("请核对服务器地址与端口，或改用诊断推荐的端点。")
    elif login_status == "auth":
        diag.conclusion = "网络与加密均正常，但账号认证被服务器拒绝（535）：这是密码类型或客户端策略的问题。"
        diag.suggestions.extend(_auth_advice(server, diag.provider))
        if rec is not None:
            diag.suggestions.append(
                f"另外，当前发信端点与实测可用端点不一致，建议同时切换到 "
                f"{rec[0]}:{rec[1]}（{rec[2]}）。")
    elif login_status == "tls":
        diag.conclusion = f"TCP 可连接，但加密协商失败：{login_detail}"
        diag.suggestions.append(
            f"核对加密方式与端口是否配套（465↔SSL、587/25↔STARTTLS、25 可能只能用「不加密」）；"
            f"当前为 {port} + {mode}。")
    elif login_status == "ok":
        diag.ok = True
        diag.conclusion = "链路完整可用：域名解析、TCP 连接、加密协商、账号认证全部成功，配置正确。"
        if invalid:
            diag.ok = False
            diag.conclusion = "SMTP 配置正确，但目标邮箱存在格式错误，发送仍会失败。"
            diag.suggestions.append("修正收件人地址：" + "、".join(invalid))
    elif login_status == "skip":
        pass  # 结论已在其他分支给出
    else:
        diag.conclusion = f"连接阶段异常：{tcp_note or login_detail}"
        diag.suggestions.append("请重试；若持续失败，检查本机防火墙、代理或杀毒软件的网络防护。")

    if invalid and not diag.ok:
        diag.suggestions.append(
            "另外，以下目标邮箱格式不正确，需一并修正：" + "、".join(invalid))

    if not diag.suggestions:
        diag.suggestions.append("未发现明显问题。")


# ------------------------------------------------------------------ 报告 --

def format_report(diag: SmtpDiagnosis) -> str:
    """把诊断结果渲染为可读的中文纯文本报告。"""
    lines = [
        "SMTP 发信链路诊断报告",
        "=" * 52,
        f"诊断时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        "",
    ]
    width = max((len(s.name) for s in diag.steps), default=8)
    for i, step in enumerate(diag.steps, 1):
        lines.append(f"[{i}] {step.name.ljust(width)}  {step.mark}  {step.detail}")
    lines += ["", "-" * 52, "结论：", diag.conclusion or "（未得出结论）"]
    if diag.suggestions:
        lines += ["", "建议："]
        for i, s in enumerate(diag.suggestions, 1):
            head, *rest = s.split("\n")      # 多行建议：首行编号，后续行缩进对齐
            lines.append(f"  {i}. {head}")
            lines.extend(f"     {ln.strip()}" for ln in rest)
    if diag.recommended:
        srv, prt, md = diag.recommended
        lines += ["", f"推荐配置：服务器 {srv}　端口 {prt}　加密 {md}"]
    lines += ["", "-" * 52,
              "说明：本报告只做连通性/服务号/认证验证，不发送任何邮件；",
              "      诊断不会在连接已被网关拒绝时提交账号密码；",
              "      但登录验证仍会提交一次账号密码 —— 部分邮件安全网关会按来源 IP 统计",
              "      认证失败次数并临时封禁，因此密码不确定时请避免反复诊断。"]
    return "\n".join(lines)
