# -*- coding: utf-8 -*-
"""Foxmail 集成模块：自动定位 Foxmail 并用它完成投递。

为什么需要它
------------
单位邮件系统启用了安全网关（Coremail + CACTER）后，第三方客户端直连 SMTP 会被
策略拒绝（`535 reject by CACTER gateway`）。此时唯一可行的投递方式是**让 Foxmail
用它自己已经配好、能正常发信的账号把邮件发出去**。

本模块负责把这件事做成"点一下就行"：

1. `detect_foxmail()`  —— 自动找到 Foxmail 的安装位置（注册表 + 常见路径 + PATH）；
2. `build_launch_argv()` —— 纯函数，构造唤起 Foxmail 的命令行（便于测试）；
3. `open_foxmail_compose()` —— 真正唤起 Foxmail（打开导出的 .eml，或预填写信窗口）。

两种唤起方式
------------
- `"eml"`    ：`Foxmail.exe "C:\\...\\权限到期提醒_xxx.eml"`
  让 Foxmail 打开我们导出的那封邮件（收件人 / 主题 / 正文已经写好），
  在 Foxmail 里点「转发 / 再次编辑 → 发送」即可。**最可靠**，因为它只是
  "用某个程序打开某个文件"，不依赖任何私有命令行参数。
- `"mailto"` ：`Foxmail.exe "mailto:a@b.com?subject=...&body=..."`
  让 Foxmail 直接弹出**新建邮件窗口**并预填内容。一步到位，但依赖 Foxmail
  对 mailto 链接作为启动参数的支持。
- `"auto"`   ：有 .eml 时用 `eml`，否则退到 `mailto`（默认）。

设计约定
--------
- 只负责"唤起"，**不负责点击发送**——最后一步始终由用户在 Foxmail 里完成；
- 不写任何注册表、不修改 Foxmail 的任何配置/数据文件；
- 找不到 Foxmail 时**不抛异常**，`detect_foxmail()` 返回空串，由界面引导用户
  手动指定路径或改用其他投递方式。
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from .outbox import build_mailto

logger = logging.getLogger(__name__)

# 唤起方式（与界面单选一致）
FOXMAIL_MODES = ("auto", "eml", "mailto")

# Foxmail 主程序名（大小写不敏感比较）
EXE_NAME = "foxmail.exe"


class FoxmailError(Exception):
    """唤起 Foxmail 失败（路径无效、程序不存在、启动报错等）。"""


# --------------------------------------------------------------- 注册表探测 ---

def _reg_value(hive: int, subkey: str, value: str) -> str:
    """读单个注册表字符串值，任何失败都返回空串（探测不应打断主流程）。"""
    if not sys.platform.startswith("win"):
        return ""
    try:
        import winreg
        with winreg.OpenKey(hive, subkey) as key:
            data = winreg.QueryValueEx(key, value)[0]
        return str(data or "")
    except (ImportError, OSError):
        return ""


def _reg_subkeys(hive: int, subkey: str) -> List[str]:
    """枚举子键名，失败返回空列表。"""
    if not sys.platform.startswith("win"):
        return []
    names: List[str] = []
    try:
        import winreg
        with winreg.OpenKey(hive, subkey) as key:
            i = 0
            while True:
                try:
                    names.append(winreg.EnumKey(key, i))
                except OSError:
                    break
                i += 1
    except (ImportError, OSError):
        pass
    return names


def parse_command_path(cmd: str) -> str:
    """从注册表的 shell\\open\\command 字符串里提取可执行文件路径。

    典型形式：
      `"C:\\Program Files (x86)\\Foxmail\\Foxmail.exe" "%1"`
      `D:\\Foxmail\\Foxmail.exe %1`
      `"C:\\Foxmail\\Foxmail.exe",0`（带资源索引后缀）
    还会展开 `%ProgramFiles%` 这类环境变量。
    """
    s = (cmd or "").strip()
    if not s:
        return ""
    path = ""
    if s.startswith('"'):
        end = s.find('"', 1)
        if end > 1:
            path = s[1:end]
    else:
        # 无引号：截到 .exe 为止（后面可能跟着参数）
        m = re.match(r"(?i)^(.*?\.exe)", s)
        path = m.group(1) if m else s.split(" ")[0]
    path = path.strip().strip('"').rstrip(",")
    # 去掉末尾的 ,0 / ,1 资源索引
    path = re.sub(r",\d+$", "", path).strip()
    return os.path.expandvars(path)


def is_valid_foxmail_exe(path: str) -> bool:
    """判断给定路径是否确实是一个存在的 Foxmail.exe。"""
    if not path:
        return False
    try:
        p = Path(path)
    except (OSError, ValueError):
        return False
    if p.name.lower() != EXE_NAME:
        return False
    try:
        return p.is_file()
    except OSError:
        return False


def _uninstall_candidates() -> List[str]:
    """从"添加或删除程序"的注册表项里找 Foxmail 的安装位置。"""
    out: List[str] = []
    if not sys.platform.startswith("win"):
        return out
    import winreg
    roots = (
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_LOCAL_MACHINE,
         r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
        (winreg.HKEY_CURRENT_USER,
         r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
    )
    for hive, root in roots:
        for name in _reg_subkeys(hive, root):
            sub = f"{root}\\{name}"
            display = _reg_value(hive, sub, "DisplayName")
            if "foxmail" not in display.lower():
                continue
            # 依次尝试 DisplayIcon / InstallLocation
            icon = parse_command_path(_reg_value(hive, sub, "DisplayIcon"))
            if icon:
                out.append(icon)
            loc = _reg_value(hive, sub, "InstallLocation")
            if loc:
                out.append(str(Path(os.path.expandvars(loc)) / "Foxmail.exe"))
    return out


def _registry_candidates() -> List[str]:
    """从三处注册表位置收集候选 exe 路径。"""
    out: List[str] = []
    if not sys.platform.startswith("win"):
        return out
    import winreg
    hives = (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE)
    # ① 作为「邮件客户端」注册过（shell\open\command）
    for hive in hives:
        for root in (r"SOFTWARE\Clients\Mail\Foxmail",
                     r"SOFTWARE\WOW6432Node\Clients\Mail\Foxmail"):
            out.append(parse_command_path(
                _reg_value(hive, f"{root}\\shell\\open\\command", "")))
    # ② App Paths（Windows 用它按名字解析可执行文件）
    for hive in hives:
        for root in (r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\Foxmail.exe",
                     r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion"
                     r"\App Paths\Foxmail.exe"):
            out.append(_reg_value(hive, root, ""))
    # ③ Foxmail 自家的键（Aerofox 是 Foxmail 的公司名）
    for hive in hives:
        for root in (r"SOFTWARE\Aerofox", r"SOFTWARE\Foxmail"):
            for sub in _reg_subkeys(hive, root):
                out.append(_reg_value(hive, f"{root}\\{sub}", "InstallPath"))
                out.append(_reg_value(hive, f"{root}\\{sub}", "Path"))
    return [c for c in out if c]


def common_install_paths() -> List[str]:
    """常见安装位置（注册表没登记时的兜底）。"""
    dirs = [
        r"C:\Program Files (x86)\Foxmail",
        r"C:\Program Files\Foxmail",
        r"C:\Foxmail",
        r"D:\Foxmail",
        r"D:\Program Files\Foxmail",
        r"D:\Program Files (x86)\Foxmail",
        r"E:\Foxmail",
        r"E:\Program Files\Foxmail",
        r"C:\Users\Public\Foxmail",
    ]
    local = os.environ.get("LOCALAPPDATA", "")
    if local:
        dirs.append(str(Path(local) / "Foxmail"))
        dirs.append(str(Path(local) / "Programs" / "Foxmail"))
    # Foxmail 安装器默认装到 "C:\Foxmail 7.2" 这类带版本号的目录，
    # 固定列表覆盖不全，因此对每个盘符做一次**浅层**目录名前缀扫描（成本极低）。
    for drive in ("C:\\", "D:\\", "E:\\"):
        root = Path(drive)
        try:
            if not root.is_dir():
                continue
            for child in root.iterdir():
                try:
                    if child.is_dir() and child.name.lower().startswith("foxmail"):
                        dirs.append(str(child))
                except OSError:
                    continue
        except OSError:
            continue
    return [str(Path(d) / "Foxmail.exe") for d in dirs]


def foxmail_data_dir() -> str:
    """返回 Foxmail 的用户数据目录（存在说明这台机器用过 Foxmail）。

    用于给出更准确的提示：比如"检测到 Foxmail 邮件数据，但没找到主程序"，
    往往意味着程序装在非默认位置或被卸载了。
    """
    cands = [Path(os.environ.get("APPDATA", "")) / "Foxmail",
             Path(os.environ.get("LOCALAPPDATA", "")) / "Foxmail",
             Path.home() / "Documents" / "Foxmail"]
    for c in cands:
        try:
            if str(c) and c.is_dir():
                return str(c)
        except OSError:
            continue
    return ""


def detect_foxmail(extra: Optional[Sequence[str]] = None) -> str:
    """自动探测 Foxmail 主程序路径；找不到返回空串。

    Args:
        extra: 额外的候选路径（例如用户上次手工指定的路径），优先检查。

    探测顺序（可靠度从高到低）：
        用户指定 → 注册表（Clients\\Mail / App Paths / Aerofox）
        → 卸载信息里的安装位置 → 常见安装目录 → PATH。
    """
    seen = set()
    ordered: List[str] = []
    for path in list(extra or []) + _registry_candidates() + \
            _uninstall_candidates() + common_install_paths():
        if not path:
            continue
        key = path.lower()
        if key in seen:
            continue
        seen.add(key)
        ordered.append(path)

    for cand in ordered:
        if is_valid_foxmail_exe(cand):
            logger.info("已检测到 Foxmail：%s", cand)
            return str(Path(cand))

    # 最后试 PATH（少见于 Foxmail，但成本极低）
    try:
        found = shutil.which("Foxmail.exe") or shutil.which("foxmail")
    except (OSError, TypeError):
        found = None
    if found and is_valid_foxmail_exe(found):
        logger.info("已在 PATH 中找到 Foxmail：%s", found)
        return found

    logger.info("未检测到 Foxmail（已检查注册表、卸载信息、常见目录与 PATH）")
    return ""


def describe_foxmail(path: str = "", detected: Optional[str] = None) -> str:
    """生成一句给人看的 Foxmail 状态描述（供界面与提示使用）。

    Args:
        path:     界面上填写的路径（可能无效或为空）。
        detected: 调用方已经探测到的结果。传 `None`（默认）表示"还没探测过"，
                  本函数会自行探测一次；传空串表示"已探测过，确实没有"，
                  此时不再重复扫描（避免同一句话里扫两遍注册表）。

    对无效路径会**自行再探测一次**，避免出现"明明装了 Foxmail 却报未检测到"的误导。
    """
    if is_valid_foxmail_exe(path):
        return f"已集成 Foxmail：{path}"

    if detected is None:
        detected = detect_foxmail(extra=[path] if path else None)
    if detected:
        note = f"已自动检测到 Foxmail：{detected}"
        if path:
            note += f"（界面填写的「{path}」无效，已改用检测结果）"
        return note

    data = foxmail_data_dir()
    if data:
        return ("未找到 Foxmail 主程序，但检测到 Foxmail 邮件数据目录"
                f"（{data}）——可能装在非默认位置，请点「浏览」指定 Foxmail.exe，"
                "或点「自动检测」重试")
    return ("未检测到 Foxmail（注册表、卸载信息、常见安装目录、PATH 都没有）。"
            "若尚未安装，请先安装 Foxmail 并在其中配好单位邮箱账号，再回来点「自动检测」")


# ------------------------------------------------------------- 唤起 Foxmail ---

def build_launch_argv(exe: str, mode: str, eml_path: str,
                      recipients: Sequence[str], subject: str,
                      plain_body: str) -> List[str]:
    """构造唤起 Foxmail 的命令行参数（纯函数，不产生副作用，便于测试）。

    Args:
        exe:        Foxmail.exe 路径。
        mode:       `"auto"` / `"eml"` / `"mailto"`。
        eml_path:   已导出的 .eml 路径（mode 为 eml/auto 时使用）。
        recipients: 收件人列表（mode 为 mailto 时使用）。
        subject:    主题。
        plain_body: 纯文本正文。

    Returns:
        交给 `subprocess.Popen` 的参数列表。

    Raises:
        FoxmailError: 参数不满足所选方式的要求（如选 eml 但文件不存在）。
    """
    if mode not in FOXMAIL_MODES:
        raise FoxmailError(f"未知的唤起方式：{mode!r}（可选 {'/'.join(FOXMAIL_MODES)}）")
    if not exe:
        raise FoxmailError("未指定 Foxmail 主程序路径")

    use = mode
    if use == "auto":
        # 有导出的 .eml 就优先用它（不依赖私有命令行参数，最可靠）
        use = "eml" if (eml_path and Path(eml_path).is_file()) else "mailto"

    if use == "eml":
        if not eml_path:
            raise FoxmailError("选择「打开导出邮件」方式时，必须先生成 .eml 文件")
        if not Path(eml_path).is_file():
            raise FoxmailError(f"找不到要打开的邮件文件：{eml_path}")
        return [exe, str(Path(eml_path))]

    # mailto：直接把 mailto: 链接作为启动参数交给 Foxmail
    url = build_mailto(list(recipients), subject, plain_body)
    return [exe, url]


def open_foxmail_compose(exe: str, mode: str, eml_path: str,
                         recipients: Sequence[str], subject: str,
                         plain_body: str) -> List[str]:
    """唤起 Foxmail 完成投递前的最后一步（打开邮件或弹出写信窗口）。

    Returns:
        实际使用的命令行参数（调用方可写进日志/提示，便于排查）。

    Raises:
        FoxmailError: 路径无效、方式不满足、或启动过程报错。

    注意：本函数**只负责唤起**。邮件是否真的发出去，取决于用户在 Foxmail 里
    点「发送」——这一点无法由本工具代劳，因此调用方必须显式向用户确认。
    """
    if not is_valid_foxmail_exe(exe):
        raise FoxmailError(f"Foxmail 主程序路径无效或不存在：{exe or '(空)'}")
    argv = build_launch_argv(exe, mode, eml_path, recipients, subject, plain_body)
    try:
        subprocess.Popen(argv, close_fds=True)  # noqa: S603 —— argv 已由本函数构造
    except OSError as exc:
        raise FoxmailError(f"启动 Foxmail 失败：{exc}") from exc
    logger.info("已唤起 Foxmail（%s）：%s", mode, argv[0])
    return argv


def open_default_apps_settings() -> bool:
    """打开 Windows「默认应用」设置页，便于把 Foxmail 设为默认邮件程序。

    Returns:
        True 表示已请求系统打开设置页；False 表示当前平台不支持。

    刻意**不直接改注册表**：把某个程序设为默认邮件客户端属于用户偏好，
    由 Windows 的设置在用户确认下完成更稳妥，也避免写坏关联。
    """
    if not sys.platform.startswith("win"):
        return False
    try:
        os.startfile("ms-settings:defaultapps")  # type: ignore[attr-defined]  # noqa: S606
        logger.info("已打开系统「默认应用」设置页")
        return True
    except OSError as exc:
        logger.warning("打开默认应用设置页失败：%s", exc)
        return False


def summary_line(exe: str, mode: str) -> str:
    """一行摘要，供日志与界面显示。"""
    mode_label = {"auto": "自动", "eml": "打开导出邮件(.eml)",
                  "mailto": "新建写信窗口(mailto)"}.get(mode, mode)
    if is_valid_foxmail_exe(exe):
        return f"Foxmail：已集成（{exe}）｜唤起方式：{mode_label}"
    return "Foxmail：未集成（未检测到主程序）"


def resolve(exe_hint: str = "") -> Tuple[str, str]:
    """综合"界面已填路径"与"自动探测"，返回 (可用 exe 路径, 探测说明)。

    界面填了有效路径就尊重它；否则自动探测，并把过程说明返回给界面显示。
    """
    if is_valid_foxmail_exe(exe_hint):
        return str(Path(exe_hint)), f"已集成 Foxmail：{exe_hint}"
    detected = detect_foxmail(extra=[exe_hint] if exe_hint else None)
    if detected:
        note = f"已自动检测到 Foxmail：{detected}"
        if exe_hint:
            note += f"（原先填写的「{exe_hint}」无效，已改用检测结果）"
        return detected, note
    # detected="" 告诉 describe_foxmail「已经扫过了」，避免重复扫描
    return "", describe_foxmail(exe_hint, detected="")
