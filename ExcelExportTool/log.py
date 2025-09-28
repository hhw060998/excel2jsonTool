# ANSI escape sequences for colors (保持原风格)
GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
RESET = '\033[0m'


_warnings: list[str] = []


LOG_WARN_IMMEDIATE = False  # 是否在 warn 发生时立即打印（默认 False -> 只在结尾汇总时打印）
def log_info(msg: str) -> None:
    print(msg)


def log_warn(msg: str) -> None:
    # 保存到警告缓存，便于在最后统一汇总输出
    try:
        _warnings.append(str(msg))
    except Exception:
        pass
    # 只有在配置为立即打印时才输出到控制台，默认行为仅缓存
    if LOG_WARN_IMMEDIATE:
        print(f"{YELLOW}[Warn] {msg}{RESET}")


def log_error(msg: str) -> None:
    print(f"{RED}{msg}{RESET}")


def log_success(msg: str) -> None:
    print(f"{GREEN}{msg}{RESET}")


def log_sep(title: str = ""):
    line = "─" * 10
    if title:
        log_info(f"{line} {title} {line}")
    else:
        log_info(line * 2)


# 新增：文件名高亮（Excel 文件名统一使用绿色）
def green_filename(name: str) -> str:
    return f"{GREEN}{name}{RESET}"


def log_warn_summary(header: str = None) -> None:
    """
    将本次运行期间收集到的所有 warn 消息一次性输出并清空缓存。
    保持对现有日志的非破坏性；若无警告则不输出任何内容。
    """
    if not _warnings:
        return
    if header:
        log_info(header)
    else:
        log_info("----- Warnings -----")
    for w in _warnings:
        # 已经带有 [Warn] 前缀的条目，直接打印
        print(f"{YELLOW}[Warn] {w}{RESET}")
    # 清空缓存，避免重复打印
    _warnings.clear()