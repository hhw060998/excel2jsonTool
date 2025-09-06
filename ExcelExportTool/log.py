# ANSI escape sequences for colors (保持原风格)
GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
RESET = '\033[0m'


def log_info(msg: str) -> None:
    print(msg)


def log_warn(msg: str) -> None:
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