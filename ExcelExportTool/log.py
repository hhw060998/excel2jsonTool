# ANSI escape sequences for colors
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