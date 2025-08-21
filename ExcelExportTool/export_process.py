# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License
# All rights reserved

import sys
import time
import openpyxl
from pathlib import Path
from typing import Optional
from cs_generation import generate_enum_file_from_sheet, get_create_files
from worksheet_data import WorksheetData

# ANSI escape sequences for colors
GREEN = '\033[92m'
RED = '\033[91m'
YELLOW = '\033[93m'
RESET = '\033[0m'


def log_info(msg: str) -> None:
    print(msg)


def log_warn(msg: str) -> None:
    print(f"{YELLOW}{msg}{RESET}")


def log_error(msg: str) -> None:
    print(f"{RED}{msg}{RESET}")


def log_success(msg: str) -> None:
    print(f"{GREEN}{msg}{RESET}")


def process_excel_file(
    excel_path: Path,
    file_sheet_map: dict[str, str],
    output_client_folder: Optional[str],
    output_project_folder: Optional[str],
    csfile_output_folder: Optional[str],
    enum_output_folder: Optional[str],
) -> None:
    """处理单个 Excel 文件"""
    try:
        wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    except Exception as e:
        log_error(f"打开文件 {excel_path} 失败：{e}")
        return

    main_sheet = wb.worksheets[0]

    # 检查 sheet 重名
    if main_sheet.title in file_sheet_map.values():
        dup_file = next(f for f, s in file_sheet_map.items() if s == main_sheet.title)
        raise RuntimeError(
            f"存在与 [{dup_file}] 中相同名称的 sheet [{main_sheet.title}]，无法重复生成"
        )

    main_sheet_data = WorksheetData(main_sheet)

    if output_client_folder:
        main_sheet_data.generate_json(output_client_folder)
    if output_project_folder:
        main_sheet_data.generate_json(output_project_folder)
    if csfile_output_folder:
        main_sheet_data.generate_script(csfile_output_folder)

    if len(wb.worksheets) > 1 and enum_output_folder:
        enum_tag = "Enum-"
        for sheet in wb.worksheets[1:]:
            if sheet.title.startswith(enum_tag):
                generate_enum_file_from_sheet(sheet, enum_tag, enum_output_folder)

    file_sheet_map[excel_path.name] = main_sheet.title


def cleanup_files(output_folders: list[Optional[str]]) -> None:
    """清理不在 create_files 列表中的文件"""
    created_files = get_create_files()
    files_to_delete: list[Path] = []

    for folder in filter(None, output_folders):
        for path in Path(folder).rglob("*"):
            if path.is_file():
                meta_file = path.with_suffix(path.suffix + ".meta")
                if (
                    str(path) not in created_files
                    and not path.suffix == ".meta"
                    and str(meta_file) not in created_files
                ):
                    files_to_delete.append(path)

    if not files_to_delete:
        log_info("没有需要删除的文件")
        return

    log_error("以下文件将被删除：")
    for f in files_to_delete:
        log_error(f" - {f}")

    confirm = input("确认删除这些文件吗？(y/n): ").strip().lower()
    if confirm == "y":
        for f in files_to_delete:
            f.unlink(missing_ok=True)
            log_error(f"删除文件 {f}")
        log_info(f"删除了 {len(files_to_delete)} 个文件")
    else:
        log_warn("用户取消了文件删除操作")


def batch_excel_to_json(
    source_folder: str,
    output_client_folder: Optional[str] = None,
    output_project_folder: Optional[str] = None,
    csfile_output_folder: Optional[str] = None,
    enum_output_folder: Optional[str] = None,
) -> None:
    """
    Converts multiple Excel files in a source folder to JSON format and optionally generates additional files.
    """
    start_time = time.time()
    log_info("开始导表……")
    log_info(f"Excel目录: {source_folder}")

    skip_count = 0
    file_count = 0
    file_sheet_map: dict[str, str] = {}

    for excel_path in Path(source_folder).rglob("*.xlsx"):
        if not excel_path.name[0].isupper():
            log_warn(f"文件 {excel_path} 首字母不是大写字母，将不会导出数据")
            skip_count += 1
            continue

        log_info("——————————————————————————————————————————————————")
        log_info(f"即将开始处理文件 {GREEN}{excel_path}{RESET}")

        try:
            process_excel_file(
                excel_path,
                file_sheet_map,
                output_client_folder,
                output_project_folder,
                csfile_output_folder,
                enum_output_folder,
            )
            file_count += 1
        except RuntimeError as e:
            log_error(str(e))
            sys.exit(1)

    log_info("——————————————————————————————————————————————————")
    log_info("检查是否存在需要清理的文件……")

    cleanup_files(
        [output_client_folder, output_project_folder, csfile_output_folder, enum_output_folder]
    )

    elapsed_time = time.time() - start_time
    log_info("——————————————————————————————————————————————————")
    log_success(
        f"导表结束，跳过了 {YELLOW}{skip_count}{GREEN} 个Excel文件，"
        f"成功处理了 {file_count} 个Excel文件，总耗时 {elapsed_time:.2f} 秒"
    )
