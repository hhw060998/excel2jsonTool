# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved

import os
import sys
import time
import openpyxl
from typing import Optional
from pathlib import Path
from cs_generation import generate_enum_file_from_sheet, get_create_files
from worksheet_data import WorksheetData

# ANSI escape sequences for colors
GREEN = '\033[92m'
RED = '\033[91m'
RESET = '\033[0m'
YELLOW = '\033[93m'

def print_red(text):
    print(f"{RED}{text}{RESET}")

def print_green(text):
    print(f"{GREEN}{text}{RESET}")

def print_yellow(text):
    print(f"{YELLOW}{text}{RESET}")

def batch_excel_to_json(
    source_folder: str,
    output_client_folder: str,
    output_project_folder: Optional[str] = None,
    csfile_output_folder: Optional[str] = None,
    enum_output_folder: Optional[str] = None
    ) -> None:
    
    """
    Converts multiple Excel files in a source folder to JSON format and optionally generates additional files.
    Args:
        source_folder (str): The directory containing the Excel files to be processed.
        output_client_folder (str): The directory where the JSON files for the client will be saved.
        output_project_folder (str, optional): The directory where the JSON files for the project will be saved. Defaults to None.
        csfile_output_folder (str, optional): The directory where the C# script files will be saved. Defaults to None.
        enum_output_folder (str, optional): The directory where the enum files will be saved. Defaults to None.
    Raises:
        SystemExit: If a worksheet with the same name has already been exported from another file.
    Prints:
        Various status messages indicating the progress of the conversion process, including the number of files processed and the time taken.
    """
    start_time = time.time()
    print(f"开始导表……")
    print(f"Excel目录:{source_folder}")

    skip_count = 0
    file_count = 0
    file_sheet_map = {}
    for folder_name, subfolders, filenames in os.walk(source_folder):
        for filename in filenames:
            if filename.endswith('.xlsx'):
                
                if filename[0].isupper() == False:
                    print(f"{YELLOW}文件{folder_name}\\{GREEN}{filename}{YELLOW}首字母不是大写字母，将不会导出数据{RESET}")
                    skip_count += 1
                    continue
                    
                excel_file = os.path.join(folder_name, filename)
                print("——————————————————————————————————————————————————")
                print(f"即将开始处理文件{folder_name}\\{GREEN}{filename}{RESET}")
                
                try:
                    wb = openpyxl.load_workbook(str(excel_file), data_only=True)
                
                except Exception as e:
                    print_red(f"打开文件{excel_file}失败：{e}")
                    continue

                # 如果worksheet的名字已经被导出过了（在file_sheet_map中），则中断导表并打印错误信息：与xx文件名的sheet重名
                if wb.worksheets[0].title in file_sheet_map.values():
                    # 获取已经导出且sheet名重复的文件名
                    for key, value in file_sheet_map.items():
                        if value == wb.worksheets[0].title:
                            print_red(f"存在与[{key}]中相同名称的sheet[{wb.worksheets[0].title}]，无法重复生成，退出导表")
                            sys.exit()

                main_sheet_data = WorksheetData(wb.worksheets[0])
                
                if output_client_folder is not None:
                    main_sheet_data.generate_json(output_client_folder)
                
                if output_project_folder is not None:
                    main_sheet_data.generate_json(output_project_folder)
                    
                if csfile_output_folder is not None:
                    main_sheet_data.generate_script(csfile_output_folder)

                if len(wb.worksheets) > 1 and enum_output_folder is not None:
                    enum_tag = "Enum-"
                    for sheet in wb.worksheets[1:]:
                        if sheet.title.startswith(enum_tag):
                            generate_enum_file_from_sheet(sheet, enum_tag, enum_output_folder)

                file_count += 1
                file_sheet_map[filename] = wb.worksheets[0].title

    print("——————————————————————————————————————————————————")

    print(f"准备清理目录其他非生成文件……")
    # 遍历output_project_folder、output_client_folder、csfile_output_folder、enum_output_folder中的所有文件
    # 如果文件不存在get_create_files中，则删除
    created_files = get_create_files()
    delete_count = 0
    for folder in [output_client_folder, output_project_folder, csfile_output_folder, enum_output_folder]:
        if folder is not None:
            for folder_name, subfolders, filenames in os.walk(folder):
                for filename in filenames:
                    file_path = os.path.abspath(os.path.join(folder_name, filename))  # 使用绝对路径
                    meta_file_path = file_path + '.meta'
                    if file_path not in created_files and not file_path.endswith(
                            '.meta') and meta_file_path not in created_files:
                        os.remove(file_path)
                        print_red(f"删除文件{file_path}")
                        delete_count += 1

    if delete_count == 0:
        print("没有需要删除的文件")
    else:
        print(f"删除了{delete_count}个文件")

    end_time = time.time()
    elapsed_time = end_time - start_time
    print("——————————————————————————————————————————————————")
    print(f"{GREEN}导表结束，跳过了{YELLOW}{skip_count}{GREEN}个Excel文件，成功处理了{GREEN}{file_count}{GREEN}个Excel文件，总耗时{elapsed_time:.2f}秒{RESET}")