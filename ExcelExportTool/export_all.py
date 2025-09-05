# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved

from export_process import batch_excel_to_json
import sys

# 兼容原参数顺序:
# root_folder project_output csfile_output enum_output
if len(sys.argv) < 5:
    print("用法: python export_all.py <Excel根目录> <工程配置输出> <C#脚本输出> <枚举输出> [--no-diff] [--dry-run]")
    sys.exit(1)

root_folder = sys.argv[1]
output_project_folder = sys.argv[2]
csfile_output_folder = sys.argv[3]
enum_output_folder = sys.argv[4]

diff_only = True
dry_run = False
if "--no-diff" in sys.argv:
    diff_only = False
if "--dry-run" in sys.argv:
    dry_run = True

batch_excel_to_json(
    root_folder,
    output_client_folder=None,
    output_project_folder=output_project_folder,
    csfile_output_folder=csfile_output_folder,
    enum_output_folder=enum_output_folder,
    diff_only=diff_only,
    dry_run=dry_run,
    auto_cleanup=True,
)