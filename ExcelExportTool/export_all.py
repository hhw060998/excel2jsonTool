# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved

from .export_process import batch_excel_to_json

import sys
from .worksheet_data import user_confirm

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
force_no_assets = False
if "--no-diff" in sys.argv:
    diff_only = False
if "--dry-run" in sys.argv:
    dry_run = True
if "--force-no-assets" in sys.argv:
    force_no_assets = True

# 检查所有输出路径是否包含 Assets
def _contains_assets(*paths):
    for p in paths:
        if p and "assets" in str(p).replace("\\", "/").lower():
            return True
    return False

if not force_no_assets and not _contains_assets(output_project_folder, csfile_output_folder, enum_output_folder):
    msg = "警告：导出路径不包含 'Assets'，这通常不是 Unity 项目目录。是否继续导出？(y/n): "
    if not user_confirm(msg, title="导出路径警告"):
        print("已取消导出。")
        sys.exit(1)

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