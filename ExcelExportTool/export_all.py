# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved

from export_process import batch_excel_to_json
import sys

# 获取命令行参数
root_folder = sys.argv[1]
output_project_folder = sys.argv[2]
csfile_output_folder = sys.argv[3]
enum_output_folder = sys.argv[4]

# 导出到工程的配置目录、游戏客户端、脚本目录和枚举目录
batch_excel_to_json(root_folder,None, output_project_folder, csfile_output_folder, enum_output_folder)