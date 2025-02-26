# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved

import sys
from export_process import batch_excel_to_json


# 获取命令行参数
root_folder = sys.argv[1]
output_client_folder = sys.argv[2]

# 仅导出到游戏客户端
batch_excel_to_json(root_folder, output_client_folder)
