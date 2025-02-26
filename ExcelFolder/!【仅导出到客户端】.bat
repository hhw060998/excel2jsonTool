@echo off
REM 设置输入和输出路径
set input_folder=D:\Excel2JsonTool\ExcelFolder
set output_client_folder=D:\Excel2JsonTool\GameClientFolder\Json

REM 运行python脚本
"C:\Program Files\Python 3.13.0\python.exe" ..\ExcelExportTool\export_game_client.py %input_folder% %output_client_folder%
pause