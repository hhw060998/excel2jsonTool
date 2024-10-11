@echo off
REM 设置输入和输出路径
set input_folder=D:\Excel2JsonTool\Excel
set output_project_folder=D:\Excel2JsonTool\ProjectFolder\Json
set output_client_folder=D:\Excel2JsonTool\GameClientFolder\Json
set csfile_output_folder=D:\Excel2JsonTool\ProjectFolder\ConfigData\AutoGeneratedScript
set enum_output_folder=D:\Excel2JsonTool\ProjectFolder\ConfigData\AutoGeneratedEnum

REM 运行python脚本
E:\Python3.10\python ..\ExcelExportTool\main.py %input_folder% %output_project_folder% %output_client_folder% %csfile_output_folder% %enum_output_folder%
pause
