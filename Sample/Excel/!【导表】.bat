@echo off
REM 设置输入和输出路径
set input_folder=D:\Sample\Excel
set output_project_folder=D:\Sample\ProjectFolder\Json
set output_client_folder=D:\Sample\GameClientFolder\Json
set csfile_output_folder=D:\Sample\ProjectFolder\AutoGeneratedScript
set enum_output_folder=D:\Sample\ProjectFolder\AutoGeneratedEnum

REM 运行python脚本
E:\Python3.10\python ..\ExcelExportTool\main.py %input_folder% %output_project_folder% %output_client_folder% %csfile_output_folder% %enum_output_folder%
pause
