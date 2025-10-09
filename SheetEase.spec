# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['ExcelExportTool', 'ExcelExportTool.export_process', 'ExcelExportTool.worksheet_data', 'ExcelExportTool.cs_generation', 'ExcelExportTool.data_processing', 'ExcelExportTool.excel_processing', 'ExcelExportTool.type_utils', 'ExcelExportTool.naming_config', 'ExcelExportTool.naming_utils', 'ExcelExportTool.log', 'ExcelExportTool.exceptions']
hiddenimports += collect_submodules('ExcelExportTool')


a = Analysis(
    ['ExcelExportTool\\app_main.py'],
    pathex=[],
    binaries=[],
    datas=[('ProjectFolder', 'ProjectFolder'), ('ExcelExportTool', 'ExcelExportTool')],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='SheetEase',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
