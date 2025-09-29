# Author: huhongwei 306463233@qq.com
# MIT License
import time
import openpyxl
from pathlib import Path
from typing import Optional
import sys
import os
import shutil
import traceback
from worksheet_data import WorksheetData
from cs_generation import generate_enum_file_from_sheet, get_created_files, set_output_options
from log import log_info, log_warn, log_error, log_success, log_sep, green_filename, log_warn_summary
from exceptions import SheetNameConflictError, ExportError
from exceptions import InvalidFieldNameError
from exceptions import WriteFileError
from exceptions import DuplicateFieldError, HeaderFormatError, UnknownCustomTypeError
REPORT = None  # 报表文件输出已移除

def process_excel_file(
    excel_path: Path,
    file_sheet_map: dict[str, str],
    output_client_folder: Optional[str],
    output_project_folder: Optional[str],
    csfile_output_folder: Optional[str],
    enum_output_folder: Optional[str],
) -> Optional[WorksheetData]:
    try:
        wb = openpyxl.load_workbook(str(excel_path), data_only=True)
    except Exception as e:
        log_error(f"打开失败: {green_filename(excel_path.name)} -> {e}")
        return None
    main_sheet = wb.worksheets[0]
    if main_sheet.title in file_sheet_map.values():
        dup = next(f for f, s in file_sheet_map.items() if s == main_sheet.title)
        raise SheetNameConflictError(main_sheet.title, dup, excel_path.name)

    log_sep(f"开始 {green_filename(excel_path.name)}")
    main_sheet_data = WorksheetData(main_sheet)
    # 记录来源 Excel 文件名，供日志使用
    try:
        setattr(main_sheet_data, "source_file", excel_path.name)
    except Exception:
        pass

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
    log_info(f"完成 {excel_path.name} \n")
    return main_sheet_data

def cleanup_files(output_folders):
    created = set(get_created_files())
    from pathlib import Path
    stale = []
    for folder in filter(None, output_folders):
        p = Path(folder)
        if not p.exists():
            continue
        for f in p.rglob("*"):
            if f.is_file() and f.suffix != ".meta":
                if str(f.resolve()) not in created and str(f.with_suffix(f.suffix + ".meta").resolve()) not in created:
                    stale.append(f)
    if not stale:
        log_info("没有需要删除的文件")
        return
    # 临时切换 warning 为立即输出
    log_warn("以下文件未在本次生成中出现：", immediate=True)
    for f in stale:
        log_warn(f" - {f}", immediate=True)
    from worksheet_data import user_confirm
    msg = "是否删除这些文件?(y/n): "
    if user_confirm(msg, title="文件删除确认"):
        for f in stale:
            try:
                f.unlink(missing_ok=True)
                log_info(f"删除: {f}")
            except Exception as e:
                log_error(f"删除失败 {f}: {e}")
    else:
        log_warn("已取消清理")

def batch_excel_to_json(
    source_folder: str,
    output_client_folder: Optional[str] = None,
    output_project_folder: Optional[str] = None,
    csfile_output_folder: Optional[str] = None,
    enum_output_folder: Optional[str] = None,
    diff_only: bool = True,
    dry_run: bool = False,
    auto_cleanup: bool = True,
) -> None:
    start = time.time()
    log_sep("开始导表")
    log_info(f"Excel目录: {source_folder}")
    set_output_options(diff_only=diff_only, dry_run=dry_run)

    # 输出目录可写性与磁盘空间检查
    def _check_output_dir(folder: Optional[str]) -> None:
        if not folder:
            return
        try:
            os.makedirs(folder, exist_ok=True)
        except Exception as e:
            raise ExportError(f"无法创建输出目录: {folder} -> {e}")
        # 检查是否可写
        if not os.access(folder, os.W_OK):
            raise ExportError(f"输出目录不可写: {folder}")
        # 检查剩余空间（简单策略：至少 10MB 可用）
        try:
            stat = os.statvfs(folder)
            free = stat.f_bavail * stat.f_frsize
            if free < 10 * 1024 * 1024:
                raise ExportError(f"输出目录磁盘空间不足 (<10MB): {folder}")
        except AttributeError:
            # Windows 不支持 statvfs，尝试使用 shutil.disk_usage
            try:
                du = shutil.disk_usage(folder)
                if du.free < 10 * 1024 * 1024:
                    raise ExportError(f"输出目录磁盘空间不足 (<10MB): {folder}")
            except Exception:
                pass

    import shutil
    _check_output_dir(output_client_folder)
    _check_output_dir(output_project_folder)
    _check_output_dir(csfile_output_folder)
    _check_output_dir(enum_output_folder)

    skip = 0
    ok = 0
    file_sheet_map: dict[str, str] = {}
    # 反查 map: sheet 名 -> excel 文件名（用于日志显示目标 Excel 文件）
    sheet_to_file_map: dict[str, str] = {}
    excel_files = list(Path(source_folder).rglob("*.xlsx"))

    if not excel_files:
        log_warn("未找到 .xlsx 文件")

    sheets: list[WorksheetData] = []
    aborted = False
    for excel_path in excel_files:
        if not excel_path.name[0].isupper():
            log_warn(f"跳过(首字母非大写): {green_filename(excel_path.name)}")
            skip += 1
            continue
        try:
            ws = process_excel_file(
                excel_path,
                file_sheet_map,
                output_client_folder,
                output_project_folder,
                csfile_output_folder,
                enum_output_folder,
            )
            if ws is not None:
                sheets.append(ws)
                # 记录 sheet -> 文件名
                try:
                    sheet_to_file_map[ws.name] = excel_path.name
                except Exception:
                    pass
            ok += 1
        except Exception as e:
            # 所有异常视为致命：打印红色错误、堆栈信息，并给出建议后立即退出
            tb = traceback.format_exc()
            log_error(f"{excel_path.name} 失败: {e}\n{tb}")
            sys.exit(1)

    # 统一引用检查（导出后）
    if sheets and not aborted:
        search_dirs = [output_client_folder, output_project_folder]
        # 空行分隔阶段，并打印一次阶段标题
        log_info("")
        log_info("————开始引用检查————")
        for ws in sheets:
            try:
                ws.run_reference_checks(search_dirs, sheet_to_file_map)
            except Exception as e:
                log_error(f"[{ws.name}] 引用检查失败: {e}")

    # 打印每表错误/警告统计（若实现了内部统计则输出；当前由 worksheet 在控制台输出具体错误）

    if auto_cleanup and not aborted:
        log_sep("清理阶段")
        cleanup_files([output_client_folder, output_project_folder, csfile_output_folder, enum_output_folder])

    elapsed = time.time() - start
    log_sep("结束")
    # 统计本次生成的 JSON 文件总体积（仅统计已实际写入的文件）
    try:
        created_files = set(get_created_files())
        total_json_bytes = 0
        for p in created_files:
            try:
                if p.lower().endswith('.json') and os.path.isfile(p):
                    total_json_bytes += os.path.getsize(p)
            except Exception:
                # 忽略单个文件统计失败
                pass

        def _human_bytes(n: int) -> str:
            # 简单的人类可读格式
            if n < 1024:
                return f"{n} B"
            if n < 1024 * 1024:
                return f"{n/1024:.1f} KB"
            return f"{n/1024/1024:.2f} MB"

        total_json_str = _human_bytes(total_json_bytes)
    except Exception:
        total_json_bytes = 0
        total_json_str = "N/A"

    # 在打印最终结果前统一输出所有警告，便于快速查看
    try:
        log_warn_summary("以下为本次运行收集到的所有警告：")
    except Exception:
        pass
    if aborted:
        log_error(f"导表已中止: 字段命名不合法，已停止后续处理。成功 {ok}，跳过 {skip}，总耗时 {elapsed:.2f}s，总生成 JSON 大小: {total_json_str}")
    else:
        log_success(f"成功 {ok}，跳过 {skip}，总耗时 {elapsed:.2f}s，总生成 JSON 大小: {total_json_str}. diff_only:{diff_only}, dry_run:{dry_run}")
