# Author: huhongwei 306463233@qq.com
# MIT License
import os
import re
import tempfile
import shutil
from pathlib import Path
from typing import Dict, Optional, List
from log import log_warn, log_info
from naming_config import CS_FILE_SUFFIX

# 输出控制
_DIFF_ONLY = True
_DRY_RUN = False
_created_files: List[str] = []

def set_output_options(diff_only: bool = True, dry_run: bool = False):
    global _DIFF_ONLY, _DRY_RUN
    _DIFF_ONLY = diff_only
    _DRY_RUN = dry_run

def get_create_files():
    return _created_files

def generate_xml_summary(origin_str: str) -> str:
    if origin_str is None:
        origin_str = ""
    lines = origin_str.splitlines()
    if len(lines) <= 1:
        return f"/// <summary> {origin_str} </summary>"
    body = "\n".join(f"/// {l}" for l in lines)
    return f"/// <summary>\n{body}\n/// </summary>"

auto_generated_summary_string = generate_xml_summary("This is auto-generated, don't modify manually")
enum_namespace = "ConfigDataName"
I_CONFIG_RAW_INFO_NAME = "IConfigRawInfo"

def generate_enum_file_from_sheet(sheet, enum_tag, output_folder):
    rows = list(sheet.iter_rows(min_row=2))
    if not rows:
        log_warn(f"枚举表空: {sheet.title}")
        return
    enum_type_name = sheet.title.replace(enum_tag, "")
    enum_names, enum_values, remarks = [], [], []
    for r in rows:
        if len(r) < 2:
            continue
        name = r[0].value
        val = r[1].value
        remark = r[2].value if len(r) > 2 else None
        if name is None or val is None:
            log_warn(f"{sheet.title} 跳过一行（缺 name 或 value）")
            continue
        enum_names.append(name)
        enum_values.append(val)
        remarks.append(remark)
    if not enum_names:
        log_warn(f"{sheet.title} 没有有效枚举项")
        return
    generate_enum_file(enum_type_name, enum_names, enum_values, remarks, enum_namespace, output_folder)

def generate_enum_file(enum_type_name, enum_names, enum_values, remarks, name_space, output_folder):
    file_content = f"namespace {name_space}\n{{\n\t{auto_generated_summary_string}\n\tpublic enum {enum_type_name}\n\t{{\n"
    for i, key in enumerate(enum_names):
        if remarks and remarks[i]:
            file_content += f'\t\t{generate_xml_summary(str(remarks[i]))}\n'
        file_content += f'\t\t{key} = {enum_values[i]},\n\n'
    file_content += "\t}\n}"
    cs_file_path = os.path.join(output_folder, f"{enum_type_name}.cs")
    write_to_file(file_content, cs_file_path)

USING_NAMESPACE_STR = "\n".join([
    "using System.Collections.Generic;",
    "using Newtonsoft.Json;",
    "\n\n",
])
NAMESPACE_WRAPPER_STR = "namespace Data.TableScript\n{{\n{0}\n}}"

def add_indentation(input_str, indent="\t"):
    return "\n".join([indent + line for line in input_str.splitlines()])

def convert_type_to_csharp(type_str):
    type_mappings = {'list': 'List', 'dict': 'Dictionary'}
    for k, v in type_mappings.items():
        type_str = type_str.replace(k, v)
    return type_str.replace('(', '<').replace(')', '>')

def wrap_class_str(class_name, class_content_str, interface_name=""):
    interface_part = f" : {interface_name}" if interface_name else ""
    indented = add_indentation(class_content_str)
    return f"public class {class_name}{interface_part}\n{{\n{indented}\n}}"

def generate_script_file(sheet_name: str,
                         properties_dict: Dict[str, str],
                         property_remarks: Dict[str, str],
                         output_folder: str,
                         need_generate_keys: bool = False,
                         file_suffix: str = CS_FILE_SUFFIX,
                         composite_keys: bool = False,
                         composite_multiplier: int = 46340,
                         composite_key_fields: Optional[Dict[str, str]] = None):
    info_class = f"{auto_generated_summary_string}\n{generate_info_class(sheet_name, properties_dict, property_remarks)}"
    data_class = f"{generate_data_class(sheet_name, need_generate_keys, composite_keys, composite_multiplier)}"
    file_content = f"{info_class}\n\n{data_class}"
    final_content = USING_NAMESPACE_STR + NAMESPACE_WRAPPER_STR.format(add_indentation(file_content))
    cs_file_path = os.path.join(output_folder, f"{sheet_name}{file_suffix}.cs")
    write_to_file(final_content, cs_file_path)

def generate_info_class(class_name, properties_dict, property_remarks):
    property_access_decorator = "{ get; private set; }"
    converted = []
    for k, v in properties_dict.items():
        converted.append(
            f"{generate_xml_summary(property_remarks[k])}\n"
            f"[JsonProperty(\"{k}\")]\n"
            f"public {convert_type_to_csharp(v)} {k} {property_access_decorator}"
        )
    if "id" not in properties_dict:
        log_warn(f"{class_name}Info 缺少 id，已自动补齐")
        converted.append(
            f"{generate_xml_summary('Auto-added to satisfy IConfigRawInfo')}\n"
            "[JsonProperty(\"id\")]\npublic int id { get; private set; }"
        )
    if "name" not in properties_dict:
        log_warn(f"{class_name}Info 缺少 name，已自动补齐")
        converted.append(
            f"{generate_xml_summary('Auto-added to satisfy IConfigRawInfo')}\n"
            "[JsonProperty(\"name\")]\npublic string name { get; private set; }"
        )
    return wrap_class_str(class_name + "Info", "\n\n".join(converted), interface_name=I_CONFIG_RAW_INFO_NAME)

def generate_data_class(sheet_name: str,
                        need_generate_keys: bool,
                        composite_keys: bool,
                        composite_multiplier: int):
    class_name = f"{sheet_name}Config"
    if need_generate_keys:
        base_class = f"ConfigDataWithKey<{sheet_name}Info, {sheet_name}Keys>"
    elif composite_keys:
        base_class = f"ConfigDataWithCompositeId<{sheet_name}Info>"
    else:
        base_class = f"ConfigDataBase<{sheet_name}Info>"
    parts = []
    if composite_keys and not need_generate_keys:
        parts.append(f"protected override int CompositeMultiplier => {composite_multiplier};")
    body = "\n\n".join(parts)
    
    header = (
        f"/// <summary>\n"
        f"/// Config data class for {sheet_name}. Generated by tool.\n"
        f"/// Query methods are provided by ConfigManager; keep this class minimal.\n"
        f"/// </summary>"
    )
    return f"{header}\n{wrap_class_str(class_name, body, interface_name=base_class)}"

def write_to_file(content: str, file_path: str):
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    if _DRY_RUN:
        log_info(f"[DryRun] 生成文件(未写入): {file_path}")
        _created_files.append(str(path.resolve()))
        return
    if _DIFF_ONLY and path.exists():
        try:
            old = path.read_text(encoding="utf-8")
            if old == content:
                log_info(f"文件未变化: {file_path}")
                _created_files.append(str(path.resolve()))
                return
        except Exception:
            pass
    try:
        fd, tmp_name = tempfile.mkstemp(dir=str(path.parent), prefix=".tmp_", suffix=".part")
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            f.write(content)
        shutil.move(tmp_name, path)
        log_info(f"成功生成文件: {file_path}")
        _created_files.append(str(path.resolve()))
    except Exception as e:
        log_warn(f"写入失败 {file_path}: {e}")
