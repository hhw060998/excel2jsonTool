# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved

import os
from pathlib import Path
from typing import Dict, Optional
import re
from log import log_warn

def generate_xml_summary(origin_str: str) -> str:
    """
    生成 XML 注释的通用方法。
    - 如果 origin_str 是单行，返回单行形式: /// <summary> text </summary>
    - 如果 origin_str 是多行，返回多行形式:
      /// <summary>
      /// line1
      /// line2
      /// </summary>
    保证注释内容与之前保持一致。
    """
    if origin_str is None:
        origin_str = ""
    lines = origin_str.splitlines()
    if len(lines) <= 1:
        return f"/// <summary> {origin_str} </summary>"
    else:
        str_text = "\n".join([f"/// {line}" for line in lines])
        return f"/// <summary>\n{str_text}\n/// </summary>"

auto_generated_summary_string = generate_xml_summary("This is auto-generated, don't modify manually")

enum_namespace = "ConfigDataName"
I_CONFIG_RAW_INFO_NAME = "IConfigRawInfo"

def generate_enum_file_from_sheet(sheet, enum_tag, output_folder):
    enum_type_name = sheet.title.replace(enum_tag, "")
    enum_rows = sheet.iter_rows(min_row=2)
    enum_names, enum_values, remarks = zip(
        *[(row[0].value, row[1].value, row[2].value) for row in enum_rows])
    generate_enum_file(enum_type_name, enum_names, enum_values, remarks, enum_namespace, output_folder)


def generate_enum_file(enum_type_name, enum_names, enum_values, remarks, name_space, output_folder):
    file_content = f"namespace {name_space}\n{{\n\t{auto_generated_summary_string}\n\tpublic enum {enum_type_name}\n\t{{\n"

    for i, key in enumerate(enum_names):
        if remarks is not None and remarks[i] is not None:
            file_content += f'\t\t{generate_xml_summary(remarks[i])}\n'
        file_content += f'\t\t{key} = {enum_values[i]},\n\n'

    file_content += "\t}\n}"

    cs_file_path = os.path.join(output_folder, f"{enum_type_name}.cs")
    write_to_file(file_content, cs_file_path)


USING_NAMESPACE_STR = "\n".join([
    "using System.Collections.Generic;",
    "using Newtonsoft.Json;",
    "\n",
    ""
])

# CONFIG_DATA_ATTRIBUTE_STR = "[ConfigData]"
PRIVATE_STATIC_FIELD_STR = "private static {0} {1};"
NAMESPACE_WRAPPER_STR = "namespace Data.TableScript\n{{\n{0}\n}}"

# 提取缩进级别可配置
def add_indentation(input_str, indent="\t"):
    return "\n".join([indent + line for line in input_str.splitlines()])


def convert_type_to_csharp(type_str):
    type_mappings = {'list': 'List', 'dict': 'Dictionary'}
    for key, value in type_mappings.items():
        type_str = type_str.replace(key, value)
    return type_str.replace('(', '<').replace(')', '>')


# 包装类结构，支持可配置缩进
def wrap_class_str(class_name, class_content_str, interface_name="", indent_level=1):
    interface_part = f" : {interface_name}" if interface_name else ""
    indentation = "\t" * indent_level
    indented_content = add_indentation(class_content_str, indentation)
    return f"public class {class_name}{interface_part}\n{{\n{indented_content}\n}}"


# 工具：把任意字段名净化成合法的 C# 参数名（最小化处理）
def _sanitize_param_name(name: str) -> str:
    # 保留字母、数字、下划线，其他替换为下划线；如果以数字开头，加下划线
    s = re.sub(r'[^0-9a-zA-Z_]', '_', name)
    if re.match(r'^[0-9]', s):
        s = '_' + s
    return s


# 工具：把字段名转为 PascalCase，用于方法名后缀（例如 id -> Id, id_group -> IdGroup）
def _to_pascal_case(name: str) -> str:
    parts = re.split(r'[^0-9a-zA-Z]+', name)
    parts = [p for p in parts if p]
    return ''.join(p[0].upper() + p[1:] if len(p) > 0 else '' for p in parts)


def generate_script_file(sheet_name: str,
                         properties_dict: Dict[str, str],
                         property_remarks: Dict[str, str],
                         output_folder: str,
                         need_generate_keys: bool = False,
                         file_suffix: str = "Data",
                         composite_keys: bool = False,
                         composite_multiplier: int = 46340,
                         composite_key_fields: Optional[Dict[str, str]] = None):
    """
    生成 C# 脚本。
    参数说明：
      - composite_keys: 是否启用组合键相关方法（仅当 composite_key_fields 非 None 且为合法字典时生效）
      - composite_multiplier: 合并乘数（与 Python 端保持一致）
      - composite_key_fields: 可选，形如 {"key1": "id", "key2": "group"}，表示真实字段名用于参数与方法命名
    """

    info_class = f"{auto_generated_summary_string}\n{generate_info_class(sheet_name, properties_dict, property_remarks)}"
    data_class = f"{generate_data_class(sheet_name, need_generate_keys, composite_keys, composite_multiplier, composite_key_fields)}"
    file_content = f"{info_class}\n\n{data_class}"
    final_file_content = USING_NAMESPACE_STR + NAMESPACE_WRAPPER_STR.format(add_indentation(file_content))

    cs_file_path = os.path.join(output_folder, f"{sheet_name}{file_suffix}.cs")
    write_to_file(final_file_content, cs_file_path)


def generate_info_class(class_name, properties_dict, property_remarks):
    def format_property_remark(remark):
        # 使用通用的 XML 注释生成方法，保持原先多行/单行格式不变
        return generate_xml_summary(remark)

    # 生成已有字段
    converted_props = []
    property_access_decorator = "{ get; private set; }"
    for key, value in properties_dict.items():
        # 添加 JsonProperty 特性和私有 setter
        converted_props.append(
            f"{format_property_remark(property_remarks[key])}\n"
            f"[JsonProperty(\"{key}\")]\n"
            f"public {convert_type_to_csharp(value)} {key} {property_access_decorator}"
        )

    # 自动补齐 id / name，并打印警告
    added_any = False
    if "id" not in properties_dict:
        converted_props.append(
            f"{generate_xml_summary('Auto-added to satisfy IConfigRawInfo')}\n"
            "[JsonProperty(\"id\")]\n"
            f"public int id {property_access_decorator}"
        )
        log_warn(f"Info class is missing property 'id'. Auto-generated 'public int id {property_access_decorator}'.")
        added_any = True

    if "name" not in properties_dict:
        converted_props.append(
            f"{generate_xml_summary('Auto-added to satisfy IConfigRawInfo')}\n"
            "[JsonProperty(\"name\")]\n"
            f"public string name {property_access_decorator}"
        )
        log_warn(f"Info class is missing property 'name'. Auto-generated 'public string name {property_access_decorator}'.")
        added_any = True

    if added_any:
        log_warn(f"{class_name}Info was missing required properties for IConfigRawInfo; they were auto-added.")

    converted_properties = "\n\n".join(converted_props)

    # 让 Info 实现 IConfigRawInfo
    return wrap_class_str(class_name + "Info", converted_properties, interface_name=I_CONFIG_RAW_INFO_NAME)


def generate_data_class(sheet_name: str,
                        need_generate_keys: bool,
                        composite_keys: bool,
                        composite_multiplier: int,
                        composite_key_fields: Optional[Dict[str, str]]):
    """
    生成数据类字符串（仅包含类声明与必要的 CompositeMultiplier 覆写）。
    - 如果 need_generate_keys 为 True，则继承 ConfigDataWithKey<{SheetName}Keys, {SheetName}Info>
    - 否则如果 composite_keys 为 True，则继承 ConfigDataWithCompositeId<{SheetName}Info> 并实现 CompositeMultiplier
    - 否则继承 ConfigDataBase<{SheetName}Info>
    注意：若 need_generate_keys 和 composite_keys 同时为 True，则按 need_generate_keys 优先。
    """

    class_name = f"{sheet_name}Config"

    # 决定基类（优先级： need_generate_keys -> composite_keys -> 默认 ConfigDataBase）
    if need_generate_keys:
        base_class = f"ConfigDataWithKey<{sheet_name}Info, {sheet_name}Keys>"
    elif composite_keys:
        base_class = f"ConfigDataWithCompositeId<{sheet_name}Info>"
    else:
        base_class = f"ConfigDataBase<{sheet_name}Info>"

    # 如果选用了 composite 基类，则实现 CompositeMultiplier 属性
    composite_override = ""
    if composite_keys and base_class.startswith("ConfigDataWithCompositeId"):
        # 使用 expression-bodied 属性以保持代码简洁
        composite_override = f"protected override int CompositeMultiplier => {composite_multiplier};"

    # 生成类体内容（仅包含可能需要的覆盖成员或注释）
    parts = []

    if composite_override:
        parts.append(composite_override)

    class_content = "\n\n".join(parts) if parts else ""

    # 将原先类内的说明性注释移到类声明上方，使用 XML 注释格式
    header_comment = generate_xml_summary(
        f"Config data class for {sheet_name}. Generated by tool.\nQuery methods are provided by ConfigManager; keep this class minimal."
    )

    # wrap_class_str 会在 class 名称后加上 " : {interface_name}"，因此把 base_class 传入 interface_name 参数
    return f"{header_comment}\n" + wrap_class_str(
        f"{class_name}",
        class_content,
        interface_name=base_class
    )


created_files = []

def write_to_file(content, file_path):
    try:
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)
        print(f"成功生成文件: {file_path}")
        created_files.append(str(Path(file_path).resolve())) # 支持绝对路径和相对路径
    except Exception as e:
        print(f"写入文件失败: {file_path}, 错误: {e}")

def get_create_files():
    return created_files
