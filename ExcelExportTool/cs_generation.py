# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved

import os
from pathlib import Path
from typing import Dict, Optional
import re

def get_formatted_summary_string(origin_str):
    return f"/// <summary> {origin_str} </summary>"

auto_generated_summary_string = get_formatted_summary_string("This is auto-generated, don't modify manually")

enum_namespace = "ConfigDataName"

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
            file_content += f'\t\t{get_formatted_summary_string(remarks[i])}\n'
        file_content += f'\t\t{key} = {enum_values[i]},\n\n'

    file_content += "\t}\n}"

    cs_file_path = os.path.join(output_folder, f"{enum_type_name}.cs")
    write_to_file(file_content, cs_file_path)


USING_NAMESPACE_STR = "\n".join([
    "using System;",
    "using System.Collections.Generic;",
    "using System.Linq;"
    "\n",
    ""
])

CONFIG_DATA_ATTRIBUTE_STR = "[ConfigData]"
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
    indentation = "\t" * indent_level  # 动态调整缩进级别
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
    data_class = f"{CONFIG_DATA_ATTRIBUTE_STR}\n{generate_data_class(sheet_name, need_generate_keys, composite_keys, composite_multiplier, composite_key_fields)}"
    file_content = f"{info_class}\n\n{data_class}"
    final_file_content = USING_NAMESPACE_STR + NAMESPACE_WRAPPER_STR.format(add_indentation(file_content))

    cs_file_path = os.path.join(output_folder, f"{sheet_name}{file_suffix}.cs")
    write_to_file(final_file_content, cs_file_path)


def generate_info_class(class_name, properties_dict, property_remarks):
    def format_property_remark(remark):
        str_text = "\n".join([f"/// {line}" for line in remark.splitlines()])
        return f"/// <summary>\n{str_text}\n/// </summary>"

    converted_properties = "\n\n".join([
        f"{format_property_remark(property_remarks[key])}\npublic {convert_type_to_csharp(value)} {key} {{ get; set; }}"
        for key, value in properties_dict.items()
    ])
    return wrap_class_str(class_name + "Info", converted_properties)


def generate_data_class(sheet_name: str,
                        need_generate_keys: bool,
                        composite_keys: bool,
                        composite_multiplier: int,
                        composite_key_fields: Optional[Dict[str, str]]):
    """
    生成数据类字符串。若 composite_keys 且 composite_key_fields 提供了真实字段名，
    则生成以真实字段名为参数名的方法（例如 GetDataByIdAndGroup）。
    """

    class_name = f"{sheet_name}Config"
    property_name = "_data"
    property_type_name = f"Dictionary<int, {sheet_name}Info>"

    # Data Property
    data_property = PRIVATE_STATIC_FIELD_STR.format(property_type_name, property_name)

    # Initialization Method
    init_method = (
        f"public static void Initialize()\n{{\n"
        f"\t{property_name} = ConfigDataUtility.DeserializeConfigData<{property_type_name}>(nameof({class_name}));\n"
        f"}}"
    )

    # Get Method (By ID)
    permission_str = "private" if need_generate_keys else "public"
    default_method_name = "GetDataById"
    exception_msg_str = """$\"Can not find the config data by id: {id}.\""""
    get_method = (
        f"{permission_str} static {sheet_name}Info {default_method_name}(int id)\n{{\n"
        f"\tif({property_name}.TryGetValue(id, out var result))\n\t{{\n"
        f"\t\treturn result;\n\t}}\n"
        f"\tthrow new InvalidOperationException({exception_msg_str});\n}}"
    )

    # Get Method (By Key) - 原有枚举/字符串 key
    get_method_with_enumkey = ""
    get_method_with_strkey = ""
    if need_generate_keys:
        key_name = f"{sheet_name}Keys"
        get_method_name_by_key = "GetDataByKey"

        # Get By Enum Key
        key_enum_param = "keyEnum"
        get_method_with_enumkey = (
            f"public static {sheet_name}Info {get_method_name_by_key}({key_name} {key_enum_param})\n{{\n"
            f"\treturn {default_method_name}((int){key_enum_param});\n}}"
        )

        # Get By String Key
        key_str_param = "keyStr"
        exception_msg_str = """$\"Can not parse the config data key: {keyStr}.\""""
        get_method_with_strkey = (
            f"public static {sheet_name}Info {get_method_name_by_key}(string {key_str_param})\n{{\n"
            f"\tif(Enum.TryParse<{key_name}>({key_str_param}, out var {key_enum_param}))\n\t{{\n"
            f"\t\treturn {get_method_name_by_key}({key_enum_param});\n\t}}\n"
            f"\tthrow new InvalidOperationException({exception_msg_str});\n}}"
        )

    # 组合键常量与方法（可能基于真实字段名）
    composite_constant_str = ""
    combine_method_str = ""
    get_by_composite_str = ""
    try_get_by_composite_str = ""
    # 以及基于真实字段名的命名（若 composite_key_fields 有值）
    specific_get_str = ""
    specific_try_get_str = ""
    if composite_keys:
        composite_constant_str = f"private const int COMPOSITE_MULTIPLIER = {composite_multiplier};\n"

        combine_method_str = (
            "public static int CombineKey(int key1, int key2)\n{\n"
            "\tif (key1 is < 0 or >= COMPOSITE_MULTIPLIER) throw new ArgumentOutOfRangeException(nameof(key1));\n"
            "\tif (key2 is < 0 or >= COMPOSITE_MULTIPLIER) throw new ArgumentOutOfRangeException(nameof(key2));\n"
            "\treturn key1 * COMPOSITE_MULTIPLIER + key2;\n}"
        )

        # 如果 composite_key_fields 给出真实字段名，例如 {"key1":"id","key2":"group"}
        if composite_key_fields and isinstance(composite_key_fields, dict):
            real1 = composite_key_fields.get("key1")
            real2 = composite_key_fields.get("key2")
            if isinstance(real1, str) and isinstance(real2, str) and real1 and real2:
                # 生成合法参数名与方法后缀
                param1 = _sanitize_param_name(real1)
                param2 = _sanitize_param_name(real2)

                # Generate GetDataBy<Real1>And<Real2>
                specific_get_str = (
                    f"public static {sheet_name}Info GetDataByCompositeKey(int {param1}, int {param2})\n{{\n"
                    f"\t// Use combined key generated from ({real1},{real2})\n"
                    f"\treturn {default_method_name}(CombineKey({param1}, {param2}));\n}}"
                )
            else:
                specific_get_str = (
                    f"public static {sheet_name}Info GetDataByCompositeKey(int key1, int key2)\n{{\n"
                    f"\treturn {default_method_name}(CombineKey(key1, key2));\n}}"   
                )

    # Select Value Collection Method
    select_value_collection_method = (
        f"public static IEnumerable<TResult> SelectValueCollection<TResult>(Func<{sheet_name}Info, TResult> selector)\n{{\n"
        f"\treturn {property_name}.Values.Select(selector);\n}}"
    )

    # Get Info Collection Method
    get_info_collection_method = (
        f"public static IEnumerable<{sheet_name}Info> GetInfoCollection(Func<{sheet_name}Info, bool> predicate)\n{{\n"
        f"\treturn {property_name}.Values.Where(predicate);\n}}"
    )

    # 构建类内容列表
    class_parts = [
        data_property,
        composite_constant_str,
        init_method,
        get_method,
        get_method_with_enumkey,
        get_method_with_strkey,
        combine_method_str,
        get_by_composite_str,
        try_get_by_composite_str,
        specific_get_str,
        specific_try_get_str,
        select_value_collection_method,
        get_info_collection_method
    ]

    # 过滤空字符串并确保每个部分之间只有一个空行
    non_empty_parts = [part for part in class_parts if part.strip()]
    class_content = '\n\n'.join(non_empty_parts)

    return wrap_class_str(
        f"{class_name}",
        class_content
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