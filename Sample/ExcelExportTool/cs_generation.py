import os


def get_formatted_summary_string(origin_str):
    return f"/// <summary> {origin_str} </summary>"

auto_generated_summary_string = get_formatted_summary_string("This is auto-generated, don't modify manually")

def generate_enum_file_from_sheet(sheet, enum_tag, output_folder):
    enum_type_name = sheet.title.replace(enum_tag, "")
    enum_rows = sheet.iter_rows(min_row=2)
    enum_names, enum_values, remarks = zip(
        *[(row[0].value, row[1].value, row[2].value) for row in enum_rows])
    generate_enum_file(enum_type_name, enum_names, enum_values, remarks, namespace, output_folder)


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
namespace = "ConfigDataName"


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


def generate_script_file(sheet_name, properties_dict, property_remarks, output_folder, need_generate_keys=False, file_suffix="Data"):
    # 通用文件生成流程
    info_class = f"{auto_generated_summary_string}\n{generate_info_class(sheet_name, properties_dict, property_remarks)}"
    data_class = f"{CONFIG_DATA_ATTRIBUTE_STR}\n{generate_data_class(sheet_name, need_generate_keys)}"
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


def generate_data_class(sheet_name, need_generate_keys):
    class_name = f"{sheet_name}Config"
    property_name = "_data"
    property_type_name = f"Dictionary<int, {sheet_name}Info>"
    data_property = PRIVATE_STATIC_FIELD_STR.format(property_type_name, property_name)
    init_method = (
        f"public static void Initialize()\n{{\n"
        f"\t{property_name} = ConfigDataUtility.DeserializeConfigData<{property_type_name}>(nameof({class_name}));\n"
        f"}}"
    )
    get_method = (
        f"public static {sheet_name}Info GetDataById(int id)\n{{\n"
        f"\tif({property_name}.TryGetValue(id, out var result))\n\t{{\n"
        f"\t\treturn result;\n\t}}\n"
        f"\tthrow new InvalidOperationException();\n}}"
    )

    get_method_with_key = ""
    if need_generate_keys:
        key_name = f"{sheet_name}Keys"
        key_str_param = "keyStr"
        key_enum_param = "keyEnum"

        get_method_with_key = (
            f"public static {sheet_name}Info GetDataByKey(string {key_str_param})\n{{\n"
            f"\tif(Enum.TryParse<{key_name}>({key_str_param}, out var {key_enum_param}))\n\t{{\n"
            f"\t\treturn GetDataById((int){key_enum_param});\n\t}}\n"
            f"\tthrow new InvalidOperationException();\n}}\n\n"
        )

    select_value_collection_method = (
        f"public static IEnumerable<TResult> SelectValueCollection<TResult>(Func<{sheet_name}Info, TResult> selector)\n{{\n"
        f"\treturn {property_name}.Values.Select(selector);\n}}"
    )

    get_info_collection_method = (
        f"public static IEnumerable<{sheet_name}Info> GetInfoCollection(Func<{sheet_name}Info, bool> predicate)\n{{\n"
        f"\treturn {property_name}.Values.Where(predicate);\n}}"
    )

    return wrap_class_str(f"{class_name}",
                          f"{data_property}\n\n{init_method}\n\n{get_method}\n\n{get_method_with_key}{select_value_collection_method}\n\n{get_info_collection_method}")


created_files = []

def write_to_file(content, file_path):
    try:
        with open(file_path, 'w', encoding='utf-8') as file:
            file.write(content)
        print(f"成功生成文件: {file_path}")
        created_files.append(os.path.abspath(file_path))  # 使用绝对路径
    except Exception as e:
        print(f"写入文件失败: {file_path}, 错误: {e}")

def get_create_files():
    return created_files
