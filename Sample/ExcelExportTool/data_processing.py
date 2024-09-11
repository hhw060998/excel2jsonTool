import re
import sys

# 定义基本类型映射
PRIMITIVE_TYPE_MAPPING = {
    'int': int,
    'float': float,
    'bool': bool,
    'str': str,
    'string': str  # 处理字符串的别名
}

# 定义支持的复杂类型前缀
COMPLEX_TYPE_PREFIXES = ['dict', 'list']


def available_csharp_enum_name(name):
    # 检查是否为有效的C#枚举名称
    return re.match(r'^[A-Za-z_][A-Za-z0-9_]*$', name)


def convert_to_type(type_str, value):
    # 处理基本类型转换
    if type_str in PRIMITIVE_TYPE_MAPPING:
        return convert_to_primitive_type(type_str, value)

    # 处理复杂类型转换
    if any(prefix in type_str for prefix in COMPLEX_TYPE_PREFIXES):
        return convert_to_complex_type(type_str, value)

    print(f"Unsupported data type: {type_str}")
    sys.exit()


def convert_to_primitive_type(type_str, value):
    # 从映射中获取转换函数，提供默认值处理None的情况
    convert_func = PRIMITIVE_TYPE_MAPPING.get(type_str, lambda x: x)
    if value is None:
        return convert_func('') if type_str in ['str', 'string'] else convert_func(0)
    return convert_func(value)


def convert_to_complex_type(type_str, value):
    # 处理字典类型
    if "dict" in type_str:
        return process_dict_type(type_str, value)

    # 处理列表类型
    if "list" in type_str:
        return process_list_type(type_str, value)

    return value


def process_dict_type(type_str, value):
    dict_data = {}
    type_match = re.search(r'\((.*)\)', type_str)

    if type_match and value is not None:
        key_type_str, value_type_str = map(str.strip, type_match.group(1).split(","))
        key_type = PRIMITIVE_TYPE_MAPPING.get(key_type_str, str)  # 默认key类型为str
        value_type = PRIMITIVE_TYPE_MAPPING.get(value_type_str, str)  # 默认value类型为str

        for line in value.split("\n"):
            if ":" in line:
                key, val = map(str.strip, line.split(":"))
                dict_data[key_type(key)] = value_type(val)

    return dict_data


def process_list_type(type_str, value):
    list_data = []
    type_match = re.search(r'\((.*)\)', type_str)

    if type_match and value is not None:
        element_type_str = type_match.group(1).strip()
        element_type = PRIMITIVE_TYPE_MAPPING.get(element_type_str, str)  # 默认元素类型为str

        if isinstance(value, str):
            list_data = [element_type(elem.strip()) for elem in value.split(",") if elem]
        else:
            list_data = [element_type(value)]

    return list_data
