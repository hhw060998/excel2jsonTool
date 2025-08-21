# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License
# All rights reserved

import re
from typing import Any, Callable, Dict, List

# 基本类型映射
PRIMITIVE_TYPE_MAPPING: Dict[str, Callable[[Any], Any]] = {
    "int": int,
    "float": float,
    "bool": lambda x: str(x).lower() in ("1", "true", "yes"),
    "str": str,
    "string": str,  # 兼容别名
}


def available_csharp_enum_name(name: str) -> bool:
    """检查是否为合法的C#枚举名"""
    return bool(re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", str(name)))


def convert_to_type(type_str: str, value: Any) -> Any:
    """根据类型字符串转换值"""
    type_str = type_str.strip()

    if type_str in PRIMITIVE_TYPE_MAPPING:
        return _convert_primitive(type_str, value)

    if type_str.startswith("dict"):
        return _convert_dict(type_str, value)

    if type_str.startswith("list"):
        return _convert_list(type_str, value)

    raise ValueError(f"Unsupported data type: {type_str}")


def _convert_primitive(type_str: str, value: Any) -> Any:
    """转换为基本类型"""
    converter = PRIMITIVE_TYPE_MAPPING[type_str]
    if value is None:
        return "" if type_str in ("str", "string") else converter(0)
    return converter(value)


def _convert_dict(type_str: str, value: Any) -> Dict[Any, Any]:
    """转换为字典类型，例如 dict(int,string)"""
    result: Dict[Any, Any] = {}
    type_match = re.search(r"\((.*)\)", type_str)

    if not type_match or value is None:
        return result

    key_type_str, value_type_str = map(str.strip, type_match.group(1).split(","))
    key_type = PRIMITIVE_TYPE_MAPPING.get(key_type_str, str)
    val_type = PRIMITIVE_TYPE_MAPPING.get(value_type_str, str)

    for line in str(value).splitlines():
        if ":" in line:
            key, val = map(str.strip, line.split(":", 1))
            try:
                result[key_type(key)] = val_type(val)
            except Exception as e:
                raise ValueError(f"无法将 {line} 转换为 {type_str}: {e}")

    return result


def _convert_list(type_str: str, value: Any) -> List[Any]:
    """转换为列表类型，例如 list(int)"""
    result: List[Any] = []
    type_match = re.search(r"\((.*)\)", type_str)

    if not type_match or value is None:
        return result

    element_type_str = type_match.group(1).strip()
    elem_type = PRIMITIVE_TYPE_MAPPING.get(element_type_str, str)

    if isinstance(value, str):
        return [elem_type(v.strip()) for v in value.split(",") if v.strip()]

    try:
        return [elem_type(value)]
    except Exception as e:
        raise ValueError(f"无法将 {value} 转换为 {type_str}: {e}")
