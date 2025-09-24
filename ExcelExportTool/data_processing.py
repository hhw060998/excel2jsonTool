# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License
# All rights reserved

import re
from typing import Any, Callable, Dict, List, Optional
from exceptions import UnknownCustomTypeError, CustomTypeParseError
from log import log_warn
import keyword
import re


_C_SHARP_KEYWORDS = set([
    # C# keywords (常见)
    'abstract','as','base','bool','break','byte','case','catch','char','checked','class','const','continue',
    'decimal','default','delegate','do','double','else','enum','event','explicit','extern','false','finally',
    'fixed','float','for','foreach','goto','if','implicit','in','int','interface','internal','is','lock','long',
    'namespace','new','null','object','operator','out','override','params','private','protected','public','readonly',
    'ref','return','sbyte','sealed','short','sizeof','stackalloc','static','string','struct','switch','this','throw',
    'true','try','typeof','uint','ulong','unchecked','unsafe','ushort','using','virtual','void','volatile','while'
])


def is_valid_csharp_identifier(name: str) -> bool:
    """检查 name 是否为合法的 C# 标识符（按常规约定：以字母或下划线开头，仅包含字母、数字或下划线，且不是关键字）。"""
    if not isinstance(name, str) or not name:
        return False
    # 允许 PascalCase / camelCase: 只需首字符为字母或下划线
    if not re.match(r'^[A-Za-z_][A-Za-z0-9_]*$', name):
        return False
    if name in _C_SHARP_KEYWORDS:
        return False
    return True

# 基本类型映射
PRIMITIVE_TYPE_MAPPING: Dict[str, Callable[[Any], Any]] = {
    "int": int,
    "float": float,
    "bool": lambda x: str(x).lower() in ("1", "true", "yes"),
    "str": str,
    "string": str,  # 兼容别名
}

# ================= 自定义类型注册机制 =================
class _CustomTypeHandler:
    def __init__(self, parser: Callable[[Optional[str]], Any]):
        self.parser = parser

class CustomTypeRegistry:
    def __init__(self):
        self._handlers: Dict[str, _CustomTypeHandler] = {}

    def register(self, full_name: str, parser: Callable[[Optional[str]], Any]):
        self._handlers[full_name] = _CustomTypeHandler(parser)

    def parse(self, full_name: str, raw: Any, field: str | None = None, sheet: str | None = None):
        h = self._handlers.get(full_name)
        if not h:
            raise UnknownCustomTypeError(full_name, field, sheet)
        try:
            return h.parser(None if raw is None else str(raw))
        except Exception as e:
            raise CustomTypeParseError(full_name, str(raw), str(e), field, sheet)

    def contains(self, full_name: str) -> bool:
        return full_name in self._handlers

    def all_types(self):
        return list(self._handlers.keys())

custom_type_registry = CustomTypeRegistry()

# 是否启用未注册自定义类型的通用回退解析
GENERIC_CUSTOM_TYPE_FALLBACK = True

def _generic_custom_type_object(full_name: str, raw: Optional[str]):
    """通用自定义类型打包：按 '#' 切分为 segments，保留原串。
    JSON 结构: {"__type": full_name, "__raw": original, "segments": [..]}
    空值 -> {"__type": full_name, "segments": []}
    """
    if raw is None or raw == "":
        return {"__type": full_name, "segments": []}
    txt = raw.replace("\r\n", "\n")
    parts = [p.strip() for p in txt.split('#')]
    return {"__type": full_name, "__raw": txt, "segments": parts}

def _parse_localized_string_ref(raw: Optional[str]):
    """默认示例: Localization.LocalizedStringRef 形如 文本#上下文 (#可省)."""
    if raw is None or raw == "":
        return {"keyHash": 0, "source": "", "context": ""}
    txt = raw.replace("\r\n", "\n")
    if "#" in txt:
        src, ctx = txt.split("#", 1)
    else:
        src, ctx = txt, ""
    src = src.strip(); ctx = ctx.strip()
    return {"keyHash": 0, "source": src, "context": ctx}

# 注册示例（可通过外部扩展继续添加）
custom_type_registry.register("Localization.LocalizedStringRef", _parse_localized_string_ref)
# =====================================================


def available_csharp_enum_name(name: str) -> bool:
    """检查是否为合法的C#枚举名"""
    return bool(re.match(r"^[A-Za-z_][A-Za-z0-9_]*$", str(name)))


def convert_to_type(type_str: str, value: Any, field: str | None = None, sheet: str | None = None) -> Any:
    """根据类型字符串转换值 (支持基础/list/dict/自定义全限定类型)"""
    if not type_str:
        raise ValueError("空类型定义")
    type_str = type_str.strip()

    # 基础
    if type_str in PRIMITIVE_TYPE_MAPPING:
        return _convert_primitive(type_str, value)
    # 容器
    if type_str.startswith("dict"):
        return _convert_dict(type_str, value)
    if type_str.startswith("list"):
        return _convert_list(type_str, value)
    # 自定义(简单策略: 至少包含一个 . 视为全限定类型)
    if "." in type_str:
        if custom_type_registry.contains(type_str):
            return custom_type_registry.parse(type_str, value, field, sheet)
        if GENERIC_CUSTOM_TYPE_FALLBACK:
            return _generic_custom_type_object(type_str, None if value is None else str(value))
        # 未开启通用回退仍旧报错
        raise UnknownCustomTypeError(type_str, field, sheet)
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
