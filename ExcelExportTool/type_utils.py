"""Type utilities: parse annotations and convert to C# type strings.

Extracted from worksheet_data and cs_generation to centralize type logic.
"""
from typing import Tuple, Optional


def parse_type_annotation(type_str: str) -> Tuple[str, Optional[str]]:
    t = (type_str or "").strip().lower()

    def base_norm(s: str) -> str:
        s = s.strip().lower()
        if s in ("int", "int32", "integer"): return "int"
        if s in ("float", "double"): return "float"
        if s in ("str", "string"): return "string"
        if s in ("bool", "boolean"): return "bool"
        return s

    if t.startswith("list(") and t.endswith(")"):
        return "list", base_norm(t[5:-1])
    if t.startswith("dict(") and t.endswith(")"):
        return "dict", None
    return "scalar", base_norm(t)


def convert_type_to_csharp(type_str: str) -> str:
    """Convert a short annotation like 'list' or 'dict' to C# representation."""
    type_mappings = {"list": "List", "dict": "Dictionary"}
    for key, value in type_mappings.items():
        type_str = type_str.replace(key, value)
    return type_str.replace("(", "<").replace(")", ">")
