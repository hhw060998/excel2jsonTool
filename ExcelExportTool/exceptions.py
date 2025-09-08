class ExportError(Exception):
    """基础导表异常"""
    pass

class DuplicateFieldError(ExportError):
    def __init__(self, fields):
        super().__init__(f"发现重复字段: {sorted(fields)}")

class InvalidEnumNameError(ExportError):
    def __init__(self, name, row):
        super().__init__(f"非法枚举名 '{name}' (Excel 行: {row})")

class DuplicatePrimaryKeyError(ExportError):
    def __init__(self, key, row_a, row_b):
        super().__init__(f"主键重复: {key} (行 {row_a} 与 行 {row_b})")

class CompositeKeyOverflowError(ExportError):
    def __init__(self, combined):
        super().__init__(f"组合键溢出: {combined} >= 2^31")

class SheetNameConflictError(ExportError):
    def __init__(self, sheet, f1, f2):
        super().__init__(f"工作表命名冲突: {sheet} 出现在 {f1} 与 {f2}")

class UnknownCustomTypeError(ExportError):
    def __init__(self, type_name: str, field: str | None = None, sheet: str | None = None):
        loc = []
        if field:
            loc.append(f"字段:{field}")
        if sheet:
            loc.append(f"表:{sheet}")
        suffix = (" (" + ", ".join(loc) + ")") if loc else ""
        super().__init__(f"未注册的自定义类型: {type_name}{suffix}")

class CustomTypeParseError(ExportError):
    def __init__(self, type_name: str, raw: str, reason: str, field: str | None = None, sheet: str | None = None):
        loc = []
        if field:
            loc.append(f"字段:{field}")
        if sheet:
            loc.append(f"表:{sheet}")
        suffix = (" (" + ", ".join(loc) + ")") if loc else ""
        super().__init__(f"自定义类型解析失败: {type_name} 原值:[{raw}] -> {reason}{suffix}")