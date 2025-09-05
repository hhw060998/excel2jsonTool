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