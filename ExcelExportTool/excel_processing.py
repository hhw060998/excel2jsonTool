# Author: huhongwei 306463233@qq.com
# MIT License
from exceptions import DuplicateFieldError

def read_cell_values(sheet, row_index):
    return [cell.value for cell in sheet[row_index]]

def check_repeating_values(values):
    dup = {v for v in values if values.count(v) > 1}
    if dup:
        raise DuplicateFieldError(dup)



