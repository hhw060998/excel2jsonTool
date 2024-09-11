import sys

def read_cell_values(sheet, row_index):
    return [cell.value for cell in sheet[row_index]]

def check_repeating_values(values):
    if len(values) != len(set(values)):
        print(f"■■■■■■■■■发现重复的字段: {set([x for x in values if values.count(x) > 1])}！■■■■■■■■")
        sys.exit()



