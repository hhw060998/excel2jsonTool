# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License
# All rights reserved

import json
from typing import Any, Dict, List
from cs_generation import generate_script_file, generate_enum_file, write_to_file
from excel_processing import read_cell_values, check_repeating_values
from data_processing import convert_to_type, available_csharp_enum_name


class WorksheetData:
    def __init__(self, worksheet) -> None:
        self.name: str = worksheet.title
        self.worksheet = worksheet
        self.cell_values: Dict[int, List[Any]] = {
            i: read_cell_values(worksheet, i) for i in range(1, 7)
        }

        self.remarks = self.cell_values[1]
        self.headers = self.cell_values[2]
        self.data_types = self.cell_values[3]
        self.data_labels = self.cell_values[4]
        self.field_names = self.cell_values[5]
        self.default_values = self.cell_values[6]
        self.row_data = list(worksheet.iter_rows(min_row=7, min_col=2))

        check_repeating_values(self.field_names)
        self.need_generate_keys = self._need_generate_keys()

    def _need_generate_keys(self) -> bool:
        """判断是否需要为数据行生成自增 key"""
        property_types = self._get_properties_dict()
        return next(iter(property_types.values()), None) == "string"

    @staticmethod
    def _convert_to_csharp_type(type_str: str) -> str:
        """Excel 类型 -> C# 类型"""
        type_mappings = {"list": "List", "dict": "Dictionary"}
        for key, value in type_mappings.items():
            type_str = type_str.replace(key, value)
        return type_str.replace("(", "<").replace(")", ">")

    def _get_properties_dict(self) -> Dict[str, str]:
        """字段名 -> C# 类型"""
        return {
            field_name: self._convert_to_csharp_type(self.data_types[i])
            for i, field_name in enumerate(self.field_names)
            if self.data_labels[i] != "ignore" and i > 0
        }

    def _get_property_remarks(self) -> Dict[str, str]:
        """字段名 -> 注释（表头: 备注）"""
        return {
            field_name: f"{self.headers[i]}: {self.remarks[i]}" if self.remarks[i] else self.headers[i]
            for i, field_name in enumerate(self.field_names)
            if self.data_labels[i] != "ignore" and i > 0
        }

    def _validate_enum_name(self, name: str, row_index: int) -> None:
        """检查枚举名是否合法"""
        if not available_csharp_enum_name(name):
            raise RuntimeError(f"第{row_index + 1}行第1列的值 {name} 不是合法的C#枚举名，无法生成主键！")

    def _generate_enum_keys_csfile(self, output_folder: str) -> None:
        enum_type_name = f"{self.name}Keys"
        enum_names, enum_values = [], []

        for index, row in enumerate(self.row_data):
            first_cell_value = row[0].value
            self._validate_enum_name(first_cell_value, index)
            enum_names.append(first_cell_value)
            enum_values.append(index)

        generate_enum_file(enum_type_name, enum_names, enum_values, None, "Data.TableScript", output_folder)

    def generate_json(self, output_folder: str) -> None:
        """将表格数据导出为 JSON 文件"""
        data: Dict[Any, Dict[str, Any]] = {}
        serial_key = 0

        for row in self.row_data:
            row_data = {}
            # 处理 key
            if self.need_generate_keys:
                row_data_key = serial_key
                serial_key += 1
            else:
                row_data_key = int(row[0].value)

            for col_index, cell in enumerate(row, start=1):
                data_label = self.data_labels[col_index]

                if data_label == "ignore":
                    continue

                default_value = self.default_values[col_index]
                data_name = self.field_names[col_index]
                data_type_str = self.data_types[col_index]

                cell_value = cell.value
                if cell_value is None:
                    if default_value is None and data_label == "required":
                        raise RuntimeError(f"{data_name} 的label为 required，但值为空且没有默认值")
                    value = convert_to_type(data_type_str, default_value)
                else:
                    value = convert_to_type(data_type_str, cell_value)

                row_data[data_name] = value

            data[row_data_key] = row_data

        file_content = json.dumps(data, ensure_ascii=False, indent=4)
        file_path = f"{output_folder}/{self.name}Config.json"
        write_to_file(file_content, file_path)

    def generate_script(self, output_folder: str) -> None:
        """生成 C# 脚本（必要时生成枚举 Key 文件）"""
        properties_dict = self._get_properties_dict()
        property_remakes = self._get_property_remarks()

        generate_script_file(self.name, properties_dict, property_remakes, output_folder, self.need_generate_keys)

        if self.need_generate_keys:
            self._generate_enum_keys_csfile(output_folder)
