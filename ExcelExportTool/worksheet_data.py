# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10
# MIT License 
# All rights reserved


from cs_generation import generate_script_file, generate_enum_file, write_to_file
from excel_processing import read_cell_values, check_repeating_values
from data_processing import convert_to_type, available_csharp_enum_name
import sys
import json

import json

class WorksheetData:
    def __init__(self, worksheet):
        self.name = worksheet.title
        self.worksheet = worksheet
        self.cell_values = {i: read_cell_values(worksheet, i) for i in range(1, 7)}
        self.remarks = self.cell_values[1]
        self.headers = self.cell_values[2]
        self.data_types = self.cell_values[3]
        self.data_labels = self.cell_values[4]
        self.field_names = self.cell_values[5]
        self.default_values = self.cell_values[6]
        self.row_data = list(worksheet.iter_rows(min_row=7, min_col=2))
        check_repeating_values(self.field_names)
        self.need_generate_keys = self.__need_generate_keys()

    def __need_generate_keys(self):
        property_types = self.__get_properties_dict()
        return list(property_types.values())[0] == "string"

    @staticmethod
    def __convert_to_csharp_type(type_str):
        type_mappings = {'list': 'List', 'dict': 'Dictionary'}
        for key, value in type_mappings.items():
            type_str = type_str.replace(key, value)
        return type_str.replace('(', '<').replace(')', '>')

    def __get_properties_dict(self):
        return {
            field_name: self.__convert_to_csharp_type(self.data_types[index])
            for index, field_name in enumerate(self.field_names)
            if self.data_labels[index] != "ignore" and index > 0
        }

    def __get_property_remarks(self):
        return {
            field_name: f"{self.headers[index]}: {self.remarks[index]}" if self.remarks[index] else self.headers[index]
            for index, field_name in enumerate(self.field_names)
            if self.data_labels[index] != "ignore" and index > 0
        }

    def __generate_enum_keys_csfile(self, output_folder):
        enum_type_name = f"{self.name}Keys"
        enum_names = []
        enum_values = []
        index = 0

        for row in self.row_data:
            for col_index, cell in enumerate(row):
                if col_index == 0:
                    if available_csharp_enum_name(cell.value):
                        enum_names.append(cell.value)
                        enum_values.append(index)
                        index += 1
                    else:
                        print(f"第{index + 1}行第1列的值{cell.value}不是合法的c#枚举名，无法生成主键！")
                        sys.exit()

        generate_enum_file(enum_type_name, enum_names, enum_values, None, "Data.TableScript", output_folder)


    def generate_json(self, output_folder):
        data = {}
        serial_key = 0
        for row in self.row_data:
            row_data = {}
            row_data_key = None
            row_data_dict = {}
            for col_index, cell in enumerate(row):
                if col_index == 0 :
                    if self.need_generate_keys:
                        row_data_key = serial_key
                        serial_key += 1
                    else:
                        row_data_key = int(cell.value)

                data_label = self.data_labels[col_index + 1]

                if data_label == "ignore":
                    continue

                default_value = self.default_values[col_index + 1]
                data_name = self.field_names[col_index + 1]
                data_type_str = self.data_types[col_index + 1]

                value = None
                if cell.value is None:
                    if default_value is None:
                        if data_label == "required":
                            print(f"{data_name}的label为required！但是值为空且没有默认值，退出导表")
                            sys.exit()
                        else:
                            value = convert_to_type(data_type_str, cell.value)
                    else:
                        value = convert_to_type(data_type_str, default_value)
                else:
                    value = convert_to_type(data_type_str, cell.value)

                row_data[data_name] = value
                row_data_dict[row_data_key] = row_data

            data.update(row_data_dict)

        file_content = json.dumps(data, ensure_ascii=False, indent=4)
        file_path = f"{output_folder}/{self.name}Config.json"
        write_to_file(file_content, file_path)


    def generate_script(self, output_folder):
        properties_dict = self.__get_properties_dict()
        property_remakes = self.__get_property_remarks()
        generate_script_file(self.name, properties_dict, property_remakes, output_folder, self.need_generate_keys)
        # 如果properties_dict没有名为id的元素，或者名为id字段的元素类型不是int，则生成枚举文件
        if self.need_generate_keys:
            self.__generate_enum_keys_csfile(output_folder)




