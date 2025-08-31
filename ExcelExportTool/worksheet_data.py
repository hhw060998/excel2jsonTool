# Author: huhongwei 306463233@qq.com
# Created: 2024-09-10 (updated)
# MIT License
# All rights reserved

import json
from typing import Any, Dict, List, Optional
from collections import defaultdict
import re

from cs_generation import generate_script_file, generate_enum_file, write_to_file
from excel_processing import read_cell_values, check_repeating_values
from data_processing import convert_to_type, available_csharp_enum_name
from log import log_warn


class WorksheetData:
    """
    处理单个 Worksheet 的数据导出逻辑。
    - 支持三种主键策略：
        1) 字符串枚举主键（need_generate_keys == True）
        2) 单列 int 主键（默认旧行为）
        3) 组合 int 主键（key1:RealName, key2:RealName 出现在 field_names 的前两项）
           组合映射算法： combined = key1 * MULTIPLIER + key2 （无冲突）
    - 当使用前缀配置（key1:xxx / key2:yyy）时，会解析出真实字段名 xxx 和 yyy，
      并在生成的 C# 方法中使用真实字段名作为参数名和注释提示。
    """

    # 组合键参数（默认保证合并结果可装入 int32）
    MAX_KEY2 = 46340         # key2 的上限（exclusive）：0 <= key2 < MAX_KEY2
    MULTIPLIER = MAX_KEY2    # MULTIPLIER = MAX_KEY2

    KEY1_PREFIX_RE = re.compile(r"^\s*key1\s*:\s*(?P<name>.+)\s*$", re.IGNORECASE)
    KEY2_PREFIX_RE = re.compile(r"^\s*key2\s*:\s*(?P<name>.+)\s*$", re.IGNORECASE)

    def __init__(self, worksheet) -> None:
        self.name: str = worksheet.title
        self.worksheet = worksheet

        # cell_values 的索引遵循原来约定：1..6 分别为 remarks, headers, data_types, data_labels, field_names, default_values
        self.cell_values: Dict[int, List[Any]] = {
            i: read_cell_values(worksheet, i) for i in range(1, 7)
        }

        self.remarks = self.cell_values[1]
        self.headers = self.cell_values[2]
        self.data_types = self.cell_values[3]
        self.data_labels = self.cell_values[4]
        self.field_names = self.cell_values[5]
        self.default_values = self.cell_values[6]

        # 数据行：从 Excel 第 7 行开始，且最左列从第2列开始读取（与你原来实现一致）
        self.row_data = list(worksheet.iter_rows(min_row=7, min_col=2))

        # 保持原有的字段重复检查（针对 field_names）
        check_repeating_values(self.field_names)

        # 原有逻辑判断是否需要生成字符串枚举 key
        self.need_generate_keys = self._need_generate_keys()

        # 解析组合键配置（要求出现在 field_names 的前两列，并采用 key1:xxx / key2:yyy 格式）
        self.composite_keys = False
        self.composite_key_fields: Dict[str, str] = {}  # e.g. {"key1": "id", "key2": "group"}
        self._detect_composite_keys_with_prefixes_in_first_two_columns()

        # 如果需要生成字符串枚举键（原逻辑），初始化时做合法性与重复检查
        if self.need_generate_keys:
            self._check_duplicate_enum_keys()

        # 如果检测到组合键配置，则在初始化阶段检查 key1/key2 的合法性、范围与组合唯一性
        if self.composite_keys:
            self._check_duplicate_composite_keys()
            
        # 当使用单列 int 主键但第一列字段名不是 id 时，记录标志并准备一次性警告
        self.first_int_pk_not_named_id_warned = False

    def _need_generate_keys(self) -> bool:
        """判断是否需要为数据行生成自增 key（原逻辑）"""
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
        """
        字段名 -> C# 类型
        注意：字段名可能包含前缀 'key1:' / 'key2:'，这里会返回真实字段名（去掉前缀）
        保持原来 i>0 的约定。
        """
        result = {}
        for i, raw_field_name in enumerate(self.field_names):
            if not (self.data_labels[i] != "ignore" and i > 0):
                continue
            actual_name = self._actual_field_name(i)
            result[actual_name] = self._convert_to_csharp_type(self.data_types[i])
        return result

    def _get_property_remarks(self) -> Dict[str, str]:
        """字段名 -> 注释（表头: 备注），字段名使用真实名字（去掉 key1:/key2: 前缀）"""
        result = {}
        for i, raw_field_name in enumerate(self.field_names):
            if not (self.data_labels[i] != "ignore" and i > 0):
                continue
            actual_name = self._actual_field_name(i)
            result[actual_name] = (
                f"{self.headers[i]}: {self.remarks[i]}" if self.remarks[i] else self.headers[i]
            )
        return result

    def _actual_field_name(self, field_index: int) -> str:
        """
        返回 field_names[field_index] 对应的“真实字段名”：
        - 如果字段是 'key1:xxx' 或 'key2:yyy' 格式，返回 xxx / yyy（不包含前缀）
        - 否则返回原始 field_names[field_index]
        注意：field_index 对应你原来使用的索引（从 0 开始），generate_json 使用 enumerate(row, start=1) 时要匹配该索引。
        """
        raw = self.field_names[field_index]
        if not isinstance(raw, str):
            return str(raw)
        m1 = self.KEY1_PREFIX_RE.match(raw)
        if m1:
            return m1.group("name").strip()
        m2 = self.KEY2_PREFIX_RE.match(raw)
        if m2:
            return m2.group("name").strip()
        return raw

    def _detect_composite_keys_with_prefixes_in_first_two_columns(self) -> None:
        """
        强制检查：
          - field_names[1] 以 key1:RealName 格式出现（不区分大小写）
          - field_names[2] 以 key2:RealName 格式出现（不区分大小写）
          - data_types[1] 与 data_types[2] 都包含 "int"
        并在 self.composite_key_fields 中保存真实字段名：
          self.composite_key_fields = {"key1": "id", "key2": "group"}
        """
        try:
            if len(self.field_names) <= 2:
                self.composite_keys = False
                return

            f1 = self.field_names[1]
            f2 = self.field_names[2]
            if not (isinstance(f1, str) and isinstance(f2, str)):
                self.composite_keys = False
                return

            m1 = self.KEY1_PREFIX_RE.match(f1)
            m2 = self.KEY2_PREFIX_RE.match(f2)
            if not (m1 and m2):
                self.composite_keys = False
                return

            dt1 = self.data_types[1] if len(self.data_types) > 1 else None
            dt2 = self.data_types[2] if len(self.data_types) > 2 else None
            if not (isinstance(dt1, str) and isinstance(dt2, str) and "int" in dt1.strip().lower() and "int" in dt2.strip().lower()):
                self.composite_keys = False
                return

            # 解析真实字段名并启用 composite_keys
            real1 = m1.group("name").strip()
            real2 = m2.group("name").strip()
            if not real1 or not real2:
                self.composite_keys = False
                return

            self.composite_keys = True
            self.composite_key_fields = {"key1": real1, "key2": real2}
        except Exception:
            self.composite_keys = False
            self.composite_key_fields = {}

    def _validate_enum_name(self, name: str, excel_row: int) -> None:
        """检查枚举名是否合法（excel_row 为真实 Excel 行号，用于错误提示）"""
        if not available_csharp_enum_name(name):
            raise RuntimeError(f"第{excel_row}行第1列的值 {name} 不是合法的C#枚举名，无法生成主键！")

    def _check_duplicate_enum_keys(self) -> None:
        """
        初始化时检查用于生成枚举的首列（字符串主键）：
        - 验证每个名字是否合法
        - 收集出现的 Excel 行号，若重复则抛错（显示真实 Excel 行号）
        """
        name_to_rows = defaultdict(list)
        for idx, row in enumerate(self.row_data):
            excel_row = idx + 7  # 数据从 Excel 第 7 行开始
            first_cell_value = row[0].value
            self._validate_enum_name(first_cell_value, excel_row)
            name_to_rows[first_cell_value].append(excel_row)

        duplicates = {name: rows for name, rows in name_to_rows.items() if len(rows) > 1}
        if duplicates:
            dup_msgs = []
            for name, rows in duplicates.items():
                dup_msgs.append(f"'{name}' 出现在行 {rows}")
            raise RuntimeError("发现重复的字符串主键（将用于生成枚举），请修正后重试：\n" + "\n".join(dup_msgs))

    def _check_duplicate_composite_keys(self) -> None:
        """
        初始化时检查组合 int 键（要求位于数据前两列 -> row[0], row[1] 且用 key1:real / key2:real 标记）：
        - 检查 key1/key2 是否为整数、是否在允许范围内
        - 检查组合后的 combined 是否唯一（若重复则抛错并显示真实 Excel 行号，及对应实际 (key1,key2)）
        """
        if not self.composite_keys:
            return

        combined_to_rows = defaultdict(list)
        for idx, row in enumerate(self.row_data):
            excel_row = idx + 7
            # key1 对应 row[0]，key2 对应 row[1]
            try:
                key1_cell = row[0].value
                key2_cell = row[1].value
            except IndexError:
                raise RuntimeError(f"第{excel_row}行无法读取 key1/key2（索引越界）")

            if key1_cell is None or key2_cell is None:
                raise RuntimeError(f"第{excel_row}行的 key1/key2 为空，无法作为组合主键。")

            try:
                key1 = int(key1_cell)
                key2 = int(key2_cell)
            except Exception:
                raise RuntimeError(f"第{excel_row}行 key1 或 key2 不能转换为 int (key1={key1_cell}, key2={key2_cell})")

            # 范围检查：要求 0 <= key2 < MAX_KEY2；同时限制 key1 在相同范围（可调整）
            if not (0 <= key2 < self.MAX_KEY2):
                raise RuntimeError(f"第{excel_row}行的 key2={key2} 超出允许范围。要求 0 <= key2 < {self.MAX_KEY2}")
            if not (0 <= key1 < self.MAX_KEY2):
                raise RuntimeError(f"第{excel_row}行的 key1={key1} 超出允许范围。要求 0 <= key1 < {self.MAX_KEY2}")

            combined = key1 * self.MULTIPLIER + key2
            combined_to_rows[combined].append((excel_row, key1, key2))

        duplicates = {c: rows for c, rows in combined_to_rows.items() if len(rows) > 1}
        if duplicates:
            dup_msgs = []
            for c, rows in duplicates.items():
                # 列出每个重复组合对应的行和原始 (key1,key2)
                rows_desc = ", ".join([f"(row {r[0]} -> ({r[1]},{r[2]}))" for r in rows])
                dup_msgs.append(f"组合值 {c} 对应的行: {rows_desc}")
            raise RuntimeError("发现重复的组合主键，请修正后重试：\n" + "\n".join(dup_msgs))

    def _generate_enum_keys_csfile(self, output_folder: str) -> None:
        """当需要 string 枚举键时才调用（保留原有实现）"""
        enum_type_name = f"{self.name}Keys"
        enum_names: List[str] = []
        enum_values: List[int] = []

        for index, row in enumerate(self.row_data):
            first_cell_value = row[0].value
            # 防御性检查（用真实 Excel 行号作为提示）
            self._validate_enum_name(first_cell_value, index + 7)
            enum_names.append(first_cell_value)
            enum_values.append(index)

        generate_enum_file(enum_type_name, enum_names, enum_values, None, "Data.TableScript", output_folder)

    def generate_json(self, output_folder: str) -> None:
        """将表格数据导出为 JSON 文件（支持单列 int 主键 / 自动生成键 / 以及组合键）。
        这里同时确保给每条记录填充 info.id：
            - 字符串主键（枚举）：id = 序号（枚举 int 值）
            - 组合键：id = key1*MULTIPLIER + key2
            - 单列 int 主键：id = 第一列的 int 值；若第一列字段名不是 'id' 则打印一次警告
        """
        data: Dict[Any, Dict[str, Any]] = {}
        serial_key = 0

        # 预取“第一列真实字段名”，用来判断是否叫 id
        first_field_real_name = self._actual_field_name(1) if len(self.field_names) > 1 else None

        for row_idx, row in enumerate(self.row_data):
            row_data = {}

            # 处理主键
            if self.need_generate_keys:
                # 使用自增 serial（枚举）
                row_data_key = serial_key
                serial_key += 1
            elif self.composite_keys:
                # 组合键：取数据前两列 row[0], row[1]
                excel_row = row_idx + 7
                try:
                    key1 = int(row[0].value)
                    key2 = int(row[1].value)
                except Exception:
                    raise RuntimeError(f"第{excel_row}行无法解析 key1/key2 为 int（请检查数据）")
                if not (0 <= key2 < self.MAX_KEY2) or not (0 <= key1 < self.MAX_KEY2):
                    raise RuntimeError(f"第{excel_row}行 key1/key2 超出允许范围，要求 0 <= key < {self.MAX_KEY2}")
                row_data_key = key1 * self.MULTIPLIER + key2
            else:
                # 单列 int 主键（第一列）
                try:
                    row_data_key = int(row[0].value)
                except Exception:
                    excel_row = row_idx + 7
                    raise RuntimeError(f"第{excel_row}行的主键无法转换为 int: {row[0].value}")

                # 第一列字段名不是 id，则打印一次性警告（并在下方把 id 填入）
                if (isinstance(first_field_real_name, str)
                    and first_field_real_name.strip().lower() != "id"
                    and not self.first_int_pk_not_named_id_warned):
                    log_warn(f"表 [{self.name}] 的第一列为 int 且被视作主键，但字段名不是 'id'。"
                        f" 已将该值写入每条记录的 'id' 属性。建议将第一列字段名改为 'id'。")
                    self.first_int_pk_not_named_id_warned = True

            # 填充行字段（列索引从 1 开始对应 field_names[1]）
            for col_index, cell in enumerate(row, start=1):
                data_label = self.data_labels[col_index]
                if data_label == "ignore":
                    continue

                default_value = self.default_values[col_index]
                data_name = self._actual_field_name(col_index)
                data_type_str = self.data_types[col_index]

                cell_value = cell.value
                if cell_value is None:
                    if default_value is None and data_label == "required":
                        raise RuntimeError(f"{data_name} 的label为 required，但值为空且没有默认值")
                    value = convert_to_type(data_type_str, default_value)
                else:
                    value = convert_to_type(data_type_str, cell_value)

                row_data[data_name] = value

            # ⭐ 关键：把正确的 id 写入记录，覆盖/补齐为主键对应的 int 值
            # - 字符串枚举：row_data_key 是序号（也是生成的枚举值）
            # - 组合键：row_data_key 是合成 id
            # - 单列 int：row_data_key 就是主键
            row_data["id"] = int(row_data_key)

            data[row_data_key] = row_data

        file_content = json.dumps(data, ensure_ascii=False, indent=4)
        file_path = f"{output_folder}/{self.name}Config.json"
        write_to_file(file_content, file_path)


    def generate_script(self, output_folder: str) -> None:
        """
        生成 C# 脚本（必要时生成枚举 Key 文件）。
        会把 composite_keys 与 MULTIPLIER 及 composite_key_fields 传给 cs 生成器，
        以便生成的 C# 方法使用真实字段名作为参数名。
        """
        properties_dict = self._get_properties_dict()
        property_remakes = self._get_property_remarks()

        # 传递 composite_key_fields（若启用）给 cs 生成器，以便在 C# 中使用真实字段名（例如 id, group）
        composite_fields = self.composite_key_fields if self.composite_keys else None

        generate_script_file(
            self.name,
            properties_dict,
            property_remakes,
            output_folder,
            self.need_generate_keys,
            composite_keys=self.composite_keys,
            composite_multiplier=self.MULTIPLIER,
            composite_key_fields=composite_fields  # 可能为 None 或 {"key1":"id","key2":"group"}
        )

        # 如果需要字符串枚举键，生成对应的枚举文件
        if self.need_generate_keys:
            self._generate_enum_keys_csfile(output_folder)
