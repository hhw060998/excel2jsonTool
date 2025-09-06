# Author: huhongwei 306463233@qq.com
# MIT License
import json
import re
from typing import Any, Dict, List
from collections import defaultdict

from cs_generation import generate_script_file, generate_enum_file
from excel_processing import read_cell_values, check_repeating_values
from data_processing import convert_to_type, available_csharp_enum_name
from log import log_warn, log_error, log_info
from exceptions import (
    InvalidEnumNameError,
    DuplicatePrimaryKeyError,
    CompositeKeyOverflowError,
)
from naming_config import JSON_FILE_PATTERN, ENUM_KEYS_SUFFIX, JSON_SORT_KEYS, JSON_ID_FIRST

# 新增：可选字段统计开关（保持功能不变，默认打印一次汇总；若不需要可改为 False）
_PRINT_FIELD_SUMMARY = True


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
        # 一致性：行长度校验（字段行与类型行长度不一致提前报错）
        # 读取 1..6 行
        self.cell_values: Dict[int, List[Any]] = {
            i: read_cell_values(worksheet, i) for i in range(1, 7)
        }
        self.remarks = self.cell_values[1]
        self.headers = self.cell_values[2]
        self.data_types = self.cell_values[3]
        self.data_labels = self.cell_values[4]
        self.field_names = self.cell_values[5]
        self.default_values = self.cell_values[6]

        if len(self.data_types) != len(self.field_names):
            raise RuntimeError(
                f"字段行数量与类型行数量不一致: fields={len(self.field_names)}, types={len(self.data_types)}"
            )
        if len(self.data_labels) != len(self.field_names):
            raise RuntimeError(
                f"字段行数量与标签行数量不一致: fields={len(self.field_names)}, labels={len(self.data_labels)}"
            )
        if len(self.default_values) != len(self.field_names):
            # 不强制报错，只警告
            log_warn(
                f"默认值列数量与字段列数量不一致: fields={len(self.field_names)}, defaults={len(self.default_values)}"
            )

        # 数据行
        self.row_data = list(worksheet.iter_rows(min_row=7, min_col=2))

        # 重复字段检测
        check_repeating_values(self.field_names)

        # 统计(仅用于汇总日志）
        self._field_total = len(self.field_names) - 1 if len(self.field_names) > 0 else 0  # 去掉首列主键列
        self._ignore_count = sum(1 for i in range(len(self.field_names)) if self.data_labels[i] == "ignore")
        self._required_fields = {i for i in range(len(self.field_names)) if self.data_labels[i] == "required"}

        # 检测是否存在有效数据（至少一行中出现非空且未 ignore 的单元格）
        self._has_effective_data = self._check_has_effective_data()

        self.need_generate_keys = self._need_generate_keys()
        self.composite_keys = False
        self.composite_key_fields: Dict[str, str] = {}
        self._detect_composite_keys_with_prefixes_in_first_two_columns()
        if self.need_generate_keys:
            self._check_duplicate_enum_keys()
        if self.composite_keys:
            self._check_duplicate_composite_keys()
        self.first_int_pk_not_named_id_warned = False

        if not self._has_effective_data:
            log_warn(f"表[{self.name}] 没有有效数据行（将生成空 JSON）。")

        if _PRINT_FIELD_SUMMARY:
            log_info(
                f"[{self.name}] 字段统计: 总列={len(self.field_names)} 可用列(含主键)={len(self.field_names)} "
                f"ignore列={self._ignore_count} required列={len(self._required_fields)}"
            )

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
            raise InvalidEnumNameError(name, excel_row)

    def _check_duplicate_enum_keys(self) -> None:
        """
        初始化时检查用于生成枚举的首列（字符串主键）：
        - 验证每个名字是否合法
        - 收集出现的 Excel 行号，若重复则抛错（显示真实 Excel 行号）
        """
        name_rows = defaultdict(list)
        for idx, row in enumerate(self.row_data):
            if not row:
                continue
            val = row[0].value
            excel_row = 7 + idx
            self._validate_enum_name(val, excel_row)
            name_rows[val].append(excel_row)
        dup = {k: v for k, v in name_rows.items() if len(v) > 1}
        if dup:
            lines = "; ".join(f"{k} -> 行{v}" for k, v in dup.items())
            raise InvalidEnumNameError(f"重复的字符串主键: {lines}", -1)

    def _check_duplicate_composite_keys(self) -> None:
        """
        初始化时检查组合 int 键（要求位于数据前两列 -> row[0], row[1] 且用 key1:real / key2:real 标记）：
        - 检查 key1/key2 是否为整数、是否在允许范围内
        - 检查组合后的 combined 是否唯一（若重复则抛错并显示真实 Excel 行号，及对应实际 (key1,key2)）
        """
        seen = {}
        for idx, row in enumerate(self.row_data):
            if len(row) < 2:
                continue
            k1 = row[0].value
            k2 = row[1].value
            excel_row = 7 + idx
            if k1 is None or k2 is None:
                raise RuntimeError(f"行{excel_row} key1/key2 为空")
            try:
                i1 = int(k1); i2 = int(k2)
            except Exception:
                raise RuntimeError(f"行{excel_row} key1/key2 不是整数: {k1},{k2}")
            combined = i1 * self.MULTIPLIER + i2
            if combined in seen:
                raise DuplicatePrimaryKeyError(combined, seen[combined], excel_row)
            seen[combined] = excel_row

    def _generate_enum_keys_csfile(self, output_folder: str) -> None:
        """当需要 string 枚举键时才调用（保留原有实现）"""
        enum_type_name = f"{self.name}{ENUM_KEYS_SUFFIX}"
        enum_names = []
        enum_values = []
        idx_val = 0
        for idx, row in enumerate(self.row_data):
            if not row:
                continue
            val = row[0].value
            self._validate_enum_name(val, 7 + idx)
            enum_names.append(val)
            enum_values.append(idx_val)
            idx_val += 1
        generate_enum_file(enum_type_name, enum_names, enum_values, None, "Data.TableScript", output_folder)

    def _check_has_effective_data(self) -> bool:
        """
        检查是否至少存在一行包含至少一个非 ignore 且非空的单元格。
        不改变现有生成逻辑，仅用于日志提示。
        """
        if not self.row_data:
            return False
        for row in self.row_data:
            for col_index, cell in enumerate(row, start=1):
                if col_index >= len(self.field_names):
                    continue
                if self.data_labels[col_index] == "ignore":
                    continue
                if cell.value not in (None, ""):
                    return True
        return False

    def generate_json(self, output_folder: str) -> None:
        """将表格数据导出为 JSON 文件（支持单列 int 主键 / 自动生成键 / 以及组合键）。
        这里同时确保给每条记录填充 info.id：
            - 字符串主键（枚举）：id = 序号（枚举 int 值）
            - 组合键：id = key1*MULTIPLIER + key2
            - 单列 int 主键：id = 第一列的 int 值；若第一列字段名不是 'id' 则打印一次警告
        """
        data: Dict[Any, Dict[str, Any]] = {}
        serial_key = 0
        first_real = self._actual_field_name(1) if len(self.field_names) > 1 else None
        used_keys = {}

        # 新增：统计 required 缺失次数（虽然缺失会抛错，此计数主要用于未来扩展；保持现功能）
        required_missing_count = 0

        for row_idx, row in enumerate(self.row_data):
            if not row:
                continue
            excel_row = 7 + row_idx
            # 处理主键
            if self.need_generate_keys:
                row_key = serial_key
                serial_key += 1
            elif self.composite_keys:
                try:
                    k1 = int(row[0].value)
                    k2 = int(row[1].value)
                except Exception:
                    raise RuntimeError(f"行{excel_row} 无法解析组合键 int")
                if not (0 <= k1 < self.MAX_KEY2 and 0 <= k2 < self.MAX_KEY2):
                    raise RuntimeError(f"行{excel_row} 组合键超范围 0~{self.MAX_KEY2-1}")
                row_key = k1 * self.MULTIPLIER + k2
                if row_key >= 2**31:
                    raise CompositeKeyOverflowError(row_key)
            else:
                try:
                    row_key = int(row[0].value)
                except Exception:
                    raise RuntimeError(f"行{excel_row} 主键非 int: {row[0].value}")
                if (isinstance(first_real, str)
                        and first_real.lower() != "id"
                        and not self.first_int_pk_not_named_id_warned):
                    log_warn(f"表[{self.name}] 第一列视为主键但字段名不是 id，已写入 id 属性。建议修改表头。")
                    self.first_int_pk_not_named_id_warned = True

            if row_key in used_keys:
                raise DuplicatePrimaryKeyError(row_key, used_keys[row_key], excel_row)
            used_keys[row_key] = excel_row

            # 保持列顺序：按 Excel 顺序构建
            if JSON_ID_FIRST:
                row_obj = {"id": int(row_key)}
            else:
                row_obj = {}

            for col_index, cell in enumerate(row, start=1):
                if col_index >= len(self.field_names):
                    continue
                if self.data_labels[col_index] == "ignore":
                    continue
                data_name = self._actual_field_name(col_index)
                type_str = self.data_types[col_index]
                default_value = self.default_values[col_index]
                cell_value = cell.value

                # 截断原始值用于日志
                def _short(v, limit=60):
                    s = str(v)
                    return s if len(s) <= limit else s[:limit] + "..."

                if cell_value is None:
                    if default_value is None and self.data_labels[col_index] == "required":
                        required_missing_count += 1
                        raise RuntimeError(f"{data_name} required 但值为空且无默认值 (行{excel_row})")
                    try:
                        value = convert_to_type(type_str, default_value)
                    except Exception as e:
                        log_error(f"{self.name} 行{excel_row} 字段 {data_name} 默认值转换失败 原值[{_short(default_value)}]: {e}")
                        value = None
                else:
                    try:
                        value = convert_to_type(type_str, cell_value)
                    except Exception as e:
                        log_error(f"{self.name} 行{excel_row} 字段 {data_name} 转换失败 原值[{_short(cell_value)}]: {e}")
                        value = None
                row_obj[data_name] = value

            if not JSON_ID_FIRST:
                row_obj["id"] = int(row_key)
            data[row_key] = row_obj

        file_content = json.dumps(
            data,
            ensure_ascii=False,
            indent=4,
            sort_keys=JSON_SORT_KEYS
        )
        import os
        from cs_generation import write_to_file
        file_path = os.path.join(output_folder, JSON_FILE_PATTERN.format(name=self.name))
        write_to_file(file_content, file_path)

        if _PRINT_FIELD_SUMMARY:
            log_info(f"[{self.name}] 导出完成: 行数={len(data)} required缺失={required_missing_count}")

    def generate_script(self, output_folder: str) -> None:
        """
        生成 C# 脚本（必要时生成枚举 Key 文件）。
        会把 composite_keys 与 MULTIPLIER 及 composite_key_fields 传给 cs 生成器，
        以便生成的 C# 方法使用真实字段名作为参数名。
        """
        props = self._get_properties_dict()
        remarks = self._get_property_remarks()
        generate_script_file(
            self.name,
            props,
            remarks,
            output_folder,
            self.need_generate_keys,
            composite_keys=self.composite_keys,
            composite_multiplier=self.MULTIPLIER,
            composite_key_fields=self.composite_key_fields if self.composite_keys else None
        )
        if self.need_generate_keys:
            self._generate_enum_keys_csfile(output_folder)
