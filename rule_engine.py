"""
规则引擎：映射规则管理 + 预览匹配引擎
"""

import os
import json
from typing import List, Dict, Optional, Any, Tuple
from dataclasses import dataclass, asdict, field

from excel_reader import read_excel, find_files, get_relative_path


@dataclass
class ColumnMapping:
    """列映射规则"""
    source_value_col: int = -1      # 源文件「要复制的值」列索引 (0-based)
    source_key_col: int = -1        # 源文件「匹配键」列索引 (0-based)
    target_key_col: int = -1        # 目标文件「匹配键」列索引 (0-based)
    target_dest_col: int = -1       # 目标文件「粘贴到」列索引 (0-based)

    source_value_label: str = ""    # 源值列的标签（用户选择的单元格内容）
    source_key_label: str = ""      # 源键列的标签
    target_key_label: str = ""      # 目标键列的标签
    target_dest_label: str = ""     # 目标粘贴列的标签

    def is_source_ready(self) -> bool:
        return self.source_value_col >= 0 and self.source_key_col >= 0

    def is_target_ready(self) -> bool:
        return self.target_key_col >= 0 and self.target_dest_col >= 0

    def is_ready(self) -> bool:
        return self.is_source_ready() and self.is_target_ready()


@dataclass
class MatchEntry:
    """单条匹配结果（用于预览）"""
    key: str = ""                    # 匹配键的值（如 A001）
    value: str = ""                  # 要复制的值（如 "2月3日来么"）
    source_file: str = ""            # 来源文件相对路径
    source_sheet: str = ""           # 来源 Sheet 名
    source_row: int = 0              # 来源行号 (1-based)
    target_row: int = 0              # 目标行号 (1-based, 0 表示未找到匹配)
    has_data: bool = False           # 是否有数据（value非空）
    selected: bool = True            # 用户是否勾选此项


@dataclass
class PreviewResult:
    """预览结果"""
    entries: List[MatchEntry] = field(default_factory=list)
    unmatched_keys: List[str] = field(default_factory=list)  # 在目标中找不到匹配的键
    total_source_rows: int = 0       # 源文件总扫描行数
    matched_count: int = 0           # 匹配成功的行数
    no_data_count: int = 0           # 匹配成功但无数据的行数


def build_target_index(target_path: str, target_key_col: int,
                       target_sheet: str = None,
                       skip_hidden: bool = False) -> Dict[str, int]:
    """
    构建目标文件的「键→行号」索引。

    Args:
        target_path: 目标文件路径
        target_key_col: 目标中匹配键的列索引 (0-based)
        target_sheet: 目标 Sheet 名（默认第一个 Sheet）
        skip_hidden: 是否跳过隐藏行

    Returns:
        {key_value: row_number (1-based), ...}
    """
    data = read_excel(target_path, skip_hidden=skip_hidden)

    if target_sheet is None:
        target_sheet = list(data.keys())[0]

    rows = data.get(target_sheet, [])
    index: Dict[str, int] = {}

    for row_idx, row_data in enumerate(rows):
        if target_key_col < len(row_data):
            key = str(row_data[target_key_col]).strip()
            if key:
                index[key] = row_idx + 1  # 1-based row number

    return index


def preview_matches(source_dir: str, target_path: str,
                    mapping: ColumnMapping,
                    source_sheet: str = None,
                    target_sheet: str = None,
                    skip_hidden: bool = False) -> PreviewResult:
    """
    扫描所有源文件，生成预览匹配结果。

    Args:
        source_dir: 源文件目录
        target_path: 目标文件路径
        mapping: 列映射规则
        source_sheet: 源文件 Sheet 名（默认第一个 Sheet，适用于所有源文件）
        target_sheet: 目标 Sheet 名（默认第一个 Sheet）
        skip_hidden: 是否跳过隐藏行

    Returns:
        PreviewResult 包含所有匹配条目
    """
    result = PreviewResult()
    source_files = find_files(source_dir, recursive=True)

    # 排除目标文件自身
    if os.path.isfile(target_path):
        target_abs = os.path.abspath(target_path)
        source_files = [f for f in source_files if os.path.abspath(f) != target_abs]
    else:
        # 目标不存在，返回空结果
        result.unmatched_keys.append("[目标文件不存在] {}".format(target_path))
        return result

    # 构建目标索引（同时收集表头值用于过滤）
    target_index = build_target_index(target_path, mapping.target_key_col,
                                      target_sheet, skip_hidden=skip_hidden)

    # 识别目标中的表头值（目标文件中匹配键列的表头不应用作匹配键）
    # 读取目标文件的 header 值
    target_header_keys = _detect_header_keys(
        target_path, mapping.target_key_col, target_sheet
    )

    for src_path in source_files:
        try:
            data = read_excel(src_path, skip_hidden=skip_hidden)
        except Exception as e:
            result.unmatched_keys.append("[读取失败] {}: {}".format(
                get_relative_path(src_path, source_dir), str(e)))
            continue

        # 确定 Sheet
        if source_sheet and source_sheet in data:
            sheet = source_sheet
        else:
            sheet = list(data.keys())[0]

        rows = data.get(sheet, [])
        rel_path = get_relative_path(src_path, source_dir)

        for row_idx, row_data in enumerate(rows):
            # 读取匹配键
            key = ""
            if mapping.source_key_col < len(row_data):
                key = str(row_data[mapping.source_key_col]).strip()

            if not key:
                continue  # 跳过空键行

            # 跳过表头行（键值为常见表头文本）
            if _is_header_key(key, target_header_keys):
                continue

            result.total_source_rows += 1

            # 读取要复制的值
            value = ""
            if mapping.source_value_col < len(row_data):
                value = str(row_data[mapping.source_value_col]).strip()

            entry = MatchEntry(
                key=key,
                value=value,
                source_file=rel_path,
                source_sheet=sheet,
                source_row=row_idx + 1,
                target_row=target_index.get(key, 0),
                has_data=bool(value),
                selected=bool(value),  # 有数据的默认勾选，无数据的默认不选
            )

            if entry.target_row == 0:
                result.unmatched_keys.append(key)
            elif entry.has_data:
                result.matched_count += 1
            else:
                result.no_data_count += 1

            result.entries.append(entry)

    return result


# 常见表头关键词（用于自动识别表头行）
HEADER_KEYWORDS = {
    '序号', '姓名', '名称', '编号', '日期', '时间',
    '证件号', '身份证号', '身份证', '电话', '手机',
    '地址', '性别', '出身日期', '出生日期', '出生年月',
    '同型号', '型号', '走访人', '走访简要情况', '走访情况',
    '备注', '说明', '金额', '数量', '单价', '状态',
    '项目', '类别', '类型', '部门', '单位',
}


def _detect_header_keys(target_path: str, key_col: int,
                        target_sheet: Optional[str] = None) -> set:
    """
    从目标文件中检测表头键值。

    只添加「看起来像表头文本」的值（中文、纯字母等），
    排除「看起来像数据」的值（含数字的编号如 A001、日期等）。

    Returns:
        set of header key values
    """
    header_keys = set(HEADER_KEYWORDS)
    try:
        data = read_excel(target_path, skip_hidden=False)  # 表头检测不需要跳过隐藏
        if target_sheet is None:
            target_sheet = list(data.keys())[0]
        rows = data.get(target_sheet, [])

        for row_idx in range(min(10, len(rows))):
            if key_col < len(rows[row_idx]):
                val = str(rows[row_idx][key_col]).strip()
                if val and _looks_like_header(val):
                    header_keys.add(val)
    except Exception:
        pass
    return header_keys


def _looks_like_header(val: str) -> bool:
    """判断一个值是否「看起来像表头文本」而非数据"""
    if not val:
        return False
    # 包含中文 → 很可能是表头
    if any('一' <= c <= '鿿' for c in val):
        return True
    # 纯字母且长度 >= 2 → 可能是表头（如 ID, Name, Code）
    if val.isalpha() and len(val) >= 2:
        return True
    # 包含数字 → 很可能是数据（如 A001, 2024-01-01, 123）
    if any(c.isdigit() for c in val):
        return False
    return False


def _is_header_key(key: str, header_keys: set) -> bool:
    """判断键值是否为表头（需要同时满足：在header_keys中且看起来像表头）"""
    if not key:
        return False
    if key in header_keys:
        return True
    # 额外安全网：看起来像表头的值
    if _looks_like_header(key):
        return True
    return False


def save_rule(filepath: str, mapping: ColumnMapping,
              source_dir: str = "", target_path: str = ""):
    """保存映射规则到 JSON 文件"""
    rule = {
        'mapping': asdict(mapping),
        'source_dir': source_dir,
        'target_path': target_path,
    }
    with open(filepath, 'w', encoding='utf-8') as f:
        json.dump(rule, f, ensure_ascii=False, indent=2)


def load_rule(filepath: str) -> Tuple[ColumnMapping, str, str]:
    """从 JSON 文件加载映射规则，返回 (mapping, source_dir, target_path)"""
    with open(filepath, 'r', encoding='utf-8') as f:
        rule = json.load(f)

    mapping_data = rule.get('mapping', {})
    mapping = ColumnMapping(
        source_value_col=mapping_data.get('source_value_col', -1),
        source_key_col=mapping_data.get('source_key_col', -1),
        target_key_col=mapping_data.get('target_key_col', -1),
        target_dest_col=mapping_data.get('target_dest_col', -1),
        source_value_label=mapping_data.get('source_value_label', ''),
        source_key_label=mapping_data.get('source_key_label', ''),
        target_key_label=mapping_data.get('target_key_label', ''),
        target_dest_label=mapping_data.get('target_dest_label', ''),
    )
    return mapping, rule.get('source_dir', ''), rule.get('target_path', '')
