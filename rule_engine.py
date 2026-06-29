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
                       skip_hidden: bool = False,
                       header_label: str = "",
                       header_row: int = 0) -> Dict[str, int]:
    """
    构建目标文件的「键→行号」索引。

    Args:
        target_path: 目标文件路径
        target_key_col: 目标中匹配键的列索引 (0-based)
        target_sheet: 目标 Sheet 名（默认第一个 Sheet）
        skip_hidden: 是否跳过隐藏行
        header_label: 用户标记的目标键列标签，用于定位表头行

    Returns:
        {key_value: row_number (1-based), ...}
    """
    data = read_excel(target_path, skip_hidden=skip_hidden)

    if target_sheet is None:
        target_sheet = list(data.keys())[0]

    rows = data.get(target_sheet, [])
    index: Dict[str, int] = {}

    # ── 用标签精确定位目标文件的表头行（覆盖自动检测值）──
    if header_label:
        limit = min(5, len(rows))
        for i in range(limit):
            if target_key_col < len(rows[i]):
                val = str(rows[i][target_key_col]).strip()
                if val == header_label:
                    header_row = i
                    break

    for row_idx, row_data in enumerate(rows):
        if row_idx == header_row:
            continue  # 跳过表头行，不纳入索引
        if target_key_col < len(row_data):
            key = str(row_data[target_key_col]).strip()
            if key:
                index[key] = row_idx + 1  # 1-based row number

    return index


def preview_matches(source_dir: str, target_path: str,
                    mapping: ColumnMapping,
                    source_sheet: str = None,
                    target_sheet: str = None,
                    skip_hidden: bool = False,
                    target_header_row: int = 0) -> PreviewResult:
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

    # 构建目标索引（传入标签+自动检测的表头行）
    target_index = build_target_index(target_path, mapping.target_key_col,
                                      target_sheet, skip_hidden=skip_hidden,
                                      header_label=mapping.target_key_label,
                                      header_row=target_header_row)

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

        # ── 按标签定位源文件的列（回退到固定索引）──
        src_key_col = mapping.source_key_col
        src_val_col = mapping.source_value_col

        # 检测当前源文件的表头行（标签驱动，关键词兜底）
        src_header_rows = _find_header_rows(
            rows,
            mapping.source_key_label,
            mapping.source_value_label,
            target_header_keys,
            src_key_col,
            src_val_col,
        )

        # 按标签定位列
        if mapping.source_key_label:
            src_key_col = _resolve_column(
                rows, mapping.source_key_label, src_header_rows, src_key_col
            )
        if mapping.source_value_label:
            src_val_col = _resolve_column(
                rows, mapping.source_value_label, src_header_rows, src_val_col
            )

        # 标签找不到 → 该文件列结构不匹配，跳过并警告
        if src_key_col < 0 or src_val_col < 0:
            missing = []
            if src_key_col < 0:
                missing.append("匹配键\"{}\"".format(mapping.source_key_label))
            if src_val_col < 0:
                missing.append("复制列\"{}\"".format(mapping.source_value_label))
            result.unmatched_keys.append(
                "[列不匹配] {}: 未找到{}，已跳过".format(rel_path, "、".join(missing))
            )
            continue

        for row_idx, row_data in enumerate(rows):
            # 跳过检测到的表头行
            if row_idx in src_header_rows:
                continue

            # 读取匹配键
            key = ""
            if src_key_col < len(row_data):
                key = str(row_data[src_key_col]).strip()

            if not key:
                continue  # 跳过空键行

            # 精确表头关键词过滤（安全网：跳过恰好命中 header_keys 的值）
            if _is_header_key(key, target_header_keys):
                continue

            result.total_source_rows += 1

            # 读取要复制的值
            value = ""
            if src_val_col < len(row_data):
                value = str(row_data[src_val_col]).strip()

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
# 涵盖中文办公场景中几乎所有的列标题，持续补充
HEADER_KEYWORDS = {
    # ── 人员信息 ──
    '序号', '编号', '代码', '学号', '工号', 'ID', 'No', '编号/序号',
    '姓名', '名称', '名字', '称呼', '联系人', '负责人',
    '性别', '年龄', '民族', '籍贯', '国籍', '户口', '户籍',
    '出生日期', '出身日期', '出生年月', '生日',
    '身份证号', '身份证', '证件号', '证件号码', '证件类型',
    '电话', '手机', '手机号', '联系电话', '联系方式', '座机', '传真',
    '邮箱', '电子邮件', 'E-mail', 'Email', 'QQ', '微信', '微信号',
    '地址', '住址', '通讯地址', '家庭住址', '单位地址', '邮编', '邮政编码',
    '学历', '学位', '专业', '毕业院校', '毕业学校', '毕业时间',
    '政治面貌', '婚姻状况', '健康状况',
    # ── 工作/单位 ──
    '工作单位', '单位', '单位名称', '公司', '公司名称', '机构',
    '部门', '科室', '班组', '车间',
    '职务', '职位', '职称', '岗位', '工种',
    '入职日期', '入职时间', '参加工作时间', '工龄',
    # ── 财务/数值 ──
    '金额', '总金额', '合计金额', '金额(元)', '金额（元）',
    '数量', '份数', '件数', '个数',
    '单价', '总价', '价格', '费用', '支出', '收入',
    '工资', '薪资', '报酬', '补贴', '津贴',
    '税费', '税率', '税额',
    # ── 日期/时间 ──
    '日期', '时间', '年月', '年份', '年度', '月份', '月份/日期',
    '开始日期', '结束日期', '起止日期', '截止日期',
    '登记日期', '记录日期', '填报日期', '创建日期', '更新日期',
    # ── 状态/分类 ──
    '状态', '进度', '完成情况', '办理情况', '处理状态',
    '类别', '类型', '分类', '项目', '项目名称',
    '品牌', '规格', '型号', '同型号', '版本',
    '等级', '级别', '星级',
    # ── 备注/说明 ──
    '备注', '说明', '描述', '内容', '摘要', '简介', '详细信息',
    '补充说明', '注意事项', '其他', '其它',
    # ── 走访/调查相关 ──
    '走访人', '走访情况', '走访简要情况', '走访记录', '走访内容',
    '走访日期', '走访时间', '走访对象',
    '调查人', '调查日期', '调查情况', '调查结果',
    '记录人', '审核人', '批准人', '经办人', '签收人',
    # ── 文件/文档 ──
    '文件名', '文件编号', '文号', '档案号',
    '标题', '主题', '关键词',
    # ── 其他常见 ──
    '是否', '是否有效', '是否通过', '是否完成',
    '得分', '评分', '成绩', '分数',
    '排名', '名次',
    '网址', '链接', 'URL',
    'IP地址', 'MAC地址',
    '经度', '纬度', '坐标',
    '照片', '图片', '附件',
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
    """判断键值是否为表头（精确匹配 header_keys 集合）"""
    if not key:
        return False
    return key in header_keys


def _find_header_rows(rows: List[List], source_key_label: str,
                      source_value_label: str, header_keys: set,
                      key_col: int, val_col: int) -> set:
    """
    在源文件中检测表头行（纯标签驱动，关键词仅作兜底）。

    策略：
    1. 优先用已知列标签（如 "姓名", "电话"）在前 8 行搜索
    2. 包含任意标签的行 → 判定为表头行
    3. 若标签为空，回退到 header_keywords 关键词匹配
    4. 都匹配不到 → 默认行 0 为表头行

    Returns:
        表头行索引集合 (0-based)
    """
    header_rows = set()
    labels = [l for l in (source_key_label, source_value_label) if l]

    if labels:
        # ── 标签优先：在前 8 行中搜索包含标签的行 ──
        limit = min(8, len(rows))
        for i in range(limit):
            row_values = [str(c).strip() for c in rows[i]]
            if any(label in row_values for label in labels):
                header_rows.add(i)
    else:
        # ── 兜底：无标签时用关键词匹配 ──
        limit = min(8, len(rows))
        for i in range(limit):
            for cell in rows[i]:
                cell_str = str(cell).strip()
                if cell_str and cell_str in header_keys:
                    header_rows.add(i)
                    break

    # ── 最终兜底：什么都没找到，默认行 0 ──
    if not header_rows:
        header_rows.add(0)

    return header_rows


def _resolve_column(rows: List[List], label: str,
                    header_rows: set, fallback_col: int) -> int:
    """
    按标签在表头行中定位列索引。

    在检测到的表头行中搜索与 label 匹配的单元格，
    返回其列索引。找不到则返回 -1（而非静默回退到固定索引），
    让调用方决定是否跳过该文件。

    Args:
        rows: 源文件行数据
        label: 要匹配的列标签（用户点击时捕获的单元格值或表头值）
        header_rows: 表头行索引集合
        fallback_col: 回退列索引（仅当 label 为空时使用）

    Returns:
        解析后的列索引 (0-based)，或 -1 表示未找到
    """
    if not label:
        return fallback_col
    # 在表头行中精确搜索
    for row_idx in sorted(header_rows):
        row = rows[row_idx] if row_idx < len(rows) else []
        for col_idx, cell in enumerate(row):
            cell_str = str(cell).strip()
            if cell_str == label:
                return col_idx
    # 前缀匹配：标签是单元格值的前缀（如 "证件号" 匹配 "证件号码"）
    # 但 "证件号" 不匹配 "签证证件号码"（不是前缀）
    for row_idx in sorted(header_rows):
        row = rows[row_idx] if row_idx < len(rows) else []
        for col_idx, cell in enumerate(row):
            cell_str = str(cell).strip()
            if cell_str and cell_str.startswith(label):
                return col_idx
    # 在首行（兜底表头行）中同样搜索
    if 0 not in header_rows and 0 < len(rows):
        for col_idx, cell in enumerate(rows[0]):
            cell_str = str(cell).strip()
            if cell_str == label or (cell_str and cell_str.startswith(label)):
                return col_idx
    # 标签有意义但找不到 → 返回 -1，让调用方跳过该文件
    return -1


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
