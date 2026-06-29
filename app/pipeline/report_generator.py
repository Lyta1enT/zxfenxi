"""Excel 报表生成器"""
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Optional

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter


REPORT_TYPE_NAMES = {
    'personal': '个人征信报告',
    'corporate': '企业征信报告',
    'tax': '水母报告（税务分析）',
}


SUMMARY_HEADERS = [
    '公司高新/深房',
    '成立/诉讼/变更税等级关联风险',
    '开票纳税',
    '企业征信',
    '法人征信',
]


def _text(value: Any) -> str:
    if value is None:
        return ''
    return str(value).strip()


def _pick_first(fields: Dict[str, Any], keys: List[str]) -> str:
    for key in keys:
        data = fields.get(key, {}) or {}
        value = _text(data.get('value', ''))
        if value and value != '[未识别]':
            return value
    return ''


def _fmt_amount(value: str) -> str:
    if not value:
        return ''
    return value.replace(',', '')


def _status_from_value(value: str) -> str:
    if not value:
        return '未识别'
    if any(token in value for token in ['正常', '无异常', '无']):
        return '正常'
    return value


def _create_styles():
    header_fill = PatternFill('solid', fgColor='D9EAD3')
    sub_fill = PatternFill('solid', fgColor='EAF2E8')
    warn_fill = PatternFill('solid', fgColor='FCE5CD')
    bad_fill = PatternFill('solid', fgColor='F4CCCC')
    title_font = Font(name='Microsoft YaHei', size=14, bold=True)
    header_font = Font(name='Microsoft YaHei', size=11, bold=True, color='000000')
    body_font = Font(name='Microsoft YaHei', size=10)
    thin = Side(style='thin', color='000000')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    return {
        'header_fill': header_fill,
        'sub_fill': sub_fill,
        'warn_fill': warn_fill,
        'bad_fill': bad_fill,
        'title_font': title_font,
        'header_font': header_font,
        'body_font': body_font,
        'border': border,
        'center': center,
        'left': left,
    }


def _set_table_widths(ws, widths: List[int]):
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = width


def _fill_row(ws, row_idx: int, values: List[Any], styles, fill=None, bold=False):
    for col_idx, value in enumerate(values, start=1):
        cell = ws.cell(row=row_idx, column=col_idx, value=value)
        cell.font = styles['header_font'] if bold else styles['body_font']
        cell.border = styles['border']
        cell.alignment = styles['center'] if col_idx != 1 else styles['left']
        if fill:
            cell.fill = fill


# ===== 报告说明过滤关键词（这些行的内容不要出现在输出中）=====
BOILERPLATE_FILTER = [
    '本报告由中国人民银行', '报告说明', '征信中心', '本报告所展示',
    '本报告中信贷交易', '本报告中借贷交易', '如无特别说明',
    '信息主体有权', '如有异议', '信息主体声明',
    '征信中心说明', '数据提供机构说明', '如信息记录斜体',
    '本报告仅展示', '报数机构', '信用报告（自主查询版）',
    'NO.', '报告编号', '查询机构', '中征码',
    '第 ', '页/共', '页', '．本报告',
    # 更多报告规则说明
    '级分类为', '被追偿业务', '逾期总额（含欠息）',
    '逾期天数', '各类贷款', '常见的产品种类',
    '信贷交易包括', '借贷交易包括', '担保交易指',
    '第三人实质上', '资产管理公司处置',
    '透支未超过', '贷记卡账户及',
    '信用卡、贷款和其他信贷记录',
    '金额类数据均以人民币计算',
    '逾期记录可能影响对您的信用评价',
    '按时还最低还款额',
    '请到当地信用报告查询网点',
    '（包括商住两用）',
    '金贷款。', '这部分包含您的',
    '注：', '说明：',
    # 报告规则碎片（含关键词误匹配）
    '提供的担保服务', '保险公司提供的', '信用保证保险',
    '分别对应其中的短期', '分，分别对应其中',
    '五级分类为', '后三类的业务',
    '由信贷交易所产生的债务',
    '逾期本金是指', '除被追偿业务',
    '分类为正常、关注和后三类',
    # 多余标签
    '（自主查询版）', '（人民币账户',
    '（美元账户', '卡片尾号',
]


def _is_boilerplate(line: str) -> bool:
    """判断是否为报告说明/模板文字"""
    return any(kw in line for kw in BOILERPLATE_FILTER)


def _classify_line(line: str) -> str:
    """根据内容将一行文本分到对应的列
    
    Returns: col_name or None
    """
    # 列1: 公司名
    if any(kw in line for kw in ['企业名称', '公司名称', '单位名称']):
        return '公司高新/深房'
    
    # 列2: 成立/诉讼/变更/税务/处罚
    if any(kw in line for kw in [
        '成立', '法人', '变更', '纳税', '税务', 'A级', 'B级', 'M级',
        '滞纳金', '处罚', '罚款', '诉讼', '法院', '行政处罚',
        '经营异常', '违约', '欠税',
        '注册资本', '出资', '经济类型', '企业规模', '所属行业',
        '存续状态', '登记地址',
    ]):
        return '成立/诉讼/变更税等级关联风险'
    
    # 列3: 开票纳税
    if any(kw in line for kw in [
        '开票', '发票', '销售收入', '销售额', '纳税', '缴税', '税款',
    ]):
        return '开票纳税'
    
    # 列4: 企业征信/贷款信息
    if any(kw in line for kw in [
        '贷款', '借款', '余额', '授信', '担保', '抵押', '质押',
        '短期', '中长期', '未结清', '已结清', '信贷',
        '账户数', '余额', '到期', '循环', '万元',
        '发放形式', '担保方式', '开立日期', '到期日',
        '五级分类', '逾期总额', '逾期本金',
        '还款', '正常类', '关注类', '不良类',
        '授信机构', '业务种类', '借款金额',
    ]):
        return '企业征信'
    
    # 列5: 法人征信
    if any(kw in line for kw in [
        '姓名', '证件号码', '身份证', '性别', '出生',
        '信用卡', '贷记卡', '负债', '逾期', '查询',
        '还款责任', '担保责任', '共同还款',
        '额度', '使用率', '已用额度',
    ]):
        return '法人征信'
    
    return None


def _collect_summary_fields(report_type: str, fields: Dict[str, Any],
                            raw_text: str = '') -> Dict[str, Any]:
    """将OCR原始文本按关键词分到5列，过滤报告说明"""
    result = {h: [] for h in SUMMARY_HEADERS}
    
    # 从 raw_text 和 fields 中收集所有文本行
    all_lines = []
    
    # 从 raw_text 分行
    if raw_text:
        for line in raw_text.replace('\\n', '\n').split('\n'):
            line = line.strip()
            if line:
                all_lines.append(line)
    
    # 从 field values 中补充
    if fields:
        for key, data in fields.items():
            val = data.get('value', '') if isinstance(data, dict) else str(data)
            if val and val not in all_lines:
                for line in val.replace('\\n', '\n').split('\n'):
                    line = line.strip()
                    if line and line not in all_lines:
                        all_lines.append(line)
    
    # 过滤报告说明 + 分桶
    for line in all_lines:
        if _is_boilerplate(line):
            continue
        # 过短的碎片行（1-2个字）跳过
        if len(line) <= 2:
            continue
        # 纯数字/符号行跳过
        if line.strip().strip('-—=~').isdigit():
            continue
        col = _classify_line(line)
        if col:
            if line not in result[col]:  # 去重
                result[col].append(line)
    
    # 列1特殊处理：如果有企业名称字段直接用它
    company = _pick_first(fields, ['company_name'])
    if company:
        result['公司高新/深房'] = [company]
    
    # 将列表合并为换行分隔的文本
    return {k: '\n'.join(v) for k, v in result.items()}


def generate_report(fields: Dict[str, Any], report_type: str,
                    source_filename: str, output_dir: str = 'output',
                    raw_text: str = '', layout_result: Optional[Dict] = None) -> str:
    """生成 Excel 报表

    Args:
        fields: 抽取的字段字典
        report_type: 报告类型
        source_filename: 源文件名
        output_dir: 输出目录
        raw_text: 原始文本
        layout_result: 版面分析结果（含表格数据）
    """
    os.makedirs(output_dir, exist_ok=True)

    base_name = Path(source_filename).stem
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'{base_name}_{report_type}_报告_{timestamp}.xlsx')

    wb = Workbook()
    styles = _create_styles()

    # 总表
    ws = wb.active
    ws.title = '总表'
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A2'
    _set_table_widths(ws, [18, 28, 22, 36, 36])

    ws.merge_cells('A1:E1')
    c = ws['A1']
    c.value = f"{REPORT_TYPE_NAMES.get(report_type, '征信报告')} - 结构化总表"
    c.font = styles['title_font']
    c.alignment = styles['center']

    for idx, header in enumerate(SUMMARY_HEADERS, start=1):
        cell = ws.cell(row=2, column=idx, value=header)
        cell.font = styles['header_font']
        cell.fill = styles['header_fill']
        cell.border = styles['border']
        cell.alignment = styles['center']

    summary = _collect_summary_fields(report_type, fields, raw_text)
    row_values = [summary.get(h, '') for h in SUMMARY_HEADERS]
    for idx, value in enumerate(row_values, start=1):
        cell = ws.cell(row=3, column=idx, value=value or '未识别')
        cell.font = styles['body_font']
        cell.border = styles['border']
        cell.alignment = styles['left'] if idx in (1, 2, 4, 5) else styles['center']
        if not value:
            cell.fill = styles['warn_fill']

    # 字段明细
    detail = wb.create_sheet('字段明细')
    detail.sheet_view.showGridLines = False
    detail.freeze_panes = 'A2'
    detail_headers = ['字段分组', '字段键', '字段名称', '字段值', '置信度', '页码', '备注']
    _set_table_widths(detail, [16, 18, 20, 42, 12, 10, 16])
    for idx, header in enumerate(detail_headers, start=1):
        cell = detail.cell(row=1, column=idx, value=header)
        cell.font = styles['header_font']
        cell.fill = styles['header_fill']
        cell.border = styles['border']
        cell.alignment = styles['center']

    field_name_map = {
        'personal': {
            'name': '姓名', 'id_number': '证件号码', 'report_time': '报告时间',
            'credit_card_count': '信用卡账户数', 'loan_count': '贷款账户数',
            'overdue_count': '逾期账户数', 'total_balance': '余额',
            'settled_count': '已结清账户数', 'anomaly_notes': '异常备注',
            'total_debt': '总负债', 'credit_card_usage': '信用卡使用率',
            'personal_loans_raw': '个人贷款明细', 'overdue_history': '逾期历史',
            'repayment_responsibility': '还款责任', 'query_history': '查询次数',
        },
        'corporate': {
            'company_name': '企业名称', 'credit_code': '统一社会信用代码',
            'report_time': '报告时间', 'unsettled_institutions': '未结清机构数',
            'total_balance': '余额', 'short_term_loan': '短期借款',
            'medium_long_term_loan': '中长期借款', 'guarantee_info': '担保信息',
            'public_info': '公共信息',
            'establish_info': '成立信息', 'legal_person': '法定代表人',
            'tax_rating': '税务等级', 'penalty_info': '滞纳金信息',
            'invoice_data': '开票数据', 'tax_data': '纳税数据',
            'loan_details_raw': '贷款明细',
        },
        'tax': {
            'tax_registration': '纳税登记状态', 'has_penalty': '是否有滞纳金',
            'tax_arrears': '欠税金额', 'invoice_3year': '近三年开票汇总',
            'tax_revenue_3year': '近三年纳税数据', 'tax_anomaly': '税务异常说明',
            'invoice_data': '开票数据', 'tax_data': '纳税数据',
        },
    }

    row = 2
    group_label = REPORT_TYPE_NAMES.get(report_type, report_type)
    for key, data in fields.items():
        label = field_name_map.get(report_type, {}).get(key, key)
        value = _text(data.get('value', '')) if isinstance(data, dict) else _text(data)
        if not value:
            value = '未识别'
        detail_values = [
            group_label,
            key,
            label,
            value,
            data.get('confidence', '') if isinstance(data, dict) else '',
            data.get('page', '') if isinstance(data, dict) else '',
            data.get('note', '') if isinstance(data, dict) else '',
        ]
        for col, v in enumerate(detail_values, start=1):
            cell = detail.cell(row=row, column=col, value=v)
            cell.font = styles['body_font']
            cell.border = styles['border']
            cell.alignment = styles['left'] if col in (1, 2, 3, 4, 7) else styles['center']
            if col == 4 and v == '未识别':
                cell.fill = styles['warn_fill']
        row += 1

    # 原始文本
    raw_sheet = wb.create_sheet('原始文本')
    raw_sheet.sheet_view.showGridLines = False
    raw_sheet.freeze_panes = 'A1'
    raw_sheet.column_dimensions['A'].width = 140
    raw_cell = raw_sheet['A1']
    raw_cell.value = raw_text or '（无原始文本）'
    raw_cell.alignment = Alignment(wrap_text=True, vertical='top')
    raw_cell.font = styles['body_font']

    # 表格数据（来自版面分析）
    if layout_result and layout_result.get('tables'):
        table_sheet = wb.create_sheet('表格数据')
        table_sheet.sheet_view.showGridLines = False
        _set_table_widths(table_sheet, [10, 50, 10, 10, 60])

        tbl_headers = ['表格#', '类型', '页', '行', '内容']
        for idx, h in enumerate(tbl_headers, start=1):
            cell = table_sheet.cell(row=1, column=idx, value=h)
            cell.font = styles['header_font']
            cell.fill = styles['header_fill']
            cell.border = styles['border']
            cell.alignment = styles['center']

        row = 2
        for ti, table in enumerate(layout_result.get('tables', [])):
            cells = table.get('cells', [])
            if not cells:
                # 无 cells 时用 HTML 或原文本
                html = table.get('html', '')
                table_sheet.cell(row=row, column=1, value=ti + 1).font = styles['body_font']
                table_sheet.cell(row=row, column=2, value='HTML').font = styles['body_font']
                table_sheet.cell(row=row, column=5, value=html[:2000] if html else str(table)).font = styles['body_font']
                for c in range(1, 6):
                    table_sheet.cell(row=row, column=c).border = styles['border']
                row += 1
                continue

            # 遍历表格的行列
            for ri, cell_row in enumerate(cells):
                if isinstance(cell_row, list):
                    row_text = ' | '.join([
                        str(c.get('text', '')) if isinstance(c, dict) else str(c)
                        for c in cell_row
                    ])
                    row_type = 'table_header' if ri == 0 else 'table_data'
                elif isinstance(cell_row, dict):
                    row_text = str(cell_row.get('text', ''))
                    row_type = 'table_cell'
                else:
                    row_text = str(cell_row)
                    row_type = 'table_data'

                table_sheet.cell(row=row, column=1, value=ti + 1).font = styles['body_font']
                table_sheet.cell(row=row, column=2, value=row_type).font = styles['body_font']
                table_sheet.cell(row=row, column=3, value=table.get('page', '')).font = styles['body_font']
                table_sheet.cell(row=row, column=4, value=ri + 1).font = styles['body_font']
                table_sheet.cell(row=row, column=5, value=row_text[:2000]).font = styles['body_font']
                for c in range(1, 6):
                    table_sheet.cell(row=row, column=c).border = styles['border']
                    if row_type == 'table_header':
                        table_sheet.cell(row=row, column=c).font = Font(
                            name='Microsoft YaHei', size=10, bold=True)
                row += 1

            # 空一行分隔不同表格
            row += 1

    wb.save(output_path)
    return output_path
