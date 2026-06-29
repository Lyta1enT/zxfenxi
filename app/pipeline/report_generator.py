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


def _collect_summary_fields(report_type: str, fields: Dict[str, Any], raw_text: str = '') -> Dict[str, Any]:
    result = {h: '' for h in SUMMARY_HEADERS}

    if report_type == 'corporate':
        company_name = _pick_first(fields, ['company_name'])
        credit_code = _pick_first(fields, ['credit_code'])
        report_time = _pick_first(fields, ['report_time'])
        total_balance = _fmt_amount(_pick_first(fields, ['total_balance']))
        unsettled = _pick_first(fields, ['unsettled_institutions'])
        short_term = _fmt_amount(_pick_first(fields, ['short_term_loan']))
        medium_long = _fmt_amount(_pick_first(fields, ['medium_long_term_loan']))
        public_info = _pick_first(fields, ['public_info'])
        guarantee = _pick_first(fields, ['guarantee_info'])

        result['公司高新/深房'] = ''
        risk_bits = []
        for token in ['成立', '诉讼', '变更', '行政处罚', '法院', '经营异常', '纳税', '欠税', '税务', '违约']:
            if token in raw_text:
                risk_bits.append(token)
        result['成立/诉讼/变更税等级关联风险'] = '、'.join(dict.fromkeys(risk_bits))
        result['开票纳税'] = _pick_first(fields, ['tax_registration', 'has_penalty', 'tax_arrears', 'invoice_3year', 'tax_revenue_3year']) or (
            '、'.join([token for token in ['开票', '纳税', '欠税'] if token in raw_text])
        )
        result['企业征信'] = '; '.join([x for x in [
            f'企业名称:{company_name}' if company_name else '',
            f'统一社会信用代码:{credit_code}' if credit_code else '',
            f'报告时间:{report_time}' if report_time else '',
            f'未结清机构数:{unsettled}' if unsettled else '',
            f'余额:{total_balance}' if total_balance else '',
            f'短期借款:{short_term}' if short_term else '',
            f'中长期借款:{medium_long}' if medium_long else '',
        ] if x])
        result['法人征信'] = _pick_first(fields, ['guarantee_info', 'public_info']) or (
            '、'.join([token for token in ['法人', '责任', '担保'] if token in raw_text])
        )
        return result

    if report_type == 'personal':
        result['企业征信'] = '; '.join([x for x in [
            f'姓名:{_pick_first(fields, ["name"])}' if _pick_first(fields, ["name"]) else '',
            f'证件号码:{_pick_first(fields, ["id_number"])}' if _pick_first(fields, ["id_number"]) else '',
            f'报告时间:{_pick_first(fields, ["report_time"])}' if _pick_first(fields, ["report_time"]) else '',
            f'信用卡账户数:{_pick_first(fields, ["credit_card_count"])}' if _pick_first(fields, ["credit_card_count"]) else '',
            f'贷款账户数:{_pick_first(fields, ["loan_count"])}' if _pick_first(fields, ["loan_count"]) else '',
            f'逾期账户数:{_pick_first(fields, ["overdue_count"])}' if _pick_first(fields, ["overdue_count"]) else '',
            f'余额:{_pick_first(fields, ["total_balance"])}' if _pick_first(fields, ["total_balance"]) else '',
            f'已结清账户数:{_pick_first(fields, ["settled_count"])}' if _pick_first(fields, ["settled_count"]) else '',
        ] if x])
        result['法人征信'] = _pick_first(fields, ['name', 'id_number'])
        result['成立/诉讼/变更税等级关联风险'] = _pick_first(fields, ['anomaly_notes'])
        return result

    if report_type == 'tax':
        result['开票纳税'] = '; '.join([x for x in [
            _pick_first(fields, ['tax_registration']),
            _pick_first(fields, ['has_penalty']),
            _pick_first(fields, ['tax_arrears']),
            _pick_first(fields, ['invoice_3year']),
            _pick_first(fields, ['tax_revenue_3year']),
        ] if x])
        result['成立/诉讼/变更税等级关联风险'] = _pick_first(fields, ['tax_anomaly']) or (
            '、'.join([token for token in ['纳税', '欠税', '处罚', '异常'] if token in raw_text])
        )
        return result

    return result


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
        },
        'corporate': {
            'company_name': '企业名称', 'credit_code': '统一社会信用代码',
            'report_time': '报告时间', 'unsettled_institutions': '未结清机构数',
            'total_balance': '余额', 'short_term_loan': '短期借款',
            'medium_long_term_loan': '中长期借款', 'guarantee_info': '担保信息',
            'public_info': '公共信息',
        },
        'tax': {
            'tax_registration': '纳税登记状态', 'has_penalty': '是否有滞纳金',
            'tax_arrears': '欠税金额', 'invoice_3year': '近三年开票汇总',
            'tax_revenue_3year': '近三年纳税数据', 'tax_anomaly': '税务异常说明',
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
