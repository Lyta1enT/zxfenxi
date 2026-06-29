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


def _collect_summary_fields(report_type: str, fields: Dict[str, Any],
                            raw_text: str = '') -> Dict[str, Any]:
    """按用户要求的格式汇总5列数据

    列1: 公司高新/深房 → 企业名称
    列2: 成立/诉讼/变更税等级关联风险 → 成立信息+法人+税务等级+处罚
    列3: 开票纳税 → 按年开票纳税数据
    列4: 企业征信 → 贷款明细列表
    列5: 法人征信 → 个人负债+贷款明细+查询次数
    """
    result = {h: '' for h in SUMMARY_HEADERS}

    if report_type == 'corporate':
        company = _pick_first(fields, ['company_name'])
        credit_code = _pick_first(fields, ['credit_code'])
        establish = _pick_first(fields, ['establish_info'])
        legal = _pick_first(fields, ['legal_person'])
        tax_rating = _pick_first(fields, ['tax_rating'])
        penalty = _pick_first(fields, ['penalty_info'])
        public_info = _pick_first(fields, ['public_info'])

        # 列1: 公司高新/深房
        result['公司高新/深房'] = company

        # 列2: 成立/诉讼/变更税等级关联风险
        risk_parts = []
        if establish:
            risk_parts.append(establish)
        if legal:
            risk_parts.append(f'法人{legal}')
        if tax_rating:
            risk_parts.append(tax_rating)
        if penalty:
            risk_parts.append(penalty)
        # 从公共信息中补充诉讼/处罚
        for token in ['诉讼', '变更', '行政处罚', '法院', '经营异常', '违约']:
            if token in raw_text and token not in str(risk_parts):
                risk_parts.append(token)
        result['成立/诉讼/变更税等级关联风险'] = '，'.join(risk_parts) if risk_parts else ''

        # 列3: 开票纳税 - 按年格式化
        invoice_raw = _pick_first(fields, ['invoice_data', 'invoice_3year'])
        tax_raw = _pick_first(fields, ['tax_data', 'tax_revenue_3year'])
        invoice_parts = []

        # 尝试从原始文本中提取按年数据
        import re
        combined_text = raw_text + ' ' + invoice_raw + ' ' + tax_raw

        # 优先匹配完整格式: "23年开票7709万，纳税163万"
        full_pat = r'(?:20)?(\d{2})\s*年?\s*(?:开票|销售收入|销售额)[：:]?\s*([\d,.]+\s*万[元]?)[，,]\s*(?:纳税|缴税|税款)[：:]?\s*([\d,.]+\s*万[元]?)'
        full_matches = re.findall(full_pat, combined_text)

        year_map = {}
        for yr, inv_amt, tax_amt in full_matches:
            y = f'20{yr}' if len(yr) == 2 else yr
            year_map[y] = {'invoice': inv_amt, 'tax': tax_amt}

        # 补充匹配单独的开票或纳税行
        year_pat = r'(?:20)?(\d{2})\s*年?\s*(?:开票|销售收入|销售额)[：:]?\s*([\d,.]+\s*万[元]?)'
        tax_pat = r'(?:20)?(\d{2})\s*年?\s*(?:纳税|缴税|税款)[：:]?\s*([\d,.]+\s*万[元]?)'

        for yr, amt in re.findall(year_pat, combined_text):
            y = f'20{yr}' if len(yr) == 2 else yr
            if y not in year_map:
                year_map[y] = {}
            year_map[y]['invoice'] = amt

        for yr, amt in re.findall(tax_pat, combined_text):
            y = f'20{yr}' if len(yr) == 2 else yr
            if y not in year_map:
                year_map[y] = {}
            year_map[y]['tax'] = amt

        if year_map:
            for y in sorted(year_map.keys()):
                s = f'{y[-2:]}年'
                if 'invoice' in year_map[y]:
                    s += f'开票{year_map[y]["invoice"]}'
                if 'tax' in year_map[y]:
                    s += f'，纳税{year_map[y]["tax"]}'
                invoice_parts.append(s)
        else:
            # 无按年数据时，用原始文本
            if invoice_raw:
                invoice_parts.append(invoice_raw)
            if tax_raw and tax_raw != invoice_raw:
                invoice_parts.append(tax_raw)

        result['开票纳税'] = ' '.join(invoice_parts) if invoice_parts else ''

        # 列4: 企业征信 - 贷款明细
        loan_raw = _pick_first(fields, ['loan_details_raw'])
        credit_parts = []

        unsettled = _pick_first(fields, ['unsettled_institutions'])
        balance = _fmt_amount(_pick_first(fields, ['total_balance']))
        short_term = _fmt_amount(_pick_first(fields, ['short_term_loan']))
        medium_long = _fmt_amount(_pick_first(fields, ['medium_long_term_loan']))

        if loan_raw:
            # 分行整理贷款明细
            loan_lines = loan_raw.replace('\\n', '\n').split('\n')
            formatted_loans = []
            for line in loan_lines:
                line = line.strip()
                if not line:
                    continue
                # 去除可能已有的序号前缀
                clean_line = re.sub(r'^\d+[.、]\s*', '', line).strip()
                # 匹配含金额或到期信息的行
                if any(k in clean_line for k in ['万', '元', '到期', '循环', '结清']):
                    formatted_loans.append(clean_line)
                elif any(k in clean_line for k in ['邮政', '信托', '台新', '和运', '中国', '工商', '交通', '微e',
                                                       '中信', '富民', '飞泉', '三湘', '邮惠', '长安']):
                    formatted_loans.append(clean_line)
            # 重新编号
            credit_parts = [f'{i+1}.{l}' for i, l in enumerate(formatted_loans)]

        # 补充汇总信息
        summary_bits = []
        if unsettled:
            summary_bits.append(f'未结清:{unsettled}')
        if balance:
            summary_bits.append(f'余额:{balance}')
        if short_term:
            summary_bits.append(f'短期:{short_term}')
        if medium_long:
            summary_bits.append(f'中长期:{medium_long}')
        if summary_bits:
            credit_parts.append(' | '.join(summary_bits))

        result['企业征信'] = '\n'.join(credit_parts) if credit_parts else ''

        # 列5: 法人征信
        personal_parts = []
        guarantee = _pick_first(fields, ['guarantee_info'])
        if guarantee:
            personal_parts.append(guarantee)
        # 从公共信息补充
        if public_info:
            personal_parts.append(public_info)
        result['法人征信'] = '；'.join(personal_parts) if personal_parts else ''

        return result

    if report_type == 'personal':
        name = _pick_first(fields, ['name'])
        id_num = _pick_first(fields, ['id_number'])
        total_debt = _pick_first(fields, ['total_debt'])
        credit_usage = _pick_first(fields, ['credit_card_usage'])
        loans_raw = _pick_first(fields, ['personal_loans_raw'])
        overdue = _pick_first(fields, ['overdue_history', 'overdue_count'])
        repayment = _pick_first(fields, ['repayment_responsibility'])
        query = _pick_first(fields, ['query_history'])

        # 列1: 用姓名
        result['公司高新/深房'] = name or ''
        # 列2: 逾期/异常
        result['成立/诉讼/变更税等级关联风险'] = overdue or ''
        # 列3: 开票纳税（个人征信此列为空）
        result['开票纳税'] = ''

        # 列4: 企业征信 → 个人信贷汇总
        credit_items = []
        cc = _pick_first(fields, ['credit_card_count'])
        lc = _pick_first(fields, ['loan_count'])
        bal = _pick_first(fields, ['total_balance'])
        settled = _pick_first(fields, ['settled_count'])
        parts = []
        if name: parts.append(f'姓名:{name}')
        if id_num: parts.append(f'证件:{id_num}')
        if cc: parts.append(f'信用卡:{cc}')
        if lc: parts.append(f'贷款:{lc}')
        if bal: parts.append(f'余额:{bal}')
        if settled: parts.append(f'已结清:{settled}')
        credit_items.append(' | '.join(parts))

        if loans_raw:
            loan_lines = loans_raw.split('\n')
            for line in loan_lines[:15]:
                ls = line.strip()
                if ls and any(k in ls for k in ['万', '元', '到期', '循环']):
                    credit_items.append(ls)

        result['企业征信'] = '\n'.join(credit_items)

        # 列5: 法人征信 → 个人负债明细
        personal_detail = []
        if total_debt:
            personal_detail.append(f'总负债:{total_debt}')
        if credit_usage:
            personal_detail.append(f'信用卡使用率:{credit_usage}')
        if repayment:
            personal_detail.append(f'还款责任:{repayment}')
        if query:
            personal_detail.append(f'查询:{query}')

        # 如果 loans_raw 还没放进企业征信，放这里
        if not result['企业征信'] and loans_raw:
            personal_detail.append(loans_raw[:500])

        result['法人征信'] = '\n'.join(personal_detail) if personal_detail else ''

        return result

    if report_type == 'tax':
        # 列3: 开票纳税
        inv = _pick_first(fields, ['invoice_3year', 'invoice_data'])
        tax = _pick_first(fields, ['tax_revenue_3year', 'tax_data'])
        reg = _pick_first(fields, ['tax_registration'])
        penalty = _pick_first(fields, ['has_penalty'])
        arrears = _pick_first(fields, ['tax_arrears'])
        anomaly = _pick_first(fields, ['tax_anomaly'])

        # 按年格式化
        import re
        combined = raw_text + ' ' + inv + ' ' + tax

        # 优先匹配完整格式
        full_pat_tax = r'(?:20)?(\d{2})\s*年?\s*(?:开票|销售收入|销售额)[：:]?\s*([\d,.]+\s*万[元]?)[，,]\s*(?:纳税|缴税|税款)[：:]?\s*([\d,.]+\s*万[元]?)'
        full_matches_tax = re.findall(full_pat_tax, combined)

        year_map = {}
        for yr, inv_amt, tax_amt in full_matches_tax:
            y = f'20{yr}' if len(yr) == 2 else yr
            year_map[y] = {'invoice': inv_amt, 'tax': tax_amt}

        year_pat_t = r'(?:20)?(\d{2})\s*年?\s*(?:开票|销售收入|销售额)[：:]?\s*([\d,.]+\s*万[元]?)'
        tax_pat_t = r'(?:20)?(\d{2})\s*年?\s*(?:纳税|缴税|税款)[：:]?\s*([\d,.]+\s*万[元]?)'

        for yr, amt in re.findall(year_pat_t, combined):
            y = f'20{yr}' if len(yr) == 2 else yr
            if y not in year_map:
                year_map[y] = {}
            year_map[y]['invoice'] = amt

        for yr, amt in re.findall(tax_pat_t, combined):
            y = f'20{yr}' if len(yr) == 2 else yr
            if y not in year_map:
                year_map[y] = {}
            year_map[y]['tax'] = amt

        invoice_parts = []
        if year_map:
            for y in sorted(year_map.keys()):
                s = f'{y[-2:]}年'
                if 'invoice' in year_map[y]:
                    s += f'开票{year_map[y]["invoice"]}'
                if 'tax' in year_map[y]:
                    s += f'，纳税{year_map[y]["tax"]}'
                invoice_parts.append(s)
        else:
            if inv: invoice_parts.append(inv)
            if tax and tax != inv: invoice_parts.append(tax)

        result['开票纳税'] = ' '.join(invoice_parts) if invoice_parts else ''
        # 列2: 税务异常/风险
        risk = []
        if reg: risk.append(reg)
        if penalty: risk.append(penalty)
        if arrears: risk.append(f'欠税:{arrears}')
        if anomaly: risk.append(anomaly)
        result['成立/诉讼/变更税等级关联风险'] = '，'.join(risk)
        # 列4: 企业征信
        result['企业征信'] = tax or ''
        # 列1: 公司名
        result['公司高新/深房'] = ''
        # 列5: 法人征信
        result['法人征信'] = ''

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
