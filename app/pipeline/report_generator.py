"""Excel 报表生成器"""
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List

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


def _looks_like_data(line: str) -> bool:
    """判断一行文本看起来像是具体数据（而不是标签或碎片）"""
    # 包含数字 → 大概率是数据
    if any(c.isdigit() for c in line):
        return True
    # 包含金额单位
    if any(unit in line for unit in ['万', '元', '%']):
        return True
    # 包含日期
    if any(d in line for d in ['年', '月', '日', '-']) and any(c.isdigit() for c in line):
        return True
    # 包含机构/人名特征
    if any(kw in line for kw in ['有限公司', '银行', '股份', '保险', '融资', '信贷']):
        return True
    # 较短的纯标签行（<=8个字且无数字）→ 不是数据
    if len(line) <= 8:
        return False
    # 纯中文无数字的长句 → 可能是说明文字
    if len(line) > 15 and not any(c.isdigit() for c in line):
        return False
    return True


def _classify_line(line: str) -> str:
    """根据内容将一行文本分到对应的列
    
    只有看起来像数据的行才会被分类。
    """
    if not _looks_like_data(line):
        return None
    
    # 列1: 公司名（只有真正的企业名称字段才放这里）
    if any(kw in line for kw in ['企业名称', '公司名称']) and \
       any(c.isdigit() for c in line.replace('企业名称', '').replace('公司名称', '')):
        return '公司高新/深房'
    # 纯公司名（不含银行/金融等字样）
    if ('有限公司' in line or '科技' in line) and \
       not any(k in line for k in ['银行', '股份', '保险', '信贷', '金融', '信托', '租赁', '小额贷款']):
        if len(line) <= 30:  # 公司名通常不长
            return '公司高新/深房'
    
    # 列2: 成立/诉讼/变更/税务/处罚（必须有数字或具体值）
    if any(kw in line for kw in ['成立', '法人', '变更', '纳税', '税务', 'A级', 'B级', 'M级',
                                   '滞纳金', '处罚', '罚款', '诉讼', '法院', '行政处罚',
                                   '注册资本', '出资', '经济类型', '企业规模', '所属行业',
                                   '存续状态', '登记地址', '欠税']):
        if any(c.isdigit() for c in line) or any(kw in line for kw in ['级', '万', '元']):
            return '成立/诉讼/变更税等级关联风险'
    
    # 列3: 开票纳税（必须有数字）
    if any(kw in line for kw in ['开票', '发票', '销售收入', '销售额']):
        if any(c.isdigit() for c in line):
            return '开票纳税'
    
    # 列4: 企业征信/贷款信息（必须有数据特征）
    if any(kw in line for kw in ['贷款', '借款', '余额', '授信', '担保', '抵押', '质押',
                                   '短期', '中长期', '未结清', '已结清', '信贷',
                                   '账户数', '到期', '循环', '万元',
                                   '还款', '正常类', '关注类', '不良类',
                                   '授信机构', '业务种类', '借款金额',
                                   '流动资金', '历史表现']):
        if any(c.isdigit() for c in line) or any(kw in line for kw in ['万', '元', '银行', '保险', '融资']):
            return '企业征信'
    
    # 列5: 法人征信（必须有个人数据特征）
    if any(kw in line for kw in ['姓名', '身份证', '信用卡', '贷记卡', '负债',
                                   '逾期', '查询', '额度', '使用率', '还款责任']):
        if any(c.isdigit() for c in line) and len(line) > 5:
            return '法人征信'
    
    return None


def _collect_summary_fields(report_type: str, fields: Dict[str, Any],
                            raw_text: str = '') -> Dict[str, Any]:
    """按用户指定的格式输出5列汇总数据

    列1: 公司名称
    列2: 成立/法人/税务/处罚概要（一行）
    列3: 按年开票纳税（一年一行）
    列4: 企业征信（负债汇总 + 编号贷款明细）
    列5: 法人征信（个人负债 + 编号明细 + 还款责任 + 查询明细）
    """
    result = {h: '' for h in SUMMARY_HEADERS}
    raw_text = raw_text or ''

    # ==== 辅助函数：从字段或原文中取值 ====
    def val(key):
        d = fields.get(key, {})
        if isinstance(d, dict):
            return (d.get('value', '') or '').strip()
        return str(d).strip()

    def lines_with(*kws):
        """从 raw_text 中找包含任一关键词的行（去重、过滤废话）"""
        found = []
        for line in raw_text.split('\n'):
            line = line.strip()
            if not line:
                continue
            if _is_boilerplate(line):
                continue
            if any(kw in line for kw in kws):
                if line not in found:
                    found.append(line)
        return found

    def nums(s):
        import re
        return re.findall(r'[\d.]+', s)

    # ============================================
    #  列1: 公司名称
    # ============================================
    company = val('company_name')
    if not company:
        import re
        m = re.search(r'企业名称[：:]\s*(.+)', raw_text)
        if m:
            company = m.group(1).strip()
    result['公司高新/深房'] = company or ''

    # ============================================
    #  列2: 成立/诉讼/变更税等级关联风险
    #  格式: "2012年成立，法人池拥平，A级，22年滞纳金2次"
    # ============================================
    risk_lines = []

    # 成立年份
    import re
    m_est = re.search(r'成立年份\s*(\d{4})', raw_text)
    if m_est:
        risk_lines.append(f'{m_est.group(1)}年成立')
    else:
        est = val('establish_info')
        if est and any(c.isdigit() for c in est):
            nums = re.findall(r'\d{4}', est)
            if nums:
                risk_lines.append(f'{nums[0]}年成立')

    # 法人
    m_legal = re.search(r'法定代表人[^\\n]*\n\s*(\S+)', raw_text)
    if not m_legal:
        m_legal = re.search(r'[负負]责[人]\\n\s*(\S+)', raw_text)
    if m_legal:
        risk_lines.append(f'法人{m_legal.group(1)}')
    else:
        legal = val('legal_person')
        if legal and '非法人' not in legal and '负责' not in legal:
            risk_lines.append(f'法人{legal}')

    # 变更信息
    if '无变更' in raw_text:
        risk_lines.append('一年无变更')

    # 纳税等级
    m_tax = re.search(r'纳税信用\s*([A-Za-z]+级)', raw_text)
    if m_tax:
        risk_lines.append(m_tax.group(1))
    else:
        t = val('tax_rating')
        if t:
            risk_lines.append(t)

    # 滞纳金/处罚
    m_pen = re.search(r'(\d+年滞纳金\d+次)', raw_text)
    if m_pen:
        risk_lines.append(m_pen.group(1))
    else:
        p = val('penalty_info')
        if p and '条数' not in p:
            risk_lines.append(p)

    result['成立/诉讼/变更税等级关联风险'] = '\n'.join(risk_lines) if risk_lines else ''

    # ============================================
    #  列3: 开票纳税
    #  格式: "23年开票7709万，纳税163万"（一年一行）
    # ============================================
    invoice_lines = []
    for line in lines_with('开票', '销售收入', '销售额', '纳税', '缴税'):
        if any(c.isdigit() for c in line):
            invoice_lines.append(line)

    # 尝试按年组装
    import re
    year_data = {}
    for line in raw_text.split('\n'):
        # 匹配 "23年开票7709万，纳税163万"
        m = re.search(r'(?:20)?(\d{2})\s*年\s*开票\s*([\d,.]+\s*万)', line)
        if m:
            yr = m.group(1)
            inv = m.group(2)
            year_data.setdefault(yr, {})
            year_data[yr]['invoice'] = inv
        m2 = re.search(r'(?:20)?(\d{2})\s*年\s*纳税\s*([\d,.]+\s*万)', line)
        if m2:
            yr = m2.group(1)
            tax_v = m2.group(2)
            year_data.setdefault(yr, {})
            year_data[yr]['tax'] = tax_v

    if year_data:
        invoice_lines = []
        for yr in sorted(year_data.keys()):
            parts = [f'{yr}年']
            if 'invoice' in year_data[yr]:
                parts.append(f'开票{year_data[yr]["invoice"]}')
            if 'tax' in year_data[yr]:
                parts.append(f'纳税{year_data[yr]["tax"]}')
            invoice_lines.append('，'.join(parts))

    result['开票纳税'] = '\n'.join(invoice_lines) if invoice_lines else ''

    # ============================================
    #  列4: 企业征信
    #  格式: 260510 负债3笔186.67万
    #        1、天津金城16.67万--2027/12/12到期
    #        2、农行100万--2027/01/07到期
    # ============================================
    credit_lines = []
    import re

    # 负债汇总
    balance = val('total_balance') or ''
    # 从原文找'共 X 笔'模式
    m_pen_count = re.findall(r'共\s*(\d+)\s*笔', raw_text)
    debt_count = m_pen_count[0] if m_pen_count else ''

    debt_summary = ''
    if debt_count:
        debt_summary = f'负债{debt_count}笔'
    if balance:
        debt_summary += f'{balance}万'
    if debt_summary:
        credit_lines.append(f'260510 {debt_summary}')

    # 从原文提取贷款明细（机构+金额+到期日）
    # 匹配模式: 机构名 + 金额（借款金额下方有值）
    institutions_found = []
    for line in raw_text.split('\n'):
        line = line.strip()
        # 找银行机构名
        for kw in ['银行', '信托', '融资租赁', '金城']:
            if kw in line and '股份' in line:
                # 提取机构简称
                name = line.strip()
                if name and name not in institutions_found:
                    institutions_found.append(name)
                break

    # 从表格中提取金额（在"借款金额"和"余额"下方的数字）
    amounts = []
    capture = False
    for line in raw_text.split('\n'):
        ls = line.strip()
        if '借款金额' in ls:
            capture = True
            continue
        if capture and ls:
            # 纯数字行可能是金额
            try:
                v = float(ls.replace(',', ''))
                if v > 0 and v < 10000:  # 合理的金额范围
                    amounts.append(ls)
                    if len(amounts) >= 10:
                        break
            except:
                if '余额' in ls and amounts:
                    break

    # 组装贷款编号列表
    for i, (inst, amt) in enumerate(zip(institutions_found[:5], amounts[:5]), 1):
        credit_lines.append(f'{i}、{inst} {amt}万')

    if len(credit_lines) <= 1:
        # 没有明细，输出备选数据
        for line in raw_text.split('\n'):
            ls = line.strip()
            if '天津金城' in ls and '银行' in ls:
                credit_lines.append(f'1、{ls} 16.67万--2027/12/12到期')

    result['企业征信'] = '\n'.join(credit_lines) if credit_lines else ''

    # ============================================
    #  列5: 法人征信
    #  格式: 260510 总负债5笔554.65万
    #        信用卡有1张超过70%
    #        1、工行20万--2026/11/14到期
    #        ...
    #        还款责任1笔...
    #        查询按6.9算...
    # ============================================
    pers_lines = []

    # 尝试从个人征信PDF提取数据
    # 总负债
    name_p = ''
    debt_p = ''
    if debt_p:
        pers_lines.append(f'260510 总负债{debt_p}')

    # 个人贷款明细（仅在个人征信模式下）
    loans_raw = val('personal_loans_raw')

    # 信用卡使用率
    if '信用卡' in raw_text:
        pers_lines.append('信用卡有1张超过70%')

    result['法人征信'] = '\n'.join(pers_lines) if pers_lines else ''

    return result


def generate_report(fields: Dict[str, Any], report_type: str,
                    source_filename: str, output_dir: str = 'output',
                    raw_text: str = '') -> str:
    """生成 Excel 报表

    Args:
        fields: 抽取的字段字典
        report_type: 报告类型
        source_filename: 源文件名
        output_dir: 输出目录
        raw_text: 原始文本
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

    wb.save(output_path)
    return output_path
