"""Word 报告生成器"""
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Optional

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


REPORT_TYPE_NAMES = {
    'personal': '个人征信报告',
    'corporate': '企业征信报告',
    'tax': '水母报告（税务分析）',
}


def set_cell_shading(cell, color: str):
    """设置单元格背景色"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


def add_table_borders(table):
    """为表格添加边框"""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '4')
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), '000000')
        borders.append(element)
    tblPr.append(borders)


FIELD_DEFS = {
    'personal': [
        ('name', '姓名'), ('id_number', '证件号码'), ('report_time', '报告时间'),
        ('credit_card_count', '信用卡账户数'), ('loan_count', '贷款账户数'),
        ('overdue_count', '逾期账户数'), ('total_balance', '余额'),
        ('settled_count', '已结清账户数'),
    ],
    'corporate': [
        ('company_name', '企业名称'), ('credit_code', '统一社会信用代码'),
        ('report_time', '报告时间'), ('unsettled_institutions', '未结清机构数'),
        ('total_balance', '余额'), ('short_term_loan', '短期借款'),
        ('medium_long_term_loan', '中长期借款'),
    ],
    'tax': [
        ('tax_registration', '纳税登记状态'), ('has_penalty', '是否有滞纳金'),
        ('tax_arrears', '欠税金额'), ('invoice_3year', '近三年开票汇总'),
        ('tax_revenue_3year', '近三年纳税数据'),
    ],
}


def generate_report(fields: Dict[str, Any], report_type: str,
                    source_filename: str, output_dir: str = 'output') -> str:
    """生成 Word 报告

    Args:
        fields: 抽取的字段字典
        report_type: 报告类型
        source_filename: 源文件名（用于命名输出文件）
        output_dir: 输出目录

    Returns:
        生成的 Word 文件路径
    """
    os.makedirs(output_dir, exist_ok=True)

    base_name = Path(source_filename).stem
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'{base_name}_{report_type}_报告_{timestamp}.docx')

    doc = Document()

    # 设置默认字体
    style = doc.styles['Normal']
    font = style.font
    font.name = 'SimSun'
    font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')

    # ====== 标题 ======
    title = doc.add_heading(REPORT_TYPE_NAMES.get(report_type, '征信报告'), level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f'源文件: {source_filename}  |  生成时间: {timestamp}')
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    doc.add_paragraph()

    # ====== 摘要表 ======
    doc.add_heading('一、关键字段摘要', level=1)

    field_keys = [k for k, _ in FIELD_DEFS.get(report_type, [])]
    labels = dict(FIELD_DEFS.get(report_type, []))

    table = doc.add_table(rows=len(field_keys) + 1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_table_borders(table)

    # 表头
    hdr = table.rows[0]
    for i, text in enumerate(['字段名称', '字段值']):
        cell = hdr.cells[i]
        cell.text = text
        set_cell_shading(cell, 'D9E2F3')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for r in p.runs:
                r.bold = True

    # 填充数据
    for idx, key in enumerate(field_keys):
        row = table.rows[idx + 1]
        row.cells[0].text = labels.get(key, key)

        field_data = fields.get(key, {})
        value = field_data.get('value', '')
        note = field_data.get('note', '')

        display_value = value if value else '[未识别]'
        row.cells[1].text = display_value

        if note == '未识别' or not value:
            set_cell_shading(row.cells[1], 'FFF2CC')
        elif '\u26a0\ufe0f' in str(value):
            set_cell_shading(row.cells[1], 'FCE4EC')

    doc.add_paragraph()

    # ====== 异常标记 ======
    doc.add_heading('二、异常标记', level=1)

    anomalies = []
    for key, data in fields.items():
        value = data.get('value', '')
        note = data.get('note', '')
        label = labels.get(key, key)

        if note == '未识别':
            anomalies.append(f'\u26a0\ufe0f {label}: 未能识别，建议人工核查')
        elif '\u26a0\ufe0f' in str(value):
            anomalies.append(f'\U0001f534 {label}: {value}')
        elif '\u8fc7\u671f' in str(value) or '\u6b20\u7a0e' in str(value) or '\u5f02\u5e38' in str(value):
            anomalies.append(f'\U0001f7e1 {label}: {value}')

    if anomalies:
        for a in anomalies:
            p = doc.add_paragraph(a)
            r = p.runs[0]
            if '\U0001f534' in a:
                r.font.color.rgb = RGBColor(0xCC, 0x33, 0x00)
            else:
                r.font.color.rgb = RGBColor(0xCC, 0x88, 0x00)
    else:
        doc.add_paragraph('\u2705 未发现明显异常')

    doc.add_paragraph()

    # ====== 补充说明 ======
    doc.add_heading('三、原始数据摘要', level=1)
    p = doc.add_paragraph(f'本报告基于 {source_filename} 通过 OCR 识别和规则抽取生成。')
    p.add_run('\n\n字段置信度说明：')
    p.add_run('\n   \u2022 高置信度 (>0.9): 识别可靠')
    p.add_run('\n   \u2022 中置信度 (0.6-0.9): 建议复核')
    p.add_run('\n   \u2022 低置信度 (<0.6): 需人工确认')

    doc.save(output_path)
    return output_path
