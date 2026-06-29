"""字段抽取引擎 - 支持版面感知和传统 OCR 两种模式"""
from typing import List, Dict, Any, Optional

from app.templates.personal import PersonalCreditExtractor
from app.templates.corporate import CorporateCreditExtractor
from app.templates.tax_report import TaxReportExtractor


EXTRACTOR_MAP = {
    'personal': PersonalCreditExtractor,
    'corporate': CorporateCreditExtractor,
    'tax': TaxReportExtractor,
}

# 表格关键词映射：表格中搜索到关键词 → 字段名
TABLE_KEYWORD_MAP = {
    'personal': {
        '信用卡': 'credit_card_count',
        '贷记卡': 'credit_card_count',
        '贷款': 'loan_count',
        '逾期': 'overdue_count',
        '余额': 'total_balance',
        '已结清': 'settled_count',
        '结清': 'settled_count',
    },
    'corporate': {
        '未结清': 'unsettled_institutions',
        '余额': 'total_balance',
        '短期借款': 'short_term_loan',
        '长期借款': 'medium_long_term_loan',
        '中长期': 'medium_long_term_loan',
        '担保': 'guarantee_info',
        '对外担保': 'guarantee_info',
        '公共信息': 'public_info',
    },
    'tax': {
        '纳税': 'tax_registration',
        '滞纳金': 'has_penalty',
        '欠税': 'tax_arrears',
        '开票': 'invoice_3year',
        '纳税数据': 'tax_revenue_3year',
        '异常': 'tax_anomaly',
    },
}


def get_extractor(report_type: str):
    """获取对应报告类型的抽取器实例"""
    cls = EXTRACTOR_MAP.get(report_type)
    if not cls:
        raise ValueError(f"不支持的报告类型: {report_type}，可选: {list(EXTRACTOR_MAP.keys())}")
    return cls()


def extract_fields(report_type: str, ocr_items: List[Dict[str, Any]] = None,
                   raw_text: str = '',
                   layout_result: Optional[Dict] = None) -> Dict[str, Any]:
    """抽取字段的主入口

    优先使用版面分析结果（layout_result），
    如果不可用则降级到传统 OCR 关键词匹配。

    Args:
        report_type: 报告类型 (personal/corporate/tax)
        ocr_items: OCR 识别结果（传统模式）
        raw_text: 源文件直接提取的文本
        layout_result: 版面分析结果（优先使用）

    Returns:
        {field_key: {value, confidence, page, note}}
    """
    if layout_result:
        return _extract_with_layout(report_type, layout_result)

    # 降级到传统 OCR 关键词匹配
    extractor = get_extractor(report_type)
    return extractor.extract(ocr_items or [], raw_text)


def _extract_with_layout(report_type: str, layout: Dict) -> Dict[str, Any]:
    """基于版面结构的字段抽取"""
    extractor = get_extractor(report_type)
    text_blocks = layout.get('text_blocks', [])
    tables = layout.get('tables', [])
    raw_text = layout.get('raw_text', '')

    # 用 text_blocks 构建传统 OCR items 供 extractor 使用
    ocr_items = []
    for block in text_blocks:
        ocr_items.append({
            'text': block['text'],
            'confidence': 1.0,
            'bbox': block.get('bbox', []),
            'page': block.get('page', 1),
        })

    # 先用 extractor 做字段抽取（关键词匹配）
    result = extractor.extract(ocr_items, raw_text)

    # 从表格中补充字段
    _enrich_from_tables(result, tables, report_type)

    # 标记哪些字段来自表格
    for key, data in result.items():
        if data.get('note') == '表格抽取':
            result[key]['confidence'] = 0.9

    return result


def _enrich_from_tables(result: Dict, tables: List[Dict], report_type: str):
    """从表格数据中补充和增强字段值"""
    keywords = TABLE_KEYWORD_MAP.get(report_type, {})

    for table in tables:
        cells = table.get('cells', [])
        if not cells:
            continue

        # 将所有行文本化
        rows_text = []
        for cell_row in cells:
            if isinstance(cell_row, list):
                row_texts = []
                for cell in cell_row:
                    if isinstance(cell, dict):
                        row_texts.append(str(cell.get('text', '')).strip())
                rows_text.append(row_texts)
            elif isinstance(cell_row, dict):
                rows_text.append([str(cell_row.get('text', '')).strip()])

        # 在每行中搜索关键词
        for row in rows_text:
            row_joined = ' '.join(row)
            for kw, field_key in keywords.items():
                if kw in row_joined:
                    # 当前字段尚未有值时才补充
                    current = result.get(field_key, {})
                    if isinstance(current, dict) and not current.get('value', ''):
                        # 取关键词所在行中非关键字的其他文本
                        values = [t for t in row if t and kw not in t and t != kw]
                        if values:
                            result[field_key] = {
                                'value': ' '.join(values),
                                'confidence': 0.9,
                                'page': 1,
                                'note': '表格抽取',
                            }


def get_field_definitions(report_type: str) -> List[Dict[str, Any]]:
    """获取指定报告类型的字段定义列表"""
    extractor = get_extractor(report_type)
    return extractor.fields
