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
    cls = EXTRACTOR_MAP.get(report_type)
    if not cls:
        raise ValueError(f"不支持的报告类型: {report_type}，可选: {list(EXTRACTOR_MAP.keys())}")
    return cls()


def _ensure_ocr_items(ocr_items: Optional[List[Dict[str, Any]]],
                      raw_text: str) -> List[Dict[str, Any]]:
    """确保 ocr_items 不为空，必要时从 raw_text 构造"""
    if ocr_items and len(ocr_items) > 0:
        return ocr_items
    if raw_text:
        lines = raw_text.split('\n')
        items = []

        # 跳过报告说明页（"第 1 页"之前的内容）
        page2_start = 0
        for i, line in enumerate(lines):
            if '第 2 页' in line or '第2页' in line:
                page2_start = i
                break

        for i, line in enumerate(lines):
            stripped = line.strip()
            if not stripped:
                continue
            # 跳过页眉页码行
            if '第 ' in stripped and '页' in stripped:
                continue
            # 跳过大段报告说明（第1页内容）
            if i < page2_start and any(kw in stripped for kw in [
                '报告说明', '本报告由', '征信中心', '本报告所展示',
                '本报告中信贷交易', '如无特别说明', '信息主体有权',
                '信息主体声明',
            ]):
                continue

            items.append({
                'text': stripped,
                'confidence': 1.0,
                'bbox': [],
                'page': page2_start + 1,  # 统一标记为数据页
            })
        return items
    return []


def extract_fields(report_type: str, ocr_items: List[Dict[str, Any]] = None,
                   raw_text: str = '',
                   layout_result: Optional[Dict] = None) -> Dict[str, Any]:
    """抽取字段的主入口

    优先使用版面分析结果（layout_result），
    如果不可用则降级到关键词匹配（从 ocr_items 或 raw_text）。
    """
    if layout_result and layout_result.get('text_blocks'):
        return _extract_with_layout(report_type, layout_result)

    # 确保 ocr_items 有内容，过滤模板说明文字
    ocr_items = _ensure_ocr_items(ocr_items, raw_text)
    extractor = get_extractor(report_type)
    ocr_items = extractor.filter_boilerplate(ocr_items)
    return extractor.extract(ocr_items, raw_text)


def _extract_with_layout(report_type: str, layout: Dict) -> Dict[str, Any]:
    """基于版面结构的字段抽取"""
    extractor = get_extractor(report_type)
    text_blocks = layout.get('text_blocks', [])
    tables = layout.get('tables', [])
    raw_text = layout.get('raw_text', '')

    ocr_items = []
    for block in text_blocks:
        ocr_items.append({
            'text': block['text'],
            'confidence': 1.0,
            'bbox': block.get('bbox', []),
            'page': block.get('page', 1),
        })

    result = extractor.extract(ocr_items, raw_text)
    _enrich_from_tables(result, tables, report_type)

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

        for row in rows_text:
            row_joined = ' '.join(row)
            for kw, field_key in keywords.items():
                if kw in row_joined:
                    current = result.get(field_key, {})
                    if isinstance(current, dict) and not current.get('value', ''):
                        values = [t for t in row if t and kw not in t and t != kw]
                        if values:
                            result[field_key] = {
                                'value': ' '.join(values),
                                'confidence': 0.9,
                                'page': 1,
                                'note': '表格抽取',
                            }


def get_field_definitions(report_type: str) -> List[Dict[str, Any]]:
    extractor = get_extractor(report_type)
    return extractor.fields
