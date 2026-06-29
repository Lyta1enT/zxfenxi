"""字段抽取引擎 - 根据报告类型选择对应的抽取器"""
from typing import List, Dict, Any, Optional

from app.templates.personal import PersonalCreditExtractor
from app.templates.corporate import CorporateCreditExtractor
from app.templates.tax_report import TaxReportExtractor


EXTRACTOR_MAP = {
    'personal': PersonalCreditExtractor,
    'corporate': CorporateCreditExtractor,
    'tax': TaxReportExtractor,
}


def get_extractor(report_type: str):
    """获取对应报告类型的抽取器实例"""
    cls = EXTRACTOR_MAP.get(report_type)
    if not cls:
        raise ValueError(f"不支持的报告类型: {report_type}，可选: {list(EXTRACTOR_MAP.keys())}")
    return cls()


def extract_fields(report_type: str, ocr_items: List[Dict[str, Any]],
                   raw_text: str = '') -> Dict[str, Any]:
    """抽取字段的主入口

    Args:
        report_type: 报告类型 (personal/corporate/tax)
        ocr_items: OCR 识别结果
        raw_text: 源文件直接提取的文本

    Returns:
        {field_key: {value, confidence, page, note}}
    """
    extractor = get_extractor(report_type)
    return extractor.extract(ocr_items, raw_text)


def get_field_definitions(report_type: str) -> List[Dict[str, Any]]:
    """获取指定报告类型的字段定义列表"""
    extractor = get_extractor(report_type)
    return extractor.fields
