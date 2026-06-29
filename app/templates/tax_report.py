"""水母/税务报告字段抽取规则"""
from typing import List, Dict, Any
from app.templates.base import BaseExtractor


class TaxReportExtractor(BaseExtractor):
    """水母报告（税务报告）抽取器"""

    def _define_fields(self):
        return [
            {'key': 'tax_registration', 'label': '纳税登记状态', 'type': 'text', 'required': True},
            {'key': 'has_penalty', 'label': '是否有滞纳金', 'type': 'text', 'required': False},
            {'key': 'tax_arrears', 'label': '欠税金额', 'type': 'amount', 'required': False},
            {'key': 'invoice_3year', 'label': '近三年开票汇总', 'type': 'text', 'required': False},
            {'key': 'tax_revenue_3year', 'label': '近三年纳税数据', 'type': 'amount', 'required': False},
            {'key': 'tax_anomaly', 'label': '税务异常说明', 'type': 'text', 'required': False},
        ]

    def extract(self, ocr_items, raw_text=''):
        result = {}

        result['tax_registration'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['纳税登记', '税务登记', '纳税人状态'])
        )
        result['has_penalty'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['滞纳金', '是否有滞纳金'])
        )
        result['tax_arrears'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['欠税', '欠税金额', '欠缴税款'])
        )
        result['invoice_3year'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['开票', '开票汇总', '近三年开票'])
        )
        result['tax_revenue_3year'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['纳税', '纳税数据', '纳税金额'])
        )
        result['tax_anomaly'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['异常', '税务异常', '说明'])
        )

        # 自动检测欠税异常
        arrears = result.get('tax_arrears', {}).get('value', '')
        if arrears and '\u65e0' not in arrears and arrears.strip():
            result['tax_anomaly'] = {
                'value': f'存在欠税: {arrears}',
                'confidence': 0.8,
                'page': 0,
                'note': '自动标记'
            }

        return result

    def _make_field(self, matches):
        if matches and matches[0].get('value', '').strip():
            m = matches[0]
            return {
                'value': m['value'],
                'confidence': m['confidence'],
                'page': m['page'],
                'note': ''
            }
        return {
            'value': '',
            'confidence': 0,
            'page': 0,
            'note': '未识别'
        }
