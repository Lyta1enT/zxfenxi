"""水母/税务报告字段抽取规则"""
from typing import List, Dict, Any
from app.templates.base import BaseExtractor
import re


class TaxReportExtractor(BaseExtractor):
    """水母报告（税务报告）抽取器"""

    def _define_fields(self):
        return [
            {'key': 'tax_registration', 'label': '纳税登记状态', 'type': 'text', 'required': True},
            {'key': 'has_penalty', 'label': '是否有滞纳金', 'type': 'text', 'required': False},

            # 按年开票纳税数据
            {'key': 'invoice_year_1', 'label': '最早年开票', 'type': 'text', 'required': False},
            {'key': 'tax_year_1', 'label': '最早年纳税', 'type': 'text', 'required': False},
            {'key': 'invoice_year_2', 'label': '中间年开票', 'type': 'text', 'required': False},
            {'key': 'tax_year_2', 'label': '中间年纳税', 'type': 'text', 'required': False},
            {'key': 'invoice_year_3', 'label': '最近年开票', 'type': 'text', 'required': False},
            {'key': 'tax_year_3', 'label': '最近年纳税', 'type': 'text', 'required': False},

            # 汇总字段
            {'key': 'invoice_3year', 'label': '近三年开票汇总', 'type': 'text', 'required': False},
            {'key': 'tax_revenue_3year', 'label': '近三年纳税数据', 'type': 'amount', 'required': False},
            {'key': 'tax_arrears', 'label': '欠税金额', 'type': 'amount', 'required': False},
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

        # 开票纳税 - 段落捕捉
        invoice_section = self.collect_lines_after(ocr_items, '开票', max_lines=20)
        tax_section = self.collect_lines_after(ocr_items, '纳税', max_lines=20)

        result['invoice_3year'] = self._make_text_field(
            invoice_section or '', note='段落捕捉'
        ) if invoice_section else self._make_field(
            self.extract_by_keywords(ocr_items, ['开票', '开票汇总', '近三年开票'])
        )

        result['tax_revenue_3year'] = self._make_text_field(
            tax_section or '', note='段落捕捉'
        ) if tax_section else self._make_field(
            self.extract_by_keywords(ocr_items, ['纳税', '纳税数据', '纳税金额'])
        )

        # 从段落中解析按年数据
        combined = f"{invoice_section} {tax_section}"
        year_data = self._parse_year_data(combined)
        for key, val in year_data.items():
            result[key] = self._make_text_field(val, note='年份解析')

        # 欠税
        result['tax_arrears'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['欠税', '欠税金额', '欠缴税款'])
        )

        # 异常
        result['tax_anomaly'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['异常', '税务异常', '说明'])
        )

        arrears = result.get('tax_arrears', {}).get('value', '')
        if arrears and '无' not in arrears and arrears.strip():
            result['tax_anomaly'] = {
                'value': f'存在欠税: {arrears}',
                'confidence': 0.8,
                'page': 0,
                'note': '自动标记'
            }

        return result

    def _parse_year_data(self, text: str) -> dict:
        """从文本中解析按年开票纳税数据
        
        匹配模式: "23年开票7709万，纳税163万"
        """
        result = {}
        # 匹配 "XX年开票XXX万" 或 "XX年开票XXX万元"
        year_pattern = re.findall(
            r'(\d{2,4})\s*年?\s*(?:开票|销售收入|销售额)[：:]?\s*([\d,.]+\s*万?)',
            text
        )
        tax_pattern = re.findall(
            r'(\d{2,4})\s*年?\s*(?:纳税|缴税|税款)[：:]?\s*([\d,.]+\s*万?)',
            text
        )

        years_found = set()
        for year_str, amount in year_pattern:
            year_key = f'20{year_str}' if len(year_str) == 2 else year_str
            years_found.add(year_key)
            # 按时间排序
            result[f'invoice_{year_key}'] = f'{year_key}年开票{amount}'

        for year_str, amount in tax_pattern:
            year_key = f'20{year_str}' if len(year_str) == 2 else year_str
            years_found.add(year_key)
            if f'invoice_{year_key}' in result:
                result[f'invoice_{year_key}'] += f'，纳税{amount}'
            else:
                result[f'tax_{year_key}'] = f'{year_key}年纳税{amount}'

        return result
