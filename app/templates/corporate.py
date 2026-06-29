"""企业征信报告字段抽取规则"""
from typing import List, Dict, Any
from app.templates.base import BaseExtractor


class CorporateCreditExtractor(BaseExtractor):
    """企业征信报告抽取器"""

    def _define_fields(self):
        return [
            {'key': 'company_name', 'label': '企业名称', 'type': 'text', 'required': True},
            {'key': 'credit_code', 'label': '统一社会信用代码', 'type': 'text', 'required': True},
            {'key': 'report_time', 'label': '报告时间', 'type': 'date', 'required': True},
            {'key': 'unsettled_institutions', 'label': '未结清机构数', 'type': 'number', 'required': False},
            {'key': 'total_balance', 'label': '余额', 'type': 'amount', 'required': False},
            {'key': 'short_term_loan', 'label': '短期借款', 'type': 'amount', 'required': False},
            {'key': 'medium_long_term_loan', 'label': '中长期借款', 'type': 'amount', 'required': False},
            {'key': 'guarantee_info', 'label': '担保信息', 'type': 'text', 'required': False},
            {'key': 'public_info', 'label': '公共信息', 'type': 'text', 'required': False},
        ]

    def extract(self, ocr_items, raw_text=''):
        result = {}

        result['company_name'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['企业名称', '公司名称', '单位名称'])
        )
        result['credit_code'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['统一社会信用代码', '信用代码', '社会信用代码'])
        )
        result['report_time'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['报告时间', '查询时间', '报告日期'])
        )
        result['unsettled_institutions'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['未结清', '未结清机构', '未结清余额'])
        )
        result['total_balance'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['余额', '余额合计', '总余额'])
        )
        result['short_term_loan'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['短期借款', '短期贷款'])
        )
        result['medium_long_term_loan'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['中长期借款', '长期借款', '中长期贷款'])
        )
        result['guarantee_info'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['担保', '对外担保', '保证担保'])
        )
        result['public_info'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['公共信息', '欠税', '处罚', '法院'])
        )

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
