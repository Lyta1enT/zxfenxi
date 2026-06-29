"""个人征信报告字段抽取规则"""
from typing import List, Dict, Any
from app.templates.base import BaseExtractor


class PersonalCreditExtractor(BaseExtractor):
    """个人征信报告抽取器"""

    def _define_fields(self):
        return [
            {'key': 'name', 'label': '姓名', 'type': 'text', 'required': True},
            {'key': 'id_number', 'label': '证件号码', 'type': 'text', 'required': True},
            {'key': 'report_time', 'label': '报告时间', 'type': 'date', 'required': True},
            {'key': 'credit_card_count', 'label': '信用卡账户数', 'type': 'number', 'required': False},
            {'key': 'loan_count', 'label': '贷款账户数', 'type': 'number', 'required': False},
            {'key': 'overdue_count', 'label': '逾期账户数', 'type': 'number', 'required': False},
            {'key': 'total_balance', 'label': '余额', 'type': 'amount', 'required': False},
            {'key': 'settled_count', 'label': '已结清账户数', 'type': 'number', 'required': False},
            {'key': 'anomaly_notes', 'label': '异常备注', 'type': 'text', 'required': False},
        ]

    def extract(self, ocr_items, raw_text=''):
        result = {}
        full_text = ' '.join([item['text'] for item in ocr_items])
        if raw_text:
            full_text = raw_text + '\n' + full_text

        result['name'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['姓名', '姓名：', '姓名:'])
        )
        result['id_number'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['证件号码', '证件号', '身份证'])
        )
        result['report_time'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['报告时间', '查询时间', '报告日期'])
        )
        result['credit_card_count'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['信用卡', '贷记卡'])
        )
        result['loan_count'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['贷款', '贷款账户'])
        )
        result['overdue_count'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['逾期', '逾期账户'])
        )
        result['total_balance'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['余额', '余额合计'])
        )
        result['settled_count'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['已结清', '结清'])
        )

        # 自动标记逾期异常
        overdue_val = result.get('overdue_count', {}).get('value', '')
        import re
        nums = re.findall(r'\d+', str(overdue_val))
        if nums and int(nums[0]) > 0:
            result['anomaly_notes'] = {
                'value': f'\u26a0\ufe0f 存在逾期记录 ({nums[0]}个账户)',
                'confidence': 1.0,
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
