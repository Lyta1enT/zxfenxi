"""个人征信报告字段抽取规则 - 含法人征信明细"""
from typing import List, Dict, Any
from app.templates.base import BaseExtractor
import re


class PersonalCreditExtractor(BaseExtractor):
    """个人征信报告抽取器（法人征信）"""

    def _define_fields(self):
        return [
            # 基本信息
            {'key': 'name', 'label': '姓名', 'type': 'text', 'required': True},
            {'key': 'id_number', 'label': '证件号码', 'type': 'text', 'required': True},
            {'key': 'report_time', 'label': '报告时间', 'type': 'date', 'required': True},

            # 信贷汇总
            {'key': 'credit_card_count', 'label': '信用卡账户数', 'type': 'number', 'required': False},
            {'key': 'loan_count', 'label': '贷款账户数', 'type': 'number', 'required': False},
            {'key': 'overdue_count', 'label': '逾期账户数', 'type': 'number', 'required': False},
            {'key': 'total_balance', 'label': '余额', 'type': 'amount', 'required': False},
            {'key': 'settled_count', 'label': '已结清账户数', 'type': 'number', 'required': False},

            # 法人征信专用
            {'key': 'total_debt', 'label': '总负债', 'type': 'text', 'required': False},
            {'key': 'credit_card_usage', 'label': '信用卡使用率', 'type': 'text', 'required': False},
            {'key': 'personal_loans_raw', 'label': '个人贷款明细', 'type': 'text', 'required': False},
            {'key': 'overdue_history', 'label': '逾期历史', 'type': 'text', 'required': False},
            {'key': 'repayment_responsibility', 'label': '还款责任', 'type': 'text', 'required': False},
            {'key': 'query_history', 'label': '查询次数', 'type': 'text', 'required': False},
            {'key': 'anomaly_notes', 'label': '异常备注', 'type': 'text', 'required': False},
        ]

    def extract(self, ocr_items, raw_text=''):
        result = {}
        full_text = ' '.join([item['text'] for item in ocr_items])
        if raw_text:
            full_text = raw_text + '\n' + full_text

        # 基本信息
        result['name'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['姓名', '姓名：', '姓名:'])
        )
        result['id_number'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['证件号码', '证件号', '身份证'])
        )
        result['report_time'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['报告时间', '查询时间', '报告日期'])
        )

        # 信贷汇总
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

        # 总负债 - "总负债X笔XXX万" 模式
        debt_text = self.collect_lines_after(ocr_items, '负债', max_lines=5)
        if debt_text:
            result['total_debt'] = self._make_text_field(debt_text, note='段落捕捉')
        else:
            # 尝试从余额推断
            result['total_debt'] = self._make_field(
                self.extract_by_keywords(ocr_items, ['负债', '总负债', '合计负债'])
            )

        # 信用卡使用率
        result['credit_card_usage'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['使用率', '信用卡使用', '额度使用'])
        )

        # 个人贷款明细 - 捕捉整段贷款信息
        loan_section = self.collect_lines_after(ocr_items, '贷款', max_lines=60)
        if loan_section:
            result['personal_loans_raw'] = self._make_text_field(loan_section, note='段落捕捉')

        # 逾期历史
        overdue_section = self.collect_lines_after(ocr_items, '逾期', max_lines=30)
        if overdue_section:
            result['overdue_history'] = self._make_text_field(overdue_section, note='段落捕捉')

        # 还款责任
        result['repayment_responsibility'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['还款责任', '担保责任', '共同还款', '保证人'])
        )

        # 查询次数 - 捕捉"查询"后面的段落
        query_section = self.collect_lines_after(ocr_items, '查询', max_lines=20)
        if query_section:
            result['query_history'] = self._make_text_field(query_section, note='段落捕捉')

        # 自动标记逾期异常
        overdue_val = result.get('overdue_count', {}).get('value', '')
        nums = re.findall(r'\d+', str(overdue_val))
        if nums and int(nums[0]) > 0:
            result['anomaly_notes'] = {
                'value': f'\u26a0\ufe0f 存在逾期记录 ({nums[0]}个账户)',
                'confidence': 1.0,
                'page': 0,
                'note': '自动标记'
            }

        return result

    def capture_sections(self, raw_text: str) -> Dict[str, str]:
        sections = {}
        lines = raw_text.split('\n')
        current = 'full_text'
        for line in lines:
            ls = line.strip()
            if not ls:
                continue
            if any(kw in ls for kw in ['姓名', '证件号码', '身份信息']):
                current = 'basic'
            elif any(kw in ls for kw in ['信息概要', '信贷汇总']):
                current = 'summary'
            elif '信用卡' in ls:
                current = 'credit_card'
            elif '贷款' in ls:
                current = 'loans'
            elif '逾期' in ls:
                current = 'overdue'
            elif '查询' in ls:
                current = 'query'
            elif any(kw in ls for kw in ['负债', '还款责任']):
                current = 'debt'
            if current not in sections:
                sections[current] = []
            sections[current].append(ls)
        return {k: '\n'.join(v) for k, v in sections.items()}
