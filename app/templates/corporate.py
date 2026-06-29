"""企业征信报告字段抽取规则 - 含段落捕捉和明细字段"""
from typing import List, Dict, Any
from app.templates.base import BaseExtractor


class CorporateCreditExtractor(BaseExtractor):
    """企业征信报告抽取器"""

    def _define_fields(self):
        return [
            # 基本信息
            {'key': 'company_name', 'label': '企业名称', 'type': 'text', 'required': True},
            {'key': 'credit_code', 'label': '统一社会信用代码', 'type': 'text', 'required': True},
            {'key': 'report_time', 'label': '报告时间', 'type': 'date', 'required': True},

            # 成立/法人/变更/税务等级
            {'key': 'establish_info', 'label': '成立信息', 'type': 'text', 'required': False},
            {'key': 'legal_person', 'label': '法定代表人', 'type': 'text', 'required': False},
            {'key': 'tax_rating', 'label': '税务等级', 'type': 'text', 'required': False},
            {'key': 'penalty_info', 'label': '滞纳金信息', 'type': 'text', 'required': False},

            # 开票纳税
            {'key': 'invoice_data', 'label': '开票数据', 'type': 'text', 'required': False},
            {'key': 'tax_data', 'label': '纳税数据', 'type': 'text', 'required': False},

            # 信贷
            {'key': 'unsettled_institutions', 'label': '未结清机构数', 'type': 'number', 'required': False},
            {'key': 'total_balance', 'label': '余额', 'type': 'amount', 'required': False},
            {'key': 'short_term_loan', 'label': '短期借款', 'type': 'amount', 'required': False},
            {'key': 'medium_long_term_loan', 'label': '中长期借款', 'type': 'amount', 'required': False},

            # 明细
            {'key': 'loan_details_raw', 'label': '贷款明细', 'type': 'text', 'required': False},
            {'key': 'guarantee_info', 'label': '担保信息', 'type': 'text', 'required': False},
            {'key': 'public_info', 'label': '公共信息', 'type': 'text', 'required': False},
        ]

    def extract(self, ocr_items, raw_text=''):
        result = {}

        # 基本字段
        result['company_name'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['企业名称', '公司名称', '单位名称'])
        )
        result['credit_code'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['统一社会信用代码', '信用代码', '社会信用代码'])
        )
        result['report_time'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['报告时间', '查询时间', '报告日期'])
        )

        # 成立/法人信息
        result['establish_info'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['成立', '成立日期', '成立时间', '注册日期'])
        )
        result['legal_person'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['法定代表人', '法人代表', '法人'])
        )

        # 税务等级
        result['tax_rating'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['纳税等级', '税务等级', '纳税评级', 'A级', 'B级', 'M级'])
        )

        # 滞纳金
        result['penalty_info'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['滞纳金', '滞纳', '处罚', '罚款'])
        )

        # 开票纳税 - 用段落捕捉获取更多上下文
        invoice_text = self.collect_lines_after(ocr_items, '开票', max_lines=10)
        if invoice_text:
            result['invoice_data'] = self._make_text_field(invoice_text, note='段落捕捉')
        else:
            result['invoice_data'] = self._make_field(
                self.extract_by_keywords(ocr_items, ['开票', '开票汇总', '近三年开票', '发票'])
            )

        tax_text = self.collect_lines_after(ocr_items, '纳税', max_lines=15)
        if tax_text:
            result['tax_data'] = self._make_text_field(tax_text, note='段落捕捉')
        else:
            result['tax_data'] = self._make_field(
                self.extract_by_keywords(ocr_items, ['纳税', '纳税数据', '纳税金额', '缴税'])
            )

        # 信贷字段
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

        # 贷款明细 - 捕捉"未结清信贷"到"已结清"之间的整段文本
        loan_section = self.collect_lines_after(ocr_items, '未结清', max_lines=50)
        if not loan_section:
            loan_section = self.collect_lines_after(ocr_items, '贷款', max_lines=50)
        if loan_section:
            result['loan_details_raw'] = self._make_text_field(loan_section, note='段落捕捉')

        # 担保信息
        result['guarantee_info'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['担保', '对外担保', '保证担保'])
        )
        # 公共信息
        result['public_info'] = self._make_field(
            self.extract_by_keywords(ocr_items, ['公共信息', '欠税', '处罚', '法院', '诉讼', '行政处罚'])
        )

        return result

    def capture_sections(self, raw_text: str) -> Dict[str, str]:
        """分节捕捉企业征信的段落"""
        sections = {}
        lines = raw_text.split('\n')

        current_section = 'full_text'
        section_texts = {current_section: []}

        for line in lines:
            line_s = line.strip()
            if not line_s:
                continue

            # 根据关键词判断当前在哪个章节
            if any(kw in line_s for kw in ['企业名称', '公司名称', '身份标识', '基本信息']):
                current_section = 'basic_info'
            elif any(kw in line_s for kw in ['开票', '发票', '纳税', '税务']):
                current_section = 'tax_invoice'
            elif any(kw in line_s for kw in ['未结清信贷', '贷款明细', '借款']):
                current_section = 'loan_details'
            elif any(kw in line_s for kw in ['担保', '对外担保']):
                current_section = 'guarantee'
            elif any(kw in line_s for kw in ['公共信息', '欠税', '诉讼', '处罚']):
                current_section = 'public'
            elif any(kw in line_s for kw in ['查询']):
                current_section = 'query'

            if current_section not in section_texts:
                section_texts[current_section] = []
            section_texts[current_section].append(line_s)

        for k, v in section_texts.items():
            sections[k] = '\n'.join(v)

        return sections
