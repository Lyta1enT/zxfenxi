"""字段抽取规则模板基类"""
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional
import re


class BaseExtractor(ABC):
    """抽取规则基类，每种报告类型继承此类"""

    # 报告模板说明文字（跳过这些内容）
    BOILERPLATE_KEYWORDS = [
        '本报告由中国人民银行', '报告说明', '征信中心', '本报告所展示',
        '本报告中信贷交易', '本报告中借贷交易', '如无特别说明',
        '信息主体有权', '如有异议', '第 ', '页/共', '信息主体声明',
        '征信中心说明', '数据提供机构说明', '如信息记录斜体',
        '本报告仅展示', '信息主体对', '报数机构',
    ]

    def __init__(self):
        self.name = self.__class__.__name__
        self.fields = self._define_fields()

    def filter_boilerplate(self, ocr_items: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """过滤掉报告模板说明文字"""
        return [
            item for item in ocr_items
            if not any(kw in item['text'] for kw in self.BOILERPLATE_KEYWORDS)
        ]

    @abstractmethod
    def _define_fields(self) -> List[Dict[str, Any]]:
        pass

    @abstractmethod
    def extract(self, ocr_items: List[Dict[str, Any]],
                raw_text: str = '') -> Dict[str, Any]:
        pass

    def capture_sections(self, raw_text: str) -> Dict[str, str]:
        """捕捉整段文本并按主题分节
        
        子类可重写此方法添加更精确的分节逻辑
        """
        return {'full_text': raw_text}

    def validate_field(self, field_key: str, value: str,
                       field_type: str) -> Optional[str]:
        if not value or value.strip() == '':
            return None
        value = value.strip()
        if field_type == 'date':
            if not re.match(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', value):
                return f"日期格式异常: {value}"
        elif field_type == 'amount':
            if not re.search(r'\d+', value):
                return f"金额格式异常: {value}"
        elif field_type == 'number':
            if not value.isdigit():
                return f"数字格式异常: {value}"
        return None

    def extract_by_keywords(self, ocr_items: List[Dict[str, Any]],
                            keywords: List[str],
                            context_radius: int = 3) -> List[Dict[str, Any]]:
        """通过关键词匹配抽取附近文本

        支持两种模式：
        1. 同行模式: "公司名称：XX科技" → "XX科技"
        2. 下一行模式: "公司名称" + "XX科技"（换行）→ "XX科技"
        """
        matched = []
        for i, item in enumerate(ocr_items):
            text = item['text'].strip()
            for kw in keywords:
                if kw in text:
                    context = text.replace(kw, '').strip().lstrip('：:，,。.、')
                    # 如果本行剩余很短（像是标签），检查下一行
                    LABEL_WORDS = {'年份', '日期', '信息', '情况', '号码',
                                   '代码', '状态', '名称', '类型', '方式'}
                    if (not context or context in LABEL_WORDS) and i + 1 < len(ocr_items):
                        context = ocr_items[i + 1]['text'].strip().lstrip('：:，,。.、')
                    # 值太长(>80)且不含数字 → 过滤掉
                    if len(context) > 80 and not any(c.isdigit() for c in context):
                        continue
                    matched.append({
                        'value': context,
                        'confidence': item['confidence'],
                        'page': item['page'],
                    })
                    break
        return matched

    def extract_section(self, ocr_items: List[Dict[str, Any]],
                        section_start: str, section_end: str = None) -> str:
        """提取两个关键词之间的整段文本"""
        lines = []
        capturing = False
        for item in ocr_items:
            text = item['text'].strip()
            if section_start in text:
                capturing = True
                # 去掉起始关键词的文本
                remain = text.replace(section_start, '').strip().lstrip('：:，,。.、')
                if remain:
                    lines.append(remain)
                continue
            if capturing and section_end and section_end in text:
                # 包含结束行
                remain = text.split(section_end)[0].strip()
                if remain:
                    lines.append(remain)
                break
            if capturing:
                lines.append(text)
        return ' '.join(lines) if lines else ''

    def collect_lines_after(self, ocr_items: List[Dict[str, Any]],
                            keyword: str, max_lines: int = 20) -> str:
        """提取关键词后面的若干行"""
        lines = []
        found = False
        for item in ocr_items:
            text = item['text'].strip()
            if not found and keyword in text:
                found = True
                remain = text.replace(keyword, '').strip().lstrip('：:，,。.、')
                if remain and remain != keyword:
                    lines.append(remain)
                continue
            if found:
                if len(lines) >= max_lines:
                    break
                # 遇到明显的下一个章节标题时停止
                if any(t in text for t in ['信息概要', '身份标识', '征信报告',
                                            '未结清', '已结清', '担保',
                                            '公共信息', '查询']):
                    if len(lines) > 2:
                        break
                lines.append(text)
        return '\n'.join(lines) if lines else ''

    def _make_field(self, matches):
        if matches and matches[0].get('value', '').strip():
            m = matches[0]
            val = m['value'].strip()
            # 值太长（>80）且不含数字 → 大概率是废话
            if len(val) > 80 and not any(c.isdigit() for c in val):
                return {'value': '', 'confidence': 0, 'page': 0, 'note': '过滤(废话)'}
            return {
                'value': val,
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

    def _make_text_field(self, text: str, confidence: float = 1.0,
                         page: int = 0, note: str = ''):
        # 清理：去掉每行中的报告说明碎片
        clean_lines = []
        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue
            # 跳过包含 boilerplate 关键词的行
            skip_words = ['本报告', '如无特别', '信息主体', '征信中心',
                         '常见的产品', '是指除', '五级分类', '逾期总额',
                         '资产管理公司', '担保交易指', '第三人实质上']
            if any(kw in line for kw in skip_words):
                continue
            # 只保留含数字或机构名的行
            if any(c.isdigit() for c in line) or any(k in line for k in ['银行', '公司', '融资']):
                clean_lines.append(line)
        text = '\n'.join(clean_lines)
        
        # 过滤后没有内容则返回未识别
        if not text.strip():
            return {'value': '', 'confidence': 0, 'page': 0, 'note': '未识别'}
        return {
            'value': text[:500],
            'confidence': confidence,
            'page': page,
            'note': note,
        }
