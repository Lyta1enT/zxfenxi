"""字段抽取规则模板基类"""
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional


class BaseExtractor(ABC):
    """抽取规则基类，每种报告类型继承此类"""
    
    def __init__(self):
        self.name = self.__class__.__name__
        self.fields = self._define_fields()
    
    @abstractmethod
    def _define_fields(self) -> List[Dict[str, Any]]:
        """定义要抽取的字段列表
        
        每个字段:
        {
            'key': 'field_name',
            'label': '显示名称',
            'type': 'text|date|amount|number',
            'required': bool,
        }
        """
        pass
    
    @abstractmethod
    def extract(self, ocr_items: List[Dict[str, Any]],
                raw_text: str = '') -> Dict[str, Any]:
        """从 OCR 结果中抽取字段
        
        Args:
            ocr_items: OCR 识别结果列表，每项含 text/confidence/bbox/page
            raw_text: 直接从文件中提取的文本（如有）
            
        Returns:
            {field_key: {value, confidence, page, note}}
        """
        pass
    
    def validate_field(self, field_key: str, value: str,
                       field_type: str) -> Optional[str]:
        """校验单个字段值，返回错误信息或 None"""
        if not value or value.strip() == '':
            return None
        
        value = value.strip()
        
        if field_type == 'date':
            import re
            if not re.match(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', value):
                return f"日期格式异常: {value}"
        
        elif field_type == 'amount':
            import re
            if not re.search(r'\d+', value):
                return f"金额格式异常: {value}"
        
        elif field_type == 'number':
            if not value.isdigit():
                return f"数字格式异常: {value}"
        
        return None
    
    def extract_by_keywords(self, ocr_items: List[Dict[str, Any]],
                            keywords: List[str],
                            context_radius: int = 3) -> List[Dict[str, Any]]:
        """通过关键词匹配抽取附近文本"""
        import re
        matched = []
        for i, item in enumerate(ocr_items):
            text = item['text'].strip()
            for kw in keywords:
                if kw in text:
                    # 去掉关键词前后的标点和空格
                    context = text.replace(kw, '').strip().lstrip('：:，,。.、')
                    if not context and i + 1 < len(ocr_items):
                        context = ocr_items[i + 1]['text'].strip().lstrip('：:，,。.、')

                    matched.append({
                        'value': context,
                        'confidence': item['confidence'],
                        'page': item['page'],
                    })
                    break
        return matched
