# PaddleOCR 3.0 + PP-StructureV3 升级实施计划

**Goal:** 升级 PaddleOCR 到 3.0 并接入 PP-StructureV3，用版面分析+表格识别替代关键词硬匹配

**Files to create/modify:**
- Modify: `requirements.txt` — paddleocr>=3.0.0
- Modify: `app/pipeline/ocr_engine.py` — PP-OCRv5
- Create: `app/pipeline/layout_analyzer.py` — PP-StructureV3 封装
- Modify: `app/pipeline/field_extractor.py` — 版面感知抽取
- Modify: `app/pipeline/worker.py` — 集成 LayoutAnalyzer
- Modify: `app/pipeline/report_generator.py` — 输出结构化表格

---

### Task 1: 升级依赖 + OCR 引擎

**Files:**
- Modify: `requirements.txt`
- Modify: `app/pipeline/ocr_engine.py`

**Changes:**

1. requirements.txt: `paddleocr>=3.0.0`

2. ocr_engine.py — 升级到 PP-OCRv5:

```python
class OCREngine:
    def __init__(self, use_angle_cls=True, lang='ch', ocr_version='PP-OCRv5'):
        self._ocr = None
        self.use_angle_cls = use_angle_cls
        self.lang = lang
        self.ocr_version = ocr_version
        self._initialized = False

    def _lazy_init(self):
        if not self._initialized:
            from paddleocr import PaddleOCR
            self._ocr = PaddleOCR(
                lang=self.lang,
                ocr_version=self.ocr_version,
                use_angle_cls=self.use_angle_cls,
                show_log=False,
            )
            self._initialized = True
    # ... rest remains same
```

### Task 2: 新建 LayoutAnalyzer

**Files:**
- Create: `app/pipeline/layout_analyzer.py`

封装 PP-StructureV3，提供统一的结构化文档解析接口。

```python
"""版面分析引擎 - 封装 PP-StructureV3"""
from pathlib import Path
from typing import List, Dict, Any, Optional


class LayoutElement:
    """版面元素"""
    def __init__(self, type_tag: str, content: Any, bbox=None, page=1):
        self.type = type_tag  # text, table, figure, title, list, etc.
        self.content = content
        self.bbox = bbox
        self.page = page


class LayoutAnalyzer:
    """版面分析器 - 使用 PP-StructureV3 解析文档结构"""
    
    def __init__(self):
        self._engine = None
        self._initialized = False
    
    def _lazy_init(self):
        if not self._initialized:
            from paddleocr import PPStructureV3
            self._engine = PPStructureV3(show_log=False)
            self._initialized = True
    
    def analyze(self, file_path: str) -> Dict[str, Any]:
        """分析文档，返回结构化结果
        
        Returns:
            {
                'elements': [LayoutElement, ...],
                'tables': [table_data, ...],
                'text_blocks': [...],
                'raw_text': str,
            }
        """
        self._lazy_init()
        result = self._engine.predict(file_path)
        return self._parse_result(result)
    
    def _parse_result(self, result) -> Dict[str, Any]:
        elements = []
        tables = []
        text_blocks = []
        raw_text_parts = []
        
        for item in result:
            elem_type = item.get('type', 'text')
            content = item.get('content', {}).get('text', '') if isinstance(item.get('content'), dict) else str(item.get('content', ''))
            bbox = item.get('bbox')
            page = item.get('page', 1)
            
            elements.append(LayoutElement(elem_type, content, bbox, page))
            
            if elem_type == 'table':
                tables.append(item)
                # Extract text from table cells
                if 'cells' in item:
                    for row in item['cells']:
                        for cell in row if isinstance(row, list) else [row]:
                            if isinstance(cell, dict) and cell.get('text'):
                                raw_text_parts.append(cell['text'])
            else:
                if isinstance(content, str) and content.strip():
                    text_blocks.append({'text': content.strip(), 'bbox': bbox, 'page': page})
                    raw_text_parts.append(content.strip())
        
        return {
            'elements': elements,
            'tables': tables,
            'text_blocks': text_blocks,
            'raw_text': '\n'.join(raw_text_parts),
        }
```

### Task 3: 重写 FieldExtractor — 版面感知抽取

**Files:**
- Modify: `app/pipeline/field_extractor.py`

利用版面结构做字段抽取，而不是纯关键词匹配。

```python
def extract_fields(report_type: str, ocr_items=None, raw_text='',
                   layout_result: Optional[Dict] = None) -> Dict[str, Any]:
    """抽取字段 - 优先使用版面分析结果"""
    
    if layout_result:
        return _extract_with_layout(report_type, layout_result)
    else:
        # 降级到原来的 OCR 关键词匹配
        extractor = get_extractor(report_type)
        return extractor.extract(ocr_items or [], raw_text)


def _extract_with_layout(report_type: str, layout: Dict) -> Dict[str, Any]:
    """基于版面结构的字段抽取"""
    extractor = get_extractor(report_type)
    text_blocks = layout.get('text_blocks', [])
    tables = layout.get('tables', [])
    
    # 从 text_blocks 构建类似 OCR items 的结构供 extractor 使用
    ocr_items = []
    for block in text_blocks:
        ocr_items.append({
            'text': block['text'],
            'confidence': 1.0,
            'bbox': block.get('bbox', []),
            'page': block.get('page', 1),
        })
    
    # 用原来的 extractor 做字段抽取
    result = extractor.extract(ocr_items, layout.get('raw_text', ''))
    
    # 额外从表格中提取信息
    for table in tables:
        _enrich_from_table(result, table, report_type)
    
    return result


def _enrich_from_table(result: Dict, table: Dict, report_type: str):
    """从表格中补充字段值"""
    cells = table.get('cells', [])
    if not cells:
        return
    
    # 第一行可能是表头
    headers = []
    for cell in cells[0] if isinstance(cells[0], list) else []:
        if isinstance(cell, dict):
            headers.append(str(cell.get('text', '')).strip())
    
    # 在表格中搜索关键字段
    keyword_map = {
        'personal': {
            '信用卡': 'credit_card_count',
            '贷款': 'loan_count', 
            '逾期': 'overdue_count',
            '余额': 'total_balance',
            '已结清': 'settled_count',
        },
        'corporate': {
            '未结清': 'unsettled_institutions',
            '余额': 'total_balance',
            '短期借款': 'short_term_loan',
            '担保': 'guarantee_info',
        },
        'tax': {
            '纳税': 'tax_registration',
            '滞纳金': 'has_penalty',
            '欠税': 'tax_arrears',
            '开票': 'invoice_3year',
        },
    }
    
    keywords = keyword_map.get(report_type, {})
    
    for row in cells:
        row_text = ''
        if isinstance(row, list):
            row_text = ' '.join([str(c.get('text', '')) if isinstance(c, dict) else str(c) for c in row])
        elif isinstance(row, dict):
            row_text = str(row.get('text', ''))
        
        for kw, field_key in keywords.items():
            if kw in row_text and (not result.get(field_key, {}).get('value', '')):
                # 从行中提取值（在表头之外的列中找数字或关键值）
                values = []
                if isinstance(row, list):
                    for cell in row:
                        if isinstance(cell, dict) and cell.get('text'):
                            text = str(cell.get('text', '')).strip()
                            if text and text != kw:
                                values.append(text)
                if values:
                    result[field_key] = {
                        'value': ' '.join(values),
                        'confidence': 0.9,
                        'page': 1,
                        'note': '表格抽取',
                    }
```

### Task 4: 集成 Worker + 报告

**Files:**
- Modify: `app/pipeline/worker.py`
- Modify: `app/pipeline/report_generator.py`

Worker 中增加 LayoutAnalyzer 步骤，report_generator 增加表格数据 sheet。

---

### 执行顺序

1. Task 1: 升级依赖 + OCR 引擎（10分钟）
2. Task 2: 新建 LayoutAnalyzer（15分钟）
3. Task 3: 重写 FieldExtractor（20分钟）
4. Task 4: 集成 Worker + Report（15分钟）
5. 测试验证（10分钟）
