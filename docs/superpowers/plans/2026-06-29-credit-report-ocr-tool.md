# 征信报告OCR识别与生成工具 - 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 构建一个 PyQt5 桌面应用，支持 PDF/图片/Word/Excel 上传，PaddleOCR 识别，规则抽取字段，生成 Word 报告

**Architecture:** 单体 PyQt5 桌面应用，核心管道（FileHandler → OCREngine → FieldExtractor → ReportGenerator）在 QThread 后台执行

**Tech Stack:** Python 3, PyQt5, PaddleOCR, python-docx, PyMuPDF, pdf2image, opencv-python

---

### Task 1: 项目脚手架

**Files:**
- Create: `app/__init__.py`
- Create: `app/utils/__init__.py`
- Create: `app/pipeline/__init__.py`
- Create: `app/templates/__init__.py`
- Create: `app/widgets/__init__.py`
- Create: `requirements.txt`

- [ ] **Step 1: 创建项目目录结构**

```bash
mkdir -p app/utils app/pipeline app/templates app/widgets output
```

- [ ] **Step 2: 创建所有 `__init__.py` 文件**

```python
# app/__init__.py
# 征信报告OCR识别与生成工具
```

```python
# app/utils/__init__.py
from . import pdf_utils, image_utils
```

```python
# app/pipeline/__init__.py
from . import file_handler, ocr_engine, field_extractor, report_generator, worker
```

```python
# app/templates/__init__.py
from . import base, personal, corporate, tax_report
```

```python
# app/widgets/__init__.py
from . import upload_area, preview_panel, log_panel
```

- [ ] **Step 3: 创建 requirements.txt**

```
PyQt5>=5.15.0
paddlepaddle>=2.5.0
paddleocr>=2.7.0
python-docx>=1.0.0
PyMuPDF>=1.23.0
pdf2image>=1.16.0
opencv-python>=4.8.0
Pillow>=10.0.0
python-pptx>=0.6.21
openpyxl>=3.1.0
pytest>=7.0.0
```

- [ ] **Step 4: 提交**

```bash
git add -A && git commit -m "chore: scaffold project structure"
```

---

### Task 2: 工具模块 - PDF 和图片处理

**Files:**
- Create: `app/utils/pdf_utils.py`
- Create: `app/utils/image_utils.py`

- [ ] **Step 1: 实现 pdf_utils.py**

```python
"""PDF 处理工具"""
import io
from pathlib import Path
from typing import List, Tuple, Optional

import fitz  # PyMuPDF


def extract_text_from_pdf(pdf_path: str) -> str:
    """从 PDF 中直接提取文本"""
    doc = fitz.open(pdf_path)
    texts = []
    for page in doc:
        texts.append(page.get_text())
    doc.close()
    return "\n".join(texts)


def pdf_to_images(pdf_path: str, dpi: int = 200) -> List[str]:
    """将 PDF 每页转换为图片，返回图片路径列表
    
    使用 PyMuPDF 进行转换，比 pdf2image 更快且无需 poppler
    """
    doc = fitz.open(pdf_path)
    image_paths = []
    base_path = str(Path(pdf_path).with_suffix(""))
    
    for page_num, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        img_path = f"{base_path}_page_{page_num + 1}.png"
        pix.save(img_path)
        image_paths.append(img_path)
    
    doc.close()
    return image_paths


def extract_text_from_pdf_with_fallback(pdf_path: str) -> Tuple[str, List[str]]:
    """尝试提取 PDF 文本，如果内容不足则转为图片
    
    Returns:
        (提取的文本, 图片路径列表)
    """
    text = extract_text_from_pdf(pdf_path)
    
    # 如果提取的文本太少（<50字符），认为是扫描件，转为图片
    if len(text.strip()) < 50:
        images = pdf_to_images(pdf_path)
        return text, images
    
    return text, []
```

- [ ] **Step 2: 实现 image_utils.py**

```python
"""图片预处理工具"""
from typing import List, Tuple, Optional

import cv2
import numpy as np
from PIL import Image


def preprocess_image(image_path: str) -> np.ndarray:
    """图片预处理：灰度化、去噪、二值化"""
    img = cv2.imread(image_path)
    if img is None:
        raise ValueError(f"无法读取图片: {image_path}")
    
    # 灰度化
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    
    # 去噪
    denoised = cv2.fastNlMeansDenoising(gray, h=30)
    
    # 二值化
    _, binary = cv2.threshold(denoised, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    
    return binary


def deskew_image(image: np.ndarray) -> np.ndarray:
    """矫正图片倾斜"""
    coords = np.column_stack(np.where(image > 0))
    if len(coords) == 0:
        return image
    
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = 90 + angle
    
    h, w = image.shape[:2]
    center = (w // 2, h // 2)
    matrix = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(
        image, matrix, (w, h),
        flags=cv2.INTER_CUBIC,
        borderMode=cv2.BORDER_REPLICATE
    )
    return rotated
```

- [ ] **Step 3: 提交**

```bash
git add -A && git commit -m "feat: add pdf and image utilities"
```

---

### Task 3: 文件处理模块

**Files:**
- Create: `app/pipeline/file_handler.py`

- [ ] **Step 1: 实现 file_handler.py**

```python
"""文件类型判断和处理"""
import os
from pathlib import Path
from typing import Dict, Any, Optional

from app.utils.pdf_utils import extract_text_from_pdf_with_fallback
from app.utils.image_utils import preprocess_image


SUPPORTED_EXTENSIONS = {
    '.pdf': 'PDF',
    '.png': 'IMAGE',
    '.jpg': 'IMAGE',
    '.jpeg': 'IMAGE',
    '.bmp': 'IMAGE',
    '.tiff': 'IMAGE',
    '.docx': 'WORD',
    '.xlsx': 'EXCEL',
}


def get_file_type(file_path: str) -> str:
    """判断文件类型，返回类型代码"""
    ext = Path(file_path).suffix.lower()
    return SUPPORTED_EXTENSIONS.get(ext, 'UNKNOWN')


def is_supported(file_path: str) -> bool:
    """检查文件是否受支持"""
    return get_file_type(file_path) != 'UNKNOWN'


def process_file(file_path: str) -> Dict[str, Any]:
    """处理上传的文件，返回统一格式的结果
    
    Returns:
        {
            'file_type': str,
            'file_path': str,
            'text': str,          # 直接提取的文本（如有）
            'image_paths': list,   # 图片路径列表（用于OCR）
            'metadata': dict       # 文件元信息
        }
    """
    file_type = get_file_type(file_path)
    
    result = {
        'file_type': file_type,
        'file_path': file_path,
        'text': '',
        'image_paths': [],
        'metadata': {'pages': 0},
    }
    
    if file_type == 'PDF':
        text, images = extract_text_from_pdf_with_fallback(file_path)
        result['text'] = text
        result['image_paths'] = images
        result['metadata']['pages'] = len(images) if images else 1
    
    elif file_type == 'IMAGE':
        result['image_paths'] = [file_path]
        result['metadata']['pages'] = 1
    
    elif file_type == 'WORD':
        # 从 Word 中提取文本（可选补充信息）
        try:
            from docx import Document
            doc = Document(file_path)
            text = '\n'.join([p.text for p in doc.paragraphs])
            result['text'] = text
        except Exception:
            pass
    
    elif file_type == 'EXCEL':
        # 从 Excel 中提取文本
        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, read_only=True)
            texts = []
            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                for row in ws.iter_rows(values_only=True):
                    row_text = ' '.join([str(c) for c in row if c is not None])
                    if row_text.strip():
                        texts.append(row_text)
            result['text'] = '\n'.join(texts)
            wb.close()
        except Exception:
            pass
    
    return result
```

- [ ] **Step 2: 提交**

```bash
git add -A && git commit -m "feat: add file handler module"
```

---

### Task 4: OCR 引擎模块

**Files:**
- Create: `app/pipeline/ocr_engine.py`

- [ ] **Step 1: 实现 ocr_engine.py**

```python
"""PaddleOCR 引擎封装"""
import os
from typing import List, Dict, Any, Optional

import numpy as np
import cv2

from app.utils.image_utils import preprocess_image, deskew_image


# OCR 结果条目结构
# {
#     'text': str,
#     'confidence': float,
#     'bbox': [[x1,y1],[x2,y2],[x3,y3],[x4,y4]],
#     'page': int,
# }


class OCREngine:
    """封装 PaddleOCR，提供统一接口"""
    
    def __init__(self, use_angle_cls: bool = True, lang: str = 'ch', 
                 enable_mkldnn: bool = True):
        self._ocr = None
        self.use_angle_cls = use_angle_cls
        self.lang = lang
        self.enable_mkldnn = enable_mkldnn
        self._initialized = False
    
    def _lazy_init(self):
        """延迟初始化 PaddleOCR（避免 import 耗时阻塞 UI）"""
        if not self._initialized:
            from paddleocr import PaddleOCR
            self._ocr = PaddleOCR(
                use_angle_cls=self.use_angle_cls,
                lang=self.lang,
                enable_mkldnn=self.enable_mkldnn,
                show_log=False,
            )
            self._initialized = True
    
    def recognize_image(self, image_path: str, page: int = 1) -> List[Dict[str, Any]]:
        """识别单张图片，返回 OCR 结果列表"""
        self._lazy_init()
        
        # 预处理
        processed = preprocess_image(image_path)
        
        # OCR 识别
        result = self._ocr.ocr(processed, cls=True)
        
        # 解析结果
        items = []
        if result and result[0]:
            for line in result[0]:
                bbox, (text, confidence) = line
                items.append({
                    'text': text,
                    'confidence': confidence,
                    'bbox': bbox,
                    'page': page,
                })
        
        return items
    
    def recognize_batch(self, image_paths: List[str]) -> List[Dict[str, Any]]:
        """批量识别多张图片"""
        all_results = []
        for i, img_path in enumerate(image_paths):
            items = self.recognize_image(img_path, page=i + 1)
            all_results.extend(items)
        return all_results
    
    def recognize_text_only(self, image_paths: List[str]) -> str:
        """仅返回识别的文本内容（拼接所有结果）"""
        results = self.recognize_batch(image_paths)
        return '\n'.join([item['text'] for item in results])
```

- [ ] **Step 2: 提交**

```bash
git add -A && git commit -m "feat: add OCR engine wrapper"
```

---

### Task 5: 抽取模板基类

**Files:**
- Create: `app/templates/base.py`

- [ ] **Step 1: 实现 base.py**

```python
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
            return None  # 空值不校验
        
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
        matched = []
        for i, item in enumerate(ocr_items):
            text = item['text'].strip()
            for kw in keywords:
                if kw in text:
                    # 收集关键词后面的内容（同一条或后几条）
                    context = text.replace(kw, '').strip()
                    if not context and i + 1 < len(ocr_items):
                        context = ocr_items[i + 1]['text'].strip()
                    
                    matched.append({
                        'value': context,
                        'confidence': item['confidence'],
                        'page': item['page'],
                    })
                    break
        return matched
```

- [ ] **Step 2: 提交**

```bash
git add -A && git commit -m "feat: add base extractor template"
```

---

### Task 6: 三种报告模板实现

**Files:**
- Create: `app/templates/personal.py`
- Create: `app/templates/corporate.py`
- Create: `app/templates/tax_report.py`

- [ ] **Step 1: 实现 personal.py（个人征信）**

```python
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
        
        # 姓名
        name_matches = self.extract_by_keywords(ocr_items, ['姓名', '姓名：', '姓名:'])
        result['name'] = self._make_field(name_matches[0] if name_matches else None)
        
        # 证件号码
        id_matches = self.extract_by_keywords(ocr_items, ['证件号码', '证件号', '身份证'])
        result['id_number'] = self._make_field(id_matches[0] if id_matches else None)
        
        # 报告时间
        time_matches = self.extract_by_keywords(ocr_items, ['报告时间', '查询时间', '报告日期'])
        result['report_time'] = self._make_field(time_matches[0] if time_matches else None)
        
        # 信用卡账户数
        cc_matches = self.extract_by_keywords(ocr_items, ['信用卡', '贷记卡'])
        result['credit_card_count'] = self._make_field(cc_matches[0] if cc_matches else None)
        
        # 贷款账户数
        loan_matches = self.extract_by_keywords(ocr_items, ['贷款', '贷款账户'])
        result['loan_count'] = self._make_field(loan_matches[0] if loan_matches else None)
        
        # 逾期账户数
        overdue_matches = self.extract_by_keywords(ocr_items, ['逾期', '逾期账户'])
        result['overdue_count'] = self._make_field(overdue_matches[0] if overdue_matches else None)
        
        # 余额
        balance_matches = self.extract_by_keywords(ocr_items, ['余额', '余额合计'])
        result['total_balance'] = self._make_field(balance_matches[0] if balance_matches else None)
        
        # 已结清账户数
        settled_matches = self.extract_by_keywords(ocr_items, ['已结清', '结清'])
        result['settled_count'] = self._make_field(settled_matches[0] if settled_matches else None)
        
        # 异常备注 - 如果有逾期
        overdue_count_val = result.get('overdue_count', {}).get('value', '0')
        if any(c.isdigit() for c in overdue_count_val):
            nums = [c for c in overdue_count_val if c.isdigit()]
            if nums and int(nums[0]) > 0:
                result['anomaly_notes'] = {
                    'value': f'⚠️ 存在逾期记录 ({overdue_count_val}个账户)',
                    'confidence': 1.0,
                    'page': 0,
                    'note': '自动标记'
                }
        
        return result
    
    def _make_field(self, match):
        if match:
            return {
                'value': match['value'],
                'confidence': match['confidence'],
                'page': match['page'],
                'note': ''
            }
        return {
            'value': '',
            'confidence': 0,
            'page': 0,
            'note': '未识别'
        }
```

- [ ] **Step 2: 实现 corporate.py（企业征信）**

```python
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
        if matches and matches[0]['value']:
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
```

- [ ] **Step 3: 实现 tax_report.py（水母报告）**

```python
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
        
        # 自动标记异常
        arrears = result.get('tax_arrears', {}).get('value', '')
        if arrears and '无' not in arrears and arrears.strip():
            result['tax_anomaly'] = {
                'value': f'存在欠税: {arrears}',
                'confidence': 0.8,
                'page': 0,
                'note': '自动标记'
            }
        
        return result
    
    def _make_field(self, matches):
        if matches and matches[0]['value']:
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
```

- [ ] **Step 4: 提交**

```bash
git add -A && git commit -m "feat: add three report type extractors"
```

---

### Task 7: 字段抽取引擎

**Files:**
- Create: `app/pipeline/field_extractor.py`

- [ ] **Step 1: 实现 field_extractor.py**

```python
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
```

- [ ] **Step 2: 提交**

```bash
git add -A && git commit -m "feat: add field extraction engine"
```

---

### Task 8: Word 报告生成器

**Files:**
- Create: `app/pipeline/report_generator.py`

- [ ] **Step 1: 实现 report_generator.py**

```python
"""Word 报告生成器"""
import os
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, List, Optional

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


REPORT_TYPE_NAMES = {
    'personal': '个人征信报告',
    'corporate': '企业征信报告',
    'tax': '水母报告（税务分析）',
}


def set_cell_shading(cell, color: str):
    """设置单元格背景色"""
    shading = OxmlElement('w:shd')
    shading.set(qn('w:fill'), color)
    shading.set(qn('w:val'), 'clear')
    cell._tc.get_or_add_tcPr().append(shading)


def add_table_borders(table):
    """为表格添加边框"""
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    borders = OxmlElement('w:tblBorders')
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        element = OxmlElement(f'w:{edge}')
        element.set(qn('w:val'), 'single')
        element.set(qn('w:sz'), '4')
        element.set(qn('w:space'), '0')
        element.set(qn('w:color'), '000000')
        borders.append(element)
    tblPr.append(borders)


def generate_report(fields: Dict[str, Any], report_type: str, 
                    source_filename: str, output_dir: str = 'output') -> str:
    """生成 Word 报告
    
    Args:
        fields: 抽取的字段字典
        report_type: 报告类型
        source_filename: 源文件名（用于命名输出文件）
        output_dir: 输出目录
        
    Returns:
        生成的 Word 文件路径
    """
    os.makedirs(output_dir, exist_ok=True)
    
    # 输出文件名
    base_name = Path(source_filename).stem
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_path = os.path.join(output_dir, f'{base_name}_{report_type}_报告_{timestamp}.docx')
    
    doc = Document()
    
    # 设置默认字体
    style = doc.styles['Normal']
    font = style.font
    font.name = 'SimSun'
    font.size = Pt(10.5)
    style.element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    
    # ====== 标题 ======
    title = doc.add_heading(REPORT_TYPE_NAMES.get(report_type, '征信报告'), level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # 副标题信息
    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(f'源文件: {source_filename}  |  生成时间: {timestamp}')
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    
    doc.add_paragraph()  # 空行
    
    # ====== 摘要表 ======
    h1 = doc.add_heading('一、关键字段摘要', level=1)
    
    # 确定字段列表
    field_map = {
        'personal': [
            'name', 'id_number', 'report_time', 'credit_card_count', 
            'loan_count', 'overdue_count', 'total_balance', 'settled_count'
        ],
        'corporate': [
            'company_name', 'credit_code', 'report_time', 'unsettled_institutions',
            'total_balance', 'short_term_loan', 'medium_long_term_loan'
        ],
        'tax': [
            'tax_registration', 'has_penalty', 'tax_arrears', 
            'invoice_3year', 'tax_revenue_3year'
        ],
    }
    
    field_keys = field_map.get(report_type, [])
    field_defs = {
        'personal': [
            ('name', '姓名'), ('id_number', '证件号码'), ('report_time', '报告时间'),
            ('credit_card_count', '信用卡账户数'), ('loan_count', '贷款账户数'),
            ('overdue_count', '逾期账户数'), ('total_balance', '余额'),
            ('settled_count', '已结清账户数'),
        ],
        'corporate': [
            ('company_name', '企业名称'), ('credit_code', '统一社会信用代码'),
            ('report_time', '报告时间'), ('unsettled_institutions', '未结清机构数'),
            ('total_balance', '余额'), ('short_term_loan', '短期借款'),
            ('medium_long_term_loan', '中长期借款'),
        ],
        'tax': [
            ('tax_registration', '纳税登记状态'), ('has_penalty', '是否有滞纳金'),
            ('tax_arrears', '欠税金额'), ('invoice_3year', '近三年开票汇总'),
            ('tax_revenue_3year', '近三年纳税数据'),
        ],
    }
    
    labels = dict(field_defs.get(report_type, []))
    
    # 创建摘要表格（2列：字段名 → 值）
    table = doc.add_table(rows=len(field_keys) + 1, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    add_table_borders(table)
    
    # 表头
    hdr = table.rows[0]
    for i, text in enumerate(['字段名称', '字段值']):
        cell = hdr.cells[i]
        cell.text = text
        set_cell_shading(cell, 'D9E2F3')
        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
    
    # 填充数据
    for idx, key in enumerate(field_keys):
        row = table.rows[idx + 1]
        row.cells[0].text = labels.get(key, key)
        
        field_data = fields.get(key, {})
        value = field_data.get('value', '')
        note = field_data.get('note', '')
        
        display_value = value if value else f'[未识别]'
        row.cells[1].text = display_value
        
        # 如果字段有异常或未识别，标黄
        if note == '未识别' or not value:
            set_cell_shading(row.cells[1], 'FFF2CC')
        elif '⚠️' in str(value):
            set_cell_shading(row.cells[1], 'FCE4EC')
    
    doc.add_paragraph()
    
    # ====== 异常标记 ======
    doc.add_heading('二、异常标记', level=1)
    
    anomalies = []
    for key, data in fields.items():
        value = data.get('value', '')
        note = data.get('note', '')
        label = labels.get(key, key)
        
        if note == '未识别':
            anomalies.append(f'⚠️ {label}: 未能识别，建议人工核查')
        elif '⚠️' in str(value):
            anomalies.append(f'🔴 {label}: {value}')
        elif '逾期' in str(value) or '欠税' in str(value) or '异常' in str(value):
            anomalies.append(f'🟡 {label}: {value}')
    
    if anomalies:
        for a in anomalies:
            p = doc.add_paragraph(a)
            run = p.runs[0]
            run.font.color.rgb = RGBColor(0xCC, 0x33, 0x00) if '🔴' in a else RGBColor(0xCC, 0x88, 0x00)
    else:
        doc.add_paragraph('✅ 未发现明显异常')
    
    doc.add_paragraph()
    
    # ====== 明细 / 补充 ======
    doc.add_heading('三、原始数据摘要', level=1)
    p = doc.add_paragraph(f'本报告基于 {source_filename} 通过 OCR 识别和规则抽取生成。')
    p.add_run('\n\n字段置信度说明：')
    p.add_run('\n   • 高置信度 (>0.9): 识别可靠')
    p.add_run('\n   • 中置信度 (0.6-0.9): 建议复核')
    p.add_run('\n   • 低置信度 (<0.6): 需人工确认')
    
    # 保存文档
    doc.save(output_path)
    return output_path
```

- [ ] **Step 2: 提交**

```bash
git add -A && git commit -m "feat: add Word report generator"
```

---

### Task 9: 后台工作线程

**Files:**
- Create: `app/pipeline/worker.py`

- [ ] **Step 1: 实现 worker.py**

```python
"""后台工作线程 - 在 QThread 中执行处理管道"""
import traceback
from typing import List, Dict, Any, Optional

from PyQt5.QtCore import QThread, pyqtSignal

from app.pipeline.file_handler import process_file
from app.pipeline.ocr_engine import OCREngine
from app.pipeline.field_extractor import extract_fields, get_field_definitions
from app.pipeline.report_generator import generate_report


class ProcessingSignals(QThread):
    """工作线程信号"""
    progress = pyqtSignal(int, str)       # (进度百分比, 状态描述)
    log = pyqtSignal(str, str)            # (级别, 消息)
    result_ready = pyqtSignal(dict)       # 抽取结果就绪
    report_ready = pyqtSignal(str)        # 报告文件路径
    error = pyqtSignal(str)               # 错误信息
    finished = pyqtSignal()               # 全部完成


class ProcessingWorker(QThread):
    """文件处理工作线程"""
    
    def __init__(self, file_path: str, report_type: str, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.report_type = report_type
        self.signals = ProcessingSignals()
        self._ocr_engine = OCREngine()
    
    def run(self):
        try:
            # Step 1: 文件处理
            self.signals.progress.emit(5, '正在分析文件...')
            self.signals.log.emit('INFO', f'开始处理文件: {self.file_path}')
            
            file_result = process_file(self.file_path)
            file_type = file_result['file_type']
            image_paths = file_result['image_paths']
            raw_text = file_result['text']
            
            self.signals.log.emit('INFO', f'文件类型: {file_type}, 图片页数: {len(image_paths)}')
            
            # Step 2: OCR 识别
            self.signals.progress.emit(20, '正在进行 OCR 识别...')
            self.signals.log.emit('INFO', f'启动 OCR 识别 ({len(image_paths)} 页)...')
            
            ocr_items = []
            if image_paths:
                ocr_items = self._ocr_engine.recognize_batch(image_paths)
                self.signals.log.emit('INFO', f'OCR 识别完成，共 {len(ocr_items)} 个文本块')
            elif raw_text:
                self.signals.log.emit('INFO', f'文件已有文本内容 ({len(raw_text)} 字符)，跳过 OCR')
            
            self.signals.progress.emit(60, '正在抽取字段...')
            
            # Step 3: 字段抽取
            fields = extract_fields(self.report_type, ocr_items, raw_text)
            field_defs = get_field_definitions(self.report_type)
            
            # 统计命中率
            hit_count = sum(1 for f in field_defs if fields.get(f['key'], {}).get('value', ''))
            total = len(field_defs)
            self.signals.log.emit('INFO', f'字段抽取完成: {hit_count}/{total} 字段命中')
            
            # 发送结果预览
            self.signals.result_ready.emit(fields)
            self.signals.progress.emit(80, '正在生成 Word 报告...')
            
            # Step 4: 生成报告
            output_path = generate_report(
                fields=fields,
                report_type=self.report_type,
                source_filename=self.file_path,
            )
            
            self.signals.log.emit('INFO', f'✅ Word 报告已生成: {output_path}')
            self.signals.report_ready.emit(output_path)
            self.signals.progress.emit(100, '处理完成')
            self.signals.finished.emit()
            
        except Exception as e:
            error_msg = f'处理失败: {str(e)}\n{traceback.format_exc()}'
            self.signals.log.emit('ERROR', error_msg)
            self.signals.error.emit(str(e))
```

- [ ] **Step 2: 提交**

```bash
git add -A && git commit -m "feat: add background processing worker"
```

---

### Task 10: UI 组件

**Files:**
- Create: `app/widgets/upload_area.py`
- Create: `app/widgets/preview_panel.py`
- Create: `app/widgets/log_panel.py`

- [ ] **Step 1: 实现 upload_area.py（拖拽上传组件）**

```python
"""文件上传区域组件 - 支持拖拽和点击选择"""
import os
from pathlib import Path

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QListWidget, QListWidgetItem, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent


SUPPORTED_FORMATS = "PDF (*.pdf);;图片 (*.png *.jpg *.jpeg *.bmp);;Word (*.docx);;Excel (*.xlsx);;所有支持文件 (*.pdf *.png *.jpg *.jpeg *.bmp *.docx *.xlsx)"


class DropArea(QLabel):
    """拖拽区域"""
    file_dropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setText('📥 拖拽文件到此处\n\n或点击下方「选择文件」按钮')
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 8px;
                padding: 30px;
                font-size: 14px;
                color: #666;
                background-color: #fafafa;
                min-height: 120px;
            }
            QLabel:hover {
                border-color: #4a90d9;
                background-color: #f0f6ff;
            }
        """)
    
    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                QLabel {
                    border: 2px dashed #4a90d9;
                    border-radius: 8px;
                    padding: 30px;
                    font-size: 14px;
                    color: #4a90d9;
                    background-color: #e8f0fe;
                    min-height: 120px;
                }
            """)
    
    def dragLeaveEvent(self, event):
        self._reset_style()
    
    def dropEvent(self, event: QDropEvent):
        self._reset_style()
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path:
                self.file_dropped.emit(file_path)
    
    def _reset_style(self):
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 8px;
                padding: 30px;
                font-size: 14px;
                color: #666;
                background-color: #fafafa;
                min-height: 120px;
            }
        """)


class UploadArea(QWidget):
    """文件上传组件"""
    file_selected = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 拖拽区域
        self.drop_area = DropArea()
        self.drop_area.file_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_area)
        
        # 按钮区域
        btn_layout = QHBoxLayout()
        
        self.select_btn = QPushButton('📂 选择文件')
        self.select_btn.clicked.connect(self._on_select_clicked)
        self.select_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 20px;
                font-size: 13px;
                background-color: #4a90d9;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #357abd; }
        """)
        
        self.clear_btn = QPushButton('🗑️ 清空')
        self.clear_btn.clicked.connect(self._on_clear_clicked)
        self.clear_btn.setEnabled(False)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 20px;
                font-size: 13px;
                background-color: #e0e0e0;
                color: #333;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #d0d0d0; }
            QPushButton:disabled { color: #aaa; }
        """)
        
        btn_layout.addWidget(self.select_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)
        
        # 已选文件信息
        self.file_info = QLabel('')
        self.file_info.setStyleSheet("color: #333; font-size: 12px; padding: 4px 0;")
        layout.addWidget(self.file_info)
        
        self.setLayout(layout)
        
        self._current_file = ''
    
    def _on_file_dropped(self, file_path: str):
        self._set_file(file_path)
    
    def _on_select_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择文件', '', SUPPORTED_FORMATS
        )
        if file_path:
            self._set_file(file_path)
    
    def _on_clear_clicked(self):
        self._current_file = ''
        self.file_info.setText('')
        self.drop_area.setText('📥 拖拽文件到此处\n\n或点击下方「选择文件」按钮')
        self.clear_btn.setEnabled(False)
    
    def _set_file(self, file_path: str):
        self._current_file = file_path
        fname = Path(file_path).name
        fsize = os.path.getsize(file_path)
        size_str = self._format_size(fsize)
        
        self.drop_area.setText(f'✅ 已选择文件: {fname}')
        self.file_info.setText(f'📄 {fname}  ({size_str})')
        self.clear_btn.setEnabled(True)
        
        self.file_selected.emit(file_path)
    
    def _format_size(self, size: int) -> str:
        for unit in ['B', 'KB', 'MB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} GB"
    
    def get_file_path(self) -> str:
        return self._current_file
```

- [ ] **Step 2: 实现 preview_panel.py（结果预览面板）**

```python
"""字段抽取结果预览面板"""
from typing import Dict, Any

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTableWidget,
    QTableWidgetItem, QHeaderView, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QBrush


FIELD_LABELS = {
    'personal': {
        'name': '姓名', 'id_number': '证件号码', 'report_time': '报告时间',
        'credit_card_count': '信用卡账户数', 'loan_count': '贷款账户数',
        'overdue_count': '逾期账户数', 'total_balance': '余额',
        'settled_count': '已结清账户数', 'anomaly_notes': '异常备注',
    },
    'corporate': {
        'company_name': '企业名称', 'credit_code': '统一社会信用代码',
        'report_time': '报告时间', 'unsettled_institutions': '未结清机构数',
        'total_balance': '余额', 'short_term_loan': '短期借款',
        'medium_long_term_loan': '中长期借款', 'guarantee_info': '担保信息',
        'public_info': '公共信息',
    },
    'tax': {
        'tax_registration': '纳税登记状态', 'has_penalty': '是否有滞纳金',
        'tax_arrears': '欠税金额', 'invoice_3year': '近三年开票汇总',
        'tax_revenue_3year': '近三年纳税数据', 'tax_anomaly': '税务异常说明',
    },
}


class PreviewPanel(QWidget):
    """结果预览面板"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_type = 'personal'
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        title = QLabel('📋 字段抽取结果')
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #333; padding: 4px 0;")
        layout.addWidget(title)
        
        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(['字段名称', '字段值', '置信度', '状态'])
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 12px;
            }
            QHeaderView::section {
                background-color: #f0f0f0;
                padding: 6px;
                border: none;
                border-bottom: 1px solid #ddd;
                font-weight: bold;
            }
        """)
        
        layout.addWidget(self.table)
        
        self.status_label = QLabel('等待处理...')
        self.status_label.setStyleSheet("color: #999; font-size: 11px; padding: 4px 0;")
        layout.addWidget(self.status_label)
        
        self.setLayout(layout)
    
    def show_fields(self, fields: Dict[str, Any], report_type: str):
        """显示字段抽取结果"""
        self._current_type = report_type
        labels = FIELD_LABELS.get(report_type, {})
        
        # 筛选有值的字段
        visible_fields = [(k, v) for k, v in fields.items() if k in labels]
        
        self.table.setRowCount(len(visible_fields))
        
        for row, (key, data) in enumerate(visible_fields):
            label = labels.get(key, key)
            value = data.get('value', '')
            confidence = data.get('confidence', 0)
            note = data.get('note', '')
            
            # 字段名
            name_item = QTableWidgetItem(label)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row, 0, name_item)
            
            # 字段值
            display = value if value else '[未识别]'
            value_item = QTableWidgetItem(display)
            value_item.setFlags(value_item.flags() & ~Qt.ItemIsEditable)
            
            if not value:
                value_item.setBackground(QBrush(QColor(0xFF, 0xF2, 0xCC)))
            elif '⚠️' in str(value):
                value_item.setBackground(QBrush(QColor(0xFC, 0xE4, 0xEC)))
            
            self.table.setItem(row, 1, value_item)
            
            # 置信度
            conf_text = f'{confidence:.0%}' if isinstance(confidence, float) else str(confidence)
            conf_item = QTableWidgetItem(conf_text)
            conf_item.setFlags(conf_item.flags() & ~Qt.ItemIsEditable)
            
            if isinstance(confidence, float) and confidence < 0.6:
                conf_item.setBackground(QBrush(QColor(0xFC, 0xE4, 0xEC)))
            elif isinstance(confidence, float) and confidence < 0.9:
                conf_item.setBackground(QBrush(QColor(0xFF, 0xF2, 0xCC)))
            
            self.table.setItem(row, 2, conf_item)
            
            # 状态
            status_text = note if note else ('✅ 已识别' if value else '⏳ 待识别')
            status_item = QTableWidgetItem(status_text)
            status_item.setFlags(status_item.flags() & ~Qt.ItemIsEditable)
            
            if note == '未识别':
                status_item.setForeground(QBrush(QColor(0xCC, 0x66, 0x00)))
            elif '⚠️' in str(value):
                status_item.setForeground(QBrush(QColor(0xCC, 0x33, 0x00)))
            
            self.table.setItem(row, 3, status_item)
    
    def clear(self):
        """清空预览"""
        self.table.setRowCount(0)
        self.status_label.setText('等待处理...')
    
    def set_status(self, text: str):
        self.status_label.setText(text)
```

- [ ] **Step 3: 实现 log_panel.py（日志面板）**

```python
"""处理日志面板"""
from datetime import datetime

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QTextEdit, QLabel, QHBoxLayout, QPushButton
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor, QTextCharFormat


LOG_COLORS = {
    'INFO': QColor(0x33, 0x33, 0x33),
    'WARN': QColor(0xCC, 0x88, 0x00),
    'ERROR': QColor(0xCC, 0x33, 0x00),
    'SUCCESS': QColor(0x00, 0x88, 0x00),
}


class LogPanel(QWidget):
    """日志显示面板"""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()
    
    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        
        # 标题栏
        header = QHBoxLayout()
        title = QLabel('📝 处理日志')
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #333;")
        header.addWidget(title)
        header.addStretch()
        
        clear_btn = QPushButton('清空')
        clear_btn.clicked.connect(self.clear_log)
        clear_btn.setStyleSheet("""
            QPushButton {
                padding: 2px 12px;
                font-size: 11px;
                background: #f0f0f0;
                border: 1px solid #ddd;
                border-radius: 3px;
            }
            QPushButton:hover { background: #e0e0e0; }
        """)
        header.addWidget(clear_btn)
        layout.addLayout(header)
        
        # 日志文本框
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                font-family: monospace;
                font-size: 11px;
                background-color: #fafafa;
                padding: 6px;
            }
        """)
        layout.addWidget(self.log_area)
        
        self.setLayout(layout)
    
    def append_log(self, level: str, message: str):
        """添加日志"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        color = LOG_COLORS.get(level, LOG_COLORS['INFO'])
        
        fmt = QTextCharFormat()
        fmt.setForeground(color)
        
        prefix = {'INFO': ' ℹ️', 'WARN': ' ⚠️', 'ERROR': ' ❌', 'SUCCESS': ' ✅'}.get(level, '')
        
        self.log_area.setCurrentCharFormat(fmt)
        self.log_area.append(f'[{timestamp}]{prefix} {message}')
        
        # 自动滚动到底部
        scrollbar = self.log_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())
    
    def clear_log(self):
        self.log_area.clear()
```

- [ ] **Step 4: 提交**

```bash
git add -A && git commit -m "feat: add UI widgets (upload, preview, log)"
```

---

### Task 11: 主窗口

**Files:**
- Create: `app/main_window.py`

- [ ] **Step 1: 实现 main_window.py**

```python
"""主窗口"""
import os
from pathlib import Path

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QLabel, QRadioButton, QButtonGroup, QProgressBar,
    QMessageBox, QFrame, QSplitter, QApplication
)
from PyQt5.QtCore import Qt, QThread
from PyQt5.QtGui import QFont, QIcon, QPixmap

from app.pipeline.worker import ProcessingWorker
from app.widgets.upload_area import UploadArea
from app.widgets.preview_panel import PreviewPanel
from app.widgets.log_panel import LogPanel
from app.pipeline.file_handler import is_supported


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self._worker = None
        self._current_file = ''
        self._current_report_type = 'personal'
        self._init_ui()
    
    def _init_ui(self):
        self.setWindowTitle('征信报告OCR识别与生成工具 v1.0')
        self.setMinimumSize(900, 700)
        
        # 全局样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f5f5;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #ddd;
                border-radius: 6px;
                margin-top: 8px;
                padding-top: 16px;
                background-color: white;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
                color: #333;
            }
        """)
        
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(12, 12, 12, 12)
        
        # ====== 顶部：报告类型 + 上传区域 ======
        top_layout = QHBoxLayout()
        
        # 报告类型选择
        type_group = QGroupBox('报告类型')
        type_layout = QVBoxLayout()
        
        self.type_group = QButtonGroup(self)
        self.radio_personal = QRadioButton('个人征信报告')
        self.radio_corporate = QRadioButton('企业征信报告')
        self.radio_tax = QRadioButton('水母/税务报告')
        
        self.radio_personal.setChecked(True)
        self.radio_personal.toggled.connect(self._on_type_changed)
        
        self.type_group.addButton(self.radio_personal, 1)
        self.type_group.addButton(self.radio_corporate, 2)
        self.type_group.addButton(self.radio_tax, 3)
        
        type_layout.addWidget(self.radio_personal)
        type_layout.addWidget(self.radio_corporate)
        type_layout.addWidget(self.radio_tax)
        type_layout.addStretch()
        type_group.setLayout(type_layout)
        type_group.setFixedWidth(180)
        
        top_layout.addWidget(type_group)
        
        # 上传区域
        upload_group = QGroupBox('文件上传')
        upload_layout = QVBoxLayout()
        self.upload_area = UploadArea()
        self.upload_area.file_selected.connect(self._on_file_selected)
        upload_layout.addWidget(self.upload_area)
        upload_group.setLayout(upload_layout)
        
        top_layout.addWidget(upload_group, 1)
        main_layout.addLayout(top_layout)
        
        # ====== 进度条 + 操作按钮 ======
        progress_layout = QHBoxLayout()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(28)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #ddd;
                border-radius: 4px;
                text-align: center;
                font-size: 12px;
                background-color: #f0f0f0;
            }
            QProgressBar::chunk {
                background-color: #4a90d9;
                border-radius: 3px;
            }
        """)
        
        self.start_btn = QPushButton('▶ 开始处理')
        self.start_btn.clicked.connect(self._on_start)
        self.start_btn.setEnabled(False)
        self.start_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 24px;
                font-size: 14px;
                font-weight: bold;
                background-color: #4a90d9;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #357abd; }
            QPushButton:disabled { background-color: #b0c4de; }
        """)
        
        self.output_btn = QPushButton('📂 输出目录')
        self.output_btn.clicked.connect(self._on_open_output)
        self.output_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                font-size: 12px;
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #e0e0e0; }
        """)
        
        self.reset_btn = QPushButton('🔄 重置')
        self.reset_btn.clicked.connect(self._on_reset)
        self.reset_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                font-size: 12px;
                background-color: #f0f0f0;
                border: 1px solid #ddd;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #e0e0e0; }
        """)
        
        progress_layout.addWidget(self.progress_bar, 1)
        progress_layout.addWidget(self.start_btn)
        progress_layout.addWidget(self.output_btn)
        progress_layout.addWidget(self.reset_btn)
        main_layout.addLayout(progress_layout)
        
        # ====== 中间：结果预览 + 日志 ======
        splitter = QSplitter(Qt.Vertical)
        
        self.preview_panel = PreviewPanel()
        splitter.addWidget(self.preview_panel)
        
        self.log_panel = LogPanel()
        splitter.addWidget(self.log_panel)
        
        splitter.setSizes([350, 200])
        main_layout.addWidget(splitter, 1)
        
        # ====== 状态栏 ======
        self.statusBar().showMessage('就绪')
    
    def _on_type_changed(self):
        if self.radio_personal.isChecked():
            self._current_report_type = 'personal'
        elif self.radio_corporate.isChecked():
            self._current_report_type = 'corporate'
        else:
            self._current_report_type = 'tax'
        
        self.log_panel.append_log('INFO', f'切换报告类型: {["个人征信","企业征信","水母报告"][["personal","corporate","tax"].index(self._current_report_type)]}')
    
    def _on_file_selected(self, file_path: str):
        self._current_file = file_path
        self.preview_panel.clear()
        
        if is_supported(file_path):
            self.start_btn.setEnabled(True)
            self.log_panel.append_log('SUCCESS', f'文件加载成功: {Path(file_path).name}')
        else:
            self.start_btn.setEnabled(False)
            self.log_panel.append_log('ERROR', f'不支持的文件格式: {Path(file_path).suffix}')
            QMessageBox.warning(self, '格式不支持', f'不支持的文件格式: {Path(file_path).suffix}\n\n支持的格式: PDF, PNG, JPG, BMP, DOCX, XLSX')
    
    def _on_start(self):
        if not self._current_file:
            return
        
        # 禁用按钮
        self.start_btn.setEnabled(False)
        self.reset_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.preview_panel.clear()
        self.log_panel.append_log('INFO', '=' * 40)
        self.log_panel.append_log('INFO', f'开始处理: {Path(self._current_file).name}')
        self.log_panel.append_log('INFO', f'报告类型: {["个人征信","企业征信","水母报告"][["personal","corporate","tax"].index(self._current_report_type)]}')
        
        # 创建工作线程
        self._worker = ProcessingWorker(self._current_file, self._current_report_type)
        
        # 连接信号
        self._worker.signals.progress.connect(self._on_progress)
        self._worker.signals.log.connect(self._on_log)
        self._worker.signals.result_ready.connect(self._on_result)
        self._worker.signals.report_ready.connect(self._on_report_ready)
        self._worker.signals.error.connect(self._on_error)
        self._worker.signals.finished.connect(self._on_finished)
        
        self.statusBar().showMessage('正在处理...')
        self._worker.start()
    
    def _on_progress(self, value: int, status: str):
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(f'{status} ({value}%)')
    
    def _on_log(self, level: str, message: str):
        self.log_panel.append_log(level, message)
    
    def _on_result(self, fields: dict):
        self.preview_panel.show_fields(fields, self._current_report_type)
    
    def _on_report_ready(self, file_path: str):
        self.log_panel.append_log('SUCCESS', f'报告已生成: {file_path}')
        self.statusBar().showMessage(f'处理完成 → {file_path}')
    
    def _on_error(self, error_msg: str):
        self.statusBar().showMessage('处理失败')
        QMessageBox.critical(self, '处理错误', f'处理过程中发生错误:\n{error_msg}')
    
    def _on_finished(self):
        self.start_btn.setEnabled(True)
        self.reset_btn.setEnabled(True)
    
    def _on_open_output(self):
        output_dir = os.path.join(os.getcwd(), 'output')
        if os.path.exists(output_dir):
            os.system(f'open "{output_dir}"')
        else:
            QMessageBox.information(self, '提示', '输出目录尚不存在，请先处理一个文件。')
    
    def _on_reset(self):
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat('就绪')
        self.preview_panel.clear()
        self.preview_panel.set_status('等待处理...')
        self.start_btn.setEnabled(bool(self._current_file))
        self.reset_btn.setEnabled(True)
        self.statusBar().showMessage('已重置')
    
    def closeEvent(self, event):
        """关闭窗口时确保线程终止"""
        if self._worker and self._worker.isRunning():
            self._worker.quit()
            self._worker.wait(3000)
        event.accept()
```

- [ ] **Step 2: 提交**

```bash
git add -A && git commit -m "feat: add main window"
```

---

### Task 12: 应用入口

**Files:**
- Create: `main.py`

- [ ] **Step 1: 实现 main.py**

```python
#!/usr/bin/env python3
"""征信报告OCR识别与生成工具 - 入口"""
import sys
import os

# 确保项目根目录在 path 中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont

from app.main_window import MainWindow


def main():
    # 高DPI支持
    QApplication.setAttribute(0x10001)  # Qt.AA_EnableHighDpiScaling
    QApplication.setAttribute(0x10002)  # Qt.AA_UseHighDpiPixmaps
    
    app = QApplication(sys.argv)
    
    # 设置全局字体
    font = QFont('Microsoft YaHei', 9)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
```

- [ ] **Step 2: 验证应用可以启动**

```bash
python3 main.py
```

Expected: 窗口启动成功，界面布局完整，按钮可用

- [ ] **Step 3: 提交**

```bash
git add -A && git commit -m "feat: add main entry point"
```

---

### Task 13: 安装依赖与集成验证

- [ ] **Step 1: 安装依赖**

```bash
pip install -r requirements.txt
```

- [ ] **Step 2: 启动应用验证**

```bash
python3 main.py
```

Expected: 窗口正常启动，所有功能可用

- [ ] **Step 3: 完整流程测试**
  1. 选择报告类型（个人征信）
  2. 上传一个 PDF 或图片文件
  3. 点击"开始处理"
  4. 观察进度条和日志更新
  5. 查看字段预览结果
  6. 确认 Word 报告生成到 output/ 目录

- [ ] **Step 4: 最终提交**

```bash
git add -A && git commit -m "feat: complete credit report OCR tool v1.0"
```
