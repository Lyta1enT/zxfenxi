"""版面分析引擎 - 封装 PP-StructureV3 支持文档结构解析"""
from pathlib import Path
from typing import List, Dict, Any, Optional


class LayoutElement:
    """版面元素"""
    __slots__ = ('type', 'content', 'bbox', 'page', 'confidence')

    def __init__(self, type_tag: str, content: Any, bbox=None,
                 page: int = 1, confidence: float = 1.0):
        self.type = type_tag       # text, table, figure, title, list, formula, etc.
        self.content = content
        self.bbox = bbox
        self.page = page
        self.confidence = confidence

    def __repr__(self):
        return (f"<LayoutElement type={self.type} "
                f"page={self.page} content={str(self.content)[:40]}>")


class LayoutAnalyzer:
    """版面分析器 - 使用 PP-StructureV3 解析文档结构

    支持：
    - 文档版面分析（标题、正文、表格、图片等区域检测）
    - 表格结构化识别（行、列、合并单元格）
    - 文本阅读顺序恢复
    - 多页文档处理

    如果 PP-StructureV3 不可用，会抛出 ImportError，
    调用方应捕获并降级到传统 OCR。
    """

    def __init__(self, lang: str = 'ch'):
        self._engine = None
        self._table_engine = None
        self.lang = lang
        self._initialized = False
        self._available = True

    def is_available(self) -> bool:
        """检查 PP-StructureV3 是否可用"""
        try:
            self._lazy_init()
            return True
        except (ImportError, Exception):
            return False

    def _lazy_init(self):
        """延迟初始化"""
        if not self._initialized:
            try:
                from paddleocr import PPStructureV3
                self._engine = PPStructureV3(lang=self.lang)
                self._initialized = True
            except (ImportError, Exception) as e:
                self._available = False
                raise ImportError(
                    f'PP-StructureV3 不可用: {e}\n'
                    f'请运行: pip install "paddlex[ocr]"'
                )

    def analyze(self, file_path: str) -> Dict[str, Any]:
        """分析文档，返回结构化结果

        Args:
            file_path: PDF 或图片文件路径

        Returns:
            {
                'elements': [LayoutElement, ...],  # 所有版面元素
                'tables': [dict, ...],             # 表格结构化数据
                'text_blocks': [dict, ...],         # 纯文本块（按阅读顺序）
                'raw_text': str,                   # 全部文本拼接
                'pages': int,                      # 页数
            }
        """
        self._lazy_init()

        result = self._engine.predict(file_path)
        return self._parse_result(result)

    def analyze_image(self, image_path: str) -> Dict[str, Any]:
        """分析单张图片"""
        return self.analyze(image_path)

    def _parse_result(self, raw_result) -> Dict[str, Any]:
        """解析 PP-StructureV3 返回结果"""
        elements = []
        tables = []
        text_blocks = []
        raw_text_parts = []
        pages = 1

        if not raw_result:
            return {
                'elements': [],
                'tables': [],
                'text_blocks': [],
                'raw_text': '',
                'pages': 0,
            }

        for item in raw_result:
            # PP-StructureV3 返回格式: {type, bbox, content, img, page, ...}
            elem_type = item.get('type', 'text')
            bbox = item.get('bbox', [])
            page = item.get('page', 1)
            pages = max(pages, page)

            content_raw = item.get('content', '')

            if elem_type == 'table':
                # 表格：content 可能是 dict 或 str
                if isinstance(content_raw, dict):
                    table_data = {
                        'type': 'table',
                        'page': page,
                        'bbox': bbox,
                        'html': content_raw.get('html', ''),
                        'cells': content_raw.get('cells', []),
                    }
                    # 从表格中提取文本
                    for cell_row in table_data['cells']:
                        if isinstance(cell_row, list):
                            for cell in cell_row:
                                if isinstance(cell, dict):
                                    cell_text = str(cell.get('text', '')).strip()
                                    if cell_text:
                                        raw_text_parts.append(cell_text)
                        elif isinstance(cell_row, dict):
                            cell_text = str(cell_row.get('text', '')).strip()
                            if cell_text:
                                raw_text_parts.append(cell_text)

                    tables.append(table_data)
                    elements.append(LayoutElement(
                        'table', table_data, bbox, page
                    ))
                else:
                    # 纯文本形式的表格
                    text = str(content_raw).strip()
                    if text:
                        text_blocks.append({'text': text, 'bbox': bbox, 'page': page})
                        raw_text_parts.append(text)
                        elements.append(LayoutElement('table_text', text, bbox, page))

            elif elem_type in ('text', 'title', 'figure', 'list', 'header', 'footer'):
                text = self._extract_text(content_raw)
                if text:
                    text_blocks.append({'text': text, 'bbox': bbox, 'page': page})
                    raw_text_parts.append(text)

                elements.append(LayoutElement(
                    elem_type, text, bbox, page
                ))

            else:
                text = self._extract_text(content_raw)
                if text:
                    text_blocks.append({'text': text, 'bbox': bbox, 'page': page})
                    raw_text_parts.append(text)
                elements.append(LayoutElement(
                    elem_type, text or str(content_raw), bbox, page
                ))

        return {
            'elements': elements,
            'tables': tables,
            'text_blocks': text_blocks,
            'raw_text': '\n'.join(raw_text_parts),
            'pages': pages,
        }

    @staticmethod
    def _extract_text(content) -> str:
        """从各种格式的 content 中提取文本"""
        if isinstance(content, str):
            return content.strip()
        elif isinstance(content, dict):
            # 可能是 {text: "..."}
            text = content.get('text', '') or content.get('content', '')
            return str(text).strip()
        elif isinstance(content, list):
            texts = []
            for item in content:
                t = LayoutAnalyzer._extract_text(item)
                if t:
                    texts.append(t)
            return ' '.join(texts)
        return str(content).strip() if content else ''
