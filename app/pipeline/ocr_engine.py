"""OCR 引擎 - 支持多后端（EasyOCR / PaddleOCR），自动选择可用引擎"""
import os
from typing import List, Dict, Any, Optional


class OCREngine:
    """OCR识别引擎

    自动检测可用引擎（优先 PaddleOCR，后备 EasyOCR），
    均不可用时抛出 ImportError。
    """

    # 引擎选择顺序
    ENGINE_PRIORITY = ['easyocr', 'paddleocr']

    def __init__(self, lang: str = 'ch', **kwargs):
        self.lang = 'ch_sim' if lang == 'ch' else lang
        self._reader = None
        self._engine_name = None
        self._initialized = False
        self._kwargs = kwargs

    def _lazy_init(self):
        """延迟初始化，自动选择可用引擎"""
        if self._initialized:
            return

        for engine in self.ENGINE_PRIORITY:
            try:
                if engine == 'easyocr':
                    import easyocr
                    self._reader = easyocr.Reader(
                        [self.lang],
                        gpu=False,
                        verbose=False,
                    )
                    self._engine_name = 'easyocr'
                    break

                elif engine == 'paddleocr':
                    from paddleocr import PaddleOCR
                    self._reader = PaddleOCR(
                        lang='ch',
                        use_angle_cls=False,
                        show_log=False,
                    )
                    self._engine_name = 'paddleocr'
                    break

            except (ImportError, Exception) as e:
                continue

        if self._reader is None:
            raise ImportError(
                '没有可用的OCR引擎。请安装: pip install easyocr'
            )

        self._initialized = True

    @property
    def engine_name(self) -> str:
        """当前使用的引擎名"""
        self._lazy_init()
        return self._engine_name

    def recognize_image(self, image_path: str, page: int = 1) -> List[Dict[str, Any]]:
        """识别单张图片，返回 OCR 结果列表

        返回格式: [{text, confidence, bbox, page}, ...]
        """
        self._lazy_init()

        if self._engine_name == 'easyocr':
            return self._recognize_easyocr(image_path, page)
        else:
            return self._recognize_paddleocr(image_path, page)

    def _recognize_easyocr(self, image_path: str, page: int) -> List[Dict[str, Any]]:
        """EasyOCR 识别"""
        result = self._reader.readtext(image_path)
        items = []
        for bbox, text, confidence in result:
            items.append({
                'text': text,
                'confidence': confidence,
                'bbox': bbox,
                'page': page,
            })
        return items

    def _recognize_paddleocr(self, image_path: str, page: int) -> List[Dict[str, Any]]:
        """PaddleOCR 识别"""
        from app.utils.image_utils import preprocess_image
        processed = preprocess_image(image_path)
        result = self._reader.ocr(processed, cls=False)

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
        """仅返回识别的文本内容"""
        results = self.recognize_batch(image_paths)
        return '\n'.join([item['text'] for item in results])
