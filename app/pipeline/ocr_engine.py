"""PaddleOCR 引擎封装"""
import os
from typing import List, Dict, Any, Optional

import numpy as np
import cv2

from app.utils.image_utils import preprocess_image, deskew_image


class OCREngine:
    """封装 PaddleOCR PP-OCRv5，提供统一接口"""

    def __init__(self, use_angle_cls: bool = True, lang: str = 'ch',
                 ocr_version: str = 'PP-OCRv5', enable_mkldnn: bool = True):
        self._ocr = None
        self.use_angle_cls = use_angle_cls
        self.lang = lang
        self.ocr_version = ocr_version
        self.enable_mkldnn = enable_mkldnn
        self._initialized = False

    def _lazy_init(self):
        """延迟初始化 PaddleOCR（避免 import 耗时阻塞 UI）"""
        if not self._initialized:
            from paddleocr import PaddleOCR
            self._ocr = PaddleOCR(
                lang=self.lang,
                ocr_version=self.ocr_version,
                use_angle_cls=self.use_angle_cls,
            )
            self._initialized = True
    
    def recognize_image(self, image_path: str, page: int = 1) -> List[Dict[str, Any]]:
        """识别单张图片，返回 OCR 结果列表"""
        self._lazy_init()
        
        processed = preprocess_image(image_path)
        result = self._ocr.ocr(processed, cls=True)
        
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
