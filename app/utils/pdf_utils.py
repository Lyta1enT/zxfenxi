"""PDF 处理工具：从 PDF 中提取文本或转为图片供 OCR 使用"""
from pathlib import Path
from typing import List, Tuple

import fitz  # PyMuPDF


def extract_text_from_pdf(pdf_path: str) -> str:
    """从 PDF 中直接提取文本"""
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise ValueError(f"无法打开 PDF 文件 {pdf_path}: {e}")

    texts = []
    for page in doc:
        texts.append(page.get_text())
    doc.close()
    return "\n".join(texts)


def pdf_to_images(pdf_path: str, dpi: int = 200) -> List[str]:
    """将 PDF 每页转换为图片，返回图片路径列表

    使用 PyMuPDF 进行转换，比 pdf2image 更快且无需 poppler
    """
    try:
        doc = fitz.open(pdf_path)
    except Exception as e:
        raise ValueError(f"无法打开 PDF 文件 {pdf_path}: {e}")

    image_paths = []
    base_path = str(Path(pdf_path).with_suffix(""))

    for page_num, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        img_path = f"{base_path}_page_{page_num + 1}.png"
        try:
            pix.save(img_path)
        except Exception as e:
            raise IOError(f"保存 PDF 页面图片失败 {img_path}: {e}")
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
