"""图片预处理工具：为 OCR 提供图片预处理功能"""
import cv2
import numpy as np


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
    """矫正图片倾斜（输入应为二值化后的图片）"""
    if image is None or image.size == 0:
        return image

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
