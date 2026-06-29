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

            hit_count = sum(1 for f in field_defs if fields.get(f['key'], {}).get('value', ''))
            total = len(field_defs)
            self.signals.log.emit('INFO', f'字段抽取完成: {hit_count}/{total} 字段命中')

            self.signals.result_ready.emit(fields)
            self.signals.progress.emit(80, '正在生成 Word 报告...')

            # Step 4: 生成报告
            output_path = generate_report(
                fields=fields,
                report_type=self.report_type,
                source_filename=self.file_path,
            )

            self.signals.log.emit('SUCCESS', f'Word 报告已生成: {output_path}')
            self.signals.report_ready.emit(output_path)
            self.signals.progress.emit(100, '处理完成')
            self.signals.finished.emit()

        except Exception as e:
            error_msg = f'处理失败: {str(e)}\n{traceback.format_exc()}'
            self.signals.log.emit('ERROR', error_msg)
            self.signals.error.emit(str(e))
