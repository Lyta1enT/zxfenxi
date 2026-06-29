"""后台工作线程 - 支持多文件批量处理"""
import traceback
from pathlib import Path
from typing import List, Dict, Any

from PyQt5.QtCore import QThread, pyqtSignal

from app.pipeline.file_handler import process_file
from app.pipeline.ocr_engine import OCREngine
from app.pipeline.field_extractor import extract_fields, get_field_definitions
from app.pipeline.report_generator import generate_report


class ProcessingSignals(QThread):
    """工作线程信号"""
    progress = pyqtSignal(int, str)       # (整体进度百分比, 状态描述)
    file_progress = pyqtSignal(int, int, str)  # (当前文件索引, 总文件数, 文件名)
    log = pyqtSignal(str, str)            # (级别, 消息)
    result_ready = pyqtSignal(int, dict)  # (文件索引, 抽取结果)
    report_ready = pyqtSignal(int, str)   # (文件索引, 报告路径)
    all_done = pyqtSignal(list)           # 全部完成，发送报告路径列表
    error = pyqtSignal(int, str)          # (文件索引, 错误信息)


class ProcessingWorker(QThread):
    """多文件批量处理工作线程"""

    def __init__(self, file_paths: List[str], report_type: str, parent=None):
        super().__init__(parent)
        self.file_paths = file_paths
        self.report_type = report_type
        self.signals = ProcessingSignals()
        self._ocr_engine = OCREngine()

    def run(self):
        """依次处理所有文件"""
        total = len(self.file_paths)
        report_paths = []

        for idx, file_path in enumerate(self.file_paths):
            filename = Path(file_path).name
            self.signals.file_progress.emit(idx, total, filename)

            try:
                self._process_single_file(idx, file_path, total, report_paths)
            except Exception as e:
                error_msg = f'[{filename}] 处理失败: {str(e)}'
                self.signals.log.emit('ERROR', error_msg)
                self.signals.error.emit(idx, str(e))
                # 继续处理下一个文件，不中断整个流程

        # 全部完成
        self.signals.all_done.emit(report_paths)
        self.signals.progress.emit(100, f'全部完成 ({len(report_paths)}/{total} 成功)')

    def _process_single_file(self, idx: int, file_path: str,
                             total: int, report_paths: List[str]):
        """处理单个文件"""
        filename = Path(file_path).name
        file_num = idx + 1

        # 当前文件的进度范围: [idx*100/total, (idx+1)*100/total)
        base_progress = int(idx * 100 / total)

        def emit_progress(sub_pct: int, msg: str):
            pct = min(base_progress + int(sub_pct / total), 99)
            self.signals.progress.emit(pct, f'[{file_num}/{total}] {msg}')

        # Step 1: 文件处理
        emit_progress(0, f'分析文件: {filename}')
        self.signals.log.emit('INFO', f'[{file_num}/{total}] 处理文件: {filename}')

        file_result = process_file(file_path)
        file_type = file_result['file_type']
        image_paths = file_result['image_paths']
        raw_text = file_result['text']

        self.signals.log.emit('INFO',
            f'[{file_num}/{total}] 类型: {file_type}, 图片页数: {len(image_paths)}')

        # Step 2: OCR 识别
        emit_progress(20, f'OCR 识别: {filename}')

        ocr_items = []
        if image_paths:
            ocr_items = self._ocr_engine.recognize_batch(image_paths)
            self.signals.log.emit('INFO',
                f'[{file_num}/{total}] OCR 完成: {len(ocr_items)} 个文本块')
        elif raw_text:
            self.signals.log.emit('INFO',
                f'[{file_num}/{total}] 跳过 OCR (已有文本 {len(raw_text)} 字符)')

        # Step 3: 字段抽取
        emit_progress(60, f'字段抽取: {filename}')
        fields = extract_fields(self.report_type, ocr_items, raw_text)
        field_defs = get_field_definitions(self.report_type)

        hit = sum(1 for f in field_defs if fields.get(f['key'], {}).get('value', ''))
        total_fields = len(field_defs)
        self.signals.log.emit('INFO',
            f'[{file_num}/{total}] 字段抽取: {hit}/{total_fields} 命中')

        self.signals.result_ready.emit(idx, fields)

        # Step 4: 报告生成
        emit_progress(80, f'生成报告: {filename}')
        output_path = generate_report(
            fields=fields,
            report_type=self.report_type,
            source_filename=file_path,
            raw_text=raw_text,
        )

        self.signals.log.emit('SUCCESS',
            f'[{file_num}/{total}] 报告已生成: {output_path}')
        self.signals.report_ready.emit(idx, output_path)
        report_paths.append(output_path)
