"""后台工作线程 - 多文件合并为1份报告"""
import traceback
from pathlib import Path
from typing import List, Dict, Any, Optional
from datetime import datetime

from PyQt5.QtCore import QThread, pyqtSignal

from app.pipeline.file_handler import process_file, get_file_type
from app.pipeline.ocr_engine import OCREngine
from app.pipeline.field_extractor import extract_fields, get_field_definitions
from app.pipeline.report_generator import generate_report


class ProcessingSignals(QThread):
    """工作线程信号"""
    progress = pyqtSignal(int, str)       # (进度百分比, 状态描述)
    file_progress = pyqtSignal(int, int, str)  # (当前文件索引, 总文件数, 文件名)
    log = pyqtSignal(str, str)            # (级别, 消息)
    all_done = pyqtSignal(str, object) # (报告路径, 源文件名列表)
    error = pyqtSignal(int, str)          # (文件索引, 错误信息)


def _merge_fields(all_fields: List[Dict[str, Any]]) -> Dict[str, Any]:
    """合并多个文件提取的字段，优先取非空值"""
    merged = {}
    for fields in all_fields:
        for key, data in fields.items():
            if key not in merged:
                merged[key] = data
            else:
                # 已有值但为空，用新值覆盖
                existing = merged[key]
                if isinstance(existing, dict) and isinstance(data, dict):
                    if not existing.get('value', '') and data.get('value', ''):
                        merged[key] = data
                    elif existing.get('value', '') and not data.get('value', ''):
                        pass  # 保留现有
                    elif data.get('value', ''):
                        # 都有值，追加（用换行分隔）
                        existing_val = existing.get('value', '')
                        new_val = data.get('value', '')
                        if new_val not in existing_val:
                            merged[key] = {
                                'value': existing_val + '\n' + new_val,
                                'confidence': max(
                                    existing.get('confidence', 0),
                                    data.get('confidence', 0)
                                ),
                                'page': existing.get('page', 0),
                                'note': '多源合并',
                            }
    return merged


def _merge_text(all_texts: List[str]) -> str:
    """合并多个文件的原始文本"""
    seen = set()
    merged = []
    for text in all_texts:
        for line in text.split('\n'):
            line = line.strip()
            if line and line not in seen:
                seen.add(line)
                merged.append(line)
    return '\n'.join(merged)


class ProcessingWorker(QThread):
    """多文件合并处理工作线程"""

    def __init__(self, file_paths: List[str], report_type: str, parent=None):
        super().__init__(parent)
        self.file_paths = file_paths
        self.report_type = report_type
        self.signals = ProcessingSignals()
        self._ocr_engine = OCREngine()

    def run(self):
        """处理所有文件，合并为1份报告"""
        total = len(self.file_paths)
        all_fields = []
        all_text = []

        for idx, file_path in enumerate(self.file_paths):
            filename = Path(file_path).name
            self.signals.file_progress.emit(idx, total, filename)

            try:
                fields, raw_text = self._process_single_file(idx, file_path, total)
                all_fields.append(fields)
                all_text.append(raw_text)
            except Exception as e:
                self.signals.log.emit('ERROR', f'[{filename}] 处理失败: {str(e)}')
                self.signals.error.emit(idx, str(e))

        # 合并所有字段和文本 → 生成1份报告
        if all_fields:
            self.signals.progress.emit(90, '合并数据生成最终报告...')
            self.signals.log.emit('INFO', f'合并 {len(all_fields)} 个文件的数据...')

            merged_fields = _merge_fields(all_fields)
            merged_text = _merge_text(all_text)

            # 用第一个源文件名作为报告名称
            first_name = Path(self.file_paths[0]).stem
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            combined_name = f'{first_name}_合并报告_{timestamp}'

            output_path = generate_report(
                fields=merged_fields,
                report_type=self.report_type,
                source_filename=combined_name,
                raw_text=merged_text,
            )

            self.signals.log.emit('SUCCESS',
                f'✅ 综合报告已生成: {output_path}')
            self.signals.progress.emit(100, f'完成 ({len(all_fields)}个文件合并)')
            self.signals.all_done.emit(output_path, self.file_paths)
        else:
            self.signals.log.emit('ERROR', '没有成功处理任何文件')

    def _process_single_file(self, idx: int, file_path: str,
                              total: int) -> tuple:
        """处理单个文件，返回 (fields, raw_text)"""
        filename = Path(file_path).name
        file_num = idx + 1
        base_progress = int(idx * 100 / total)

        def emit_progress(sub_pct: int, msg: str):
            pct = min(base_progress + int(sub_pct / total), 99)
            self.signals.progress.emit(pct, f'[{file_num}/{total}] {msg}')

        # Step 1: 文件处理
        emit_progress(0, f'分析: {filename}')
        self.signals.log.emit('INFO', f'[{file_num}/{total}] {filename}')

        file_result = process_file(file_path)
        image_paths = file_result['image_paths']
        raw_text = file_result['text']

        # Step 2: OCR 识别（图片文件）
        ocr_items = []
        if image_paths:
            emit_progress(20, f'OCR: {filename}')
            ocr_items = self._ocr_engine.recognize_batch(image_paths)
            self.signals.log.emit('INFO',
                f'[{file_num}/{total}] OCR {len(ocr_items)} 项')
        elif raw_text:
            self.signals.log.emit('INFO',
                f'[{file_num}/{total}] 文本 {len(raw_text)} 字符')

        # Step 3: 字段抽取
        emit_progress(50, f'抽取: {filename}')
        fields = extract_fields(
            self.report_type,
            ocr_items=ocr_items,
            raw_text=raw_text,
        )
        field_defs = get_field_definitions(self.report_type)
        hit = sum(1 for f in field_defs if fields.get(f['key'], {}).get('value', ''))
        self.signals.log.emit('INFO',
            f'[{file_num}/{total}] 命中 {hit}/{len(field_defs)} 字段')

        return fields, raw_text
