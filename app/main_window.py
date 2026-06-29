"""主窗口"""
import os
from pathlib import Path

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QLabel, QRadioButton, QButtonGroup, QProgressBar,
    QMessageBox, QSplitter, QApplication
)
from PyQt5.QtCore import Qt

from app.pipeline.worker import ProcessingWorker
from app.widgets.upload_area import UploadArea
from app.widgets.preview_panel import PreviewPanel
from app.widgets.log_panel import LogPanel
from app.pipeline.file_handler import is_supported


class MainWindow(QMainWindow):
    """主窗口"""

    REPORT_TYPES = ['personal', 'corporate', 'tax']
    REPORT_NAMES = ['个人征信报告', '企业征信报告', '水母/税务报告']

    def __init__(self):
        super().__init__()
        self._worker = None
        self._current_file = ''
        self._current_report_type = 'personal'
        self._init_ui()

    def _init_ui(self):
        self.setWindowTitle('征信报告OCR识别与生成工具 v1.0')
        self.setMinimumSize(900, 700)

        self.setStyleSheet("""
            QMainWindow { background-color: #f5f5f5; }
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
        self.radio_buttons = []
        for i, name in enumerate(self.REPORT_NAMES):
            rb = QRadioButton(name)
            rb.toggled.connect(self._on_type_changed)
            self.type_group.addButton(rb, i)
            type_layout.addWidget(rb)
            self.radio_buttons.append(rb)

        self.radio_buttons[0].setChecked(True)
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

        self.start_btn = QPushButton('\u25b6 开始处理')
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

        self.output_btn = QPushButton('\U0001f4c2 输出目录')
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

        self.reset_btn = QPushButton('\U0001f504 重置')
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

        self.statusBar().showMessage('就绪')

    def _on_type_changed(self):
        checked_id = self.type_group.checkedId()
        if 0 <= checked_id < len(self.REPORT_TYPES):
            self._current_report_type = self.REPORT_TYPES[checked_id]
            self.log_panel.append_log('INFO', f'切换报告类型: {self.REPORT_NAMES[checked_id]}')

    def _on_file_selected(self, file_path: str):
        self._current_file = file_path
        self.preview_panel.clear()

        if is_supported(file_path):
            self.start_btn.setEnabled(True)
            self.log_panel.append_log('SUCCESS', f'文件加载成功: {Path(file_path).name}')
        else:
            self.start_btn.setEnabled(False)
            ext = Path(file_path).suffix
            self.log_panel.append_log('ERROR', f'不支持的文件格式: {ext}')
            QMessageBox.warning(
                self, '格式不支持',
                f'不支持的文件格式: {ext}\n\n支持的格式: PDF, PNG, JPG, BMP, DOCX, XLSX'
            )

    def _on_start(self):
        if not self._current_file:
            return

        self.start_btn.setEnabled(False)
        self.reset_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.preview_panel.clear()

        type_idx = self.REPORT_TYPES.index(self._current_report_type)
        self.log_panel.append_log('INFO', '=' * 40)
        self.log_panel.append_log('INFO', f'开始处理: {Path(self._current_file).name}')
        self.log_panel.append_log('INFO', f'报告类型: {self.REPORT_NAMES[type_idx]}')

        self._worker = ProcessingWorker(self._current_file, self._current_report_type)

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
        if self._worker and self._worker.isRunning():
            self._worker.quit()
            self._worker.wait(3000)
        event.accept()
