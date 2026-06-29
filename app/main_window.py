"""主窗口 - 支持多文件上传和批量处理"""
import os
from pathlib import Path
from typing import List

from PyQt5.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGroupBox,
    QPushButton, QLabel, QRadioButton, QButtonGroup, QProgressBar,
    QMessageBox, QSplitter, QListWidget, QListWidgetItem
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
        self._current_files: List[str] = []
        self._current_report_type = 'personal'
        self._report_paths: List[str] = []
        self._init_ui()

    def _init_ui(self):
        self.setWindowTitle('征信报告OCR识别与生成工具 v1.1')
        self.setMinimumSize(960, 760)

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

        type_layout.addStretch()
        type_group.setLayout(type_layout)
        type_group.setFixedWidth(180)

        top_layout.addWidget(type_group)

        # 多文件上传区域
        upload_group = QGroupBox('文件上传（支持多文件）')
        upload_layout = QVBoxLayout()
        self.upload_area = UploadArea()
        self.upload_area.files_changed.connect(self._on_files_changed)
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

        self.start_btn = QPushButton('\u25b6 开始批量处理')
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

        # ====== 中间：文件处理状态 + 结果预览 + 日志 ======
        middle_layout = QSplitter(Qt.Horizontal)

        # 左侧：文件处理状态列表
        status_panel = QWidget()
        status_layout = QVBoxLayout(status_panel)
        status_layout.setContentsMargins(0, 0, 0, 0)

        status_title = QLabel('\U0001f4cb 处理状态')
        status_title.setStyleSheet("font-size: 14px; font-weight: bold; color: #333; padding: 4px 0;")
        status_layout.addWidget(status_title)

        self.status_list = QListWidget()
        self.status_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 12px;
                background: white;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #f0f0f0;
            }
        """)
        status_layout.addWidget(self.status_list)
        status_panel.setFixedWidth(280)

        middle_layout.addWidget(status_panel)

        # 右侧：结果预览
        self.preview_panel = PreviewPanel()
        middle_layout.addWidget(self.preview_panel)

        middle_layout.setSizes([280, 400])
        main_layout.addWidget(middle_layout, 3)

        # ====== 底部：日志 ======
        self.log_panel = LogPanel()
        main_layout.addWidget(self.log_panel, 2)

        self.statusBar().showMessage('就绪')

        # 延迟设置默认选中，确保所有组件已创建
        self.radio_buttons[0].setChecked(True)

    def _on_type_changed(self):
        checked_id = self.type_group.checkedId()
        if 0 <= checked_id < len(self.REPORT_TYPES):
            self._current_report_type = self.REPORT_TYPES[checked_id]
            self.log_panel.append_log('INFO', f'切换报告类型: {self.REPORT_NAMES[checked_id]}')

    def _on_files_changed(self, files: List[str]):
        """文件列表变化"""
        self._current_files = files
        self.start_btn.setEnabled(len(files) > 0)
        self._refresh_status_list()

        if files:
            self.log_panel.append_log('SUCCESS', f'已选择 {len(files)} 个文件')

    def _refresh_status_list(self):
        """刷新文件状态列表"""
        self.status_list.clear()
        for fpath in self._current_files:
            name = Path(fpath).name
            item = QListWidgetItem(f'\u23f3 等待处理  -  {name}')
            self.status_list.addItem(item)

    def _on_start(self):
        """开始批量处理"""
        if not self._current_files:
            return

        # 验证所有文件格式
        unsupported = [f for f in self._current_files if not is_supported(f)]
        if unsupported:
            names = '\n'.join(Path(f).name for f in unsupported)
            QMessageBox.warning(self, '格式不支持', f'以下文件格式不支持:\n{names}')
            return

        self.start_btn.setEnabled(False)
        self.reset_btn.setEnabled(False)
        self.progress_bar.setValue(0)
        self.preview_panel.clear()
        self._report_paths = []

        type_idx = self.REPORT_TYPES.index(self._current_report_type)
        self.log_panel.append_log('INFO', '=' * 50)
        self.log_panel.append_log('INFO',
            f'开始批量处理: {len(self._current_files)} 个文件, '
            f'类型: {self.REPORT_NAMES[type_idx]}')

        # 更新状态列表
        self.status_list.clear()
        for fpath in self._current_files:
            name = Path(fpath).name
            item = QListWidgetItem(f'\U0001f4e6 {name}')
            self.status_list.addItem(item)

        # 创建工作线程（合并模式：多个文件生成1份报告）
        self._worker = ProcessingWorker(self._current_files, self._current_report_type)

        self._worker.signals.progress.connect(self._on_progress)
        self._worker.signals.file_progress.connect(self._on_file_progress)
        self._worker.signals.log.connect(self._on_log)
        self._worker.signals.all_done.connect(self._on_all_done)
        self._worker.signals.error.connect(self._on_file_error)

        self.statusBar().showMessage('正在批量处理...')
        self._worker.start()

    def _on_progress(self, value: int, status: str):
        self.progress_bar.setValue(value)
        self.progress_bar.setFormat(f'{status} ({value}%)')

    def _on_file_progress(self, idx: int, total: int, filename: str):
        """更新某个文件的状态"""
        if idx < self.status_list.count():
            item = self.status_list.item(idx)
            if item:
                item.setText(f'\U0001f504 处理中... {filename}')
                self.status_list.scrollToItem(item)

        self.log_panel.append_log('INFO', f'--- 处理文件 [{idx+1}/{total}]: {filename} ---')
        self.statusBar().showMessage(f'正在处理 [{idx+1}/{total}]: {filename}')

    def _on_log(self, level: str, message: str):
        self.log_panel.append_log(level, message)

    def _on_file_error(self, idx: int, error_msg: str):
        """单个文件处理出错"""
        if idx < self.status_list.count():
            item = self.status_list.item(idx)
            if item:
                text = item.text()
                item.setText(f'\u274c 处理失败')

    def _on_all_done(self, report_path: str, source_files: List[str]):
        """全部处理完成，合并生成1份报告"""
        success = len([f for f in source_files if Path(f).exists()])
        total = len(self._current_files)
        self.progress_bar.setValue(100)
        self.progress_bar.setFormat(f'已完成 ({total}个文件合并为1份报告)')

        self.log_panel.append_log('SUCCESS',
            f'✅ 综合报告已生成: {report_path}')
        self.log_panel.append_log('INFO',
            f'   源文件: {len(source_files)}个 → 1份报告')

        self.statusBar().showMessage(f'完成: {Path(report_path).name}')
        self.start_btn.setEnabled(True)
        self.reset_btn.setEnabled(True)
        self._report_paths = [report_path]

        QMessageBox.information(self, '处理完成',
            f'✅ 所有文件已合并为1份报告\n\n'
            f'源文件: {total}个\n'
            f'输出: {Path(report_path).name}\n\n'
            f'已保存到 output/ 目录')

    def _on_open_output(self):
        output_dir = os.path.join(os.getcwd(), 'output')
        if os.path.exists(output_dir):
            os.system(f'open "{output_dir}"')
        else:
            QMessageBox.information(self, '提示', '输出目录尚不存在，请先处理文件。')

    def _on_reset(self):
        """重置所有状态"""
        if self._worker and self._worker.isRunning():
            self._worker.quit()
            self._worker.wait(2000)

        self.progress_bar.setValue(0)
        self.progress_bar.setFormat('就绪')
        self.preview_panel.clear()
        self.preview_panel.set_status('等待处理...')
        self.status_list.clear()
        self._report_paths = []
        self._current_files = self.upload_area.get_files()
        self.start_btn.setEnabled(len(self._current_files) > 0)
        self.reset_btn.setEnabled(True)
        self._refresh_status_list()
        self.statusBar().showMessage('已重置')

    def closeEvent(self, event):
        if self._worker and self._worker.isRunning():
            self._worker.quit()
            self._worker.wait(3000)
        event.accept()
