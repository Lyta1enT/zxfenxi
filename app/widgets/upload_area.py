"""文件上传区域组件 - 支持拖拽和点击选择"""
import os
from pathlib import Path

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QFrame
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent


SUPPORTED_FORMATS = "PDF (*.pdf);;图片 (*.png *.jpg *.jpeg *.bmp);;Word (*.docx);;Excel (*.xlsx);;所有支持文件 (*.pdf *.png *.jpg *.jpeg *.bmp *.docx *.xlsx)"


class DropArea(QLabel):
    """拖拽区域"""
    file_dropped = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setText('\U0001f4e5 拖拽文件到此处\n\n或点击下方「选择文件」按钮')
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 8px;
                padding: 30px;
                font-size: 14px;
                color: #666;
                background-color: #fafafa;
                min-height: 120px;
            }
            QLabel:hover {
                border-color: #4a90d9;
                background-color: #f0f6ff;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                QLabel {
                    border: 2px dashed #4a90d9;
                    border-radius: 8px;
                    padding: 30px;
                    font-size: 14px;
                    color: #4a90d9;
                    background-color: #e8f0fe;
                    min-height: 120px;
                }
            """)

    def dragLeaveEvent(self, event):
        self._reset_style()

    def dropEvent(self, event: QDropEvent):
        self._reset_style()
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path:
                self.file_dropped.emit(file_path)

    def _reset_style(self):
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 8px;
                padding: 30px;
                font-size: 14px;
                color: #666;
                background-color: #fafafa;
                min-height: 120px;
            }
        """)


class UploadArea(QWidget):
    """文件上传组件"""
    file_selected = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_file = ''
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        self.drop_area = DropArea()
        self.drop_area.file_dropped.connect(self._on_file_dropped)
        layout.addWidget(self.drop_area)

        btn_layout = QHBoxLayout()

        self.select_btn = QPushButton('\U0001f4c2 选择文件')
        self.select_btn.clicked.connect(self._on_select_clicked)
        self.select_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 20px;
                font-size: 13px;
                background-color: #4a90d9;
                color: white;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #357abd; }
        """)

        self.clear_btn = QPushButton('\U0001f5d1\ufe0f 清空')
        self.clear_btn.clicked.connect(self._on_clear_clicked)
        self.clear_btn.setEnabled(False)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 20px;
                font-size: 13px;
                background-color: #e0e0e0;
                color: #333;
                border: none;
                border-radius: 4px;
            }
            QPushButton:hover { background-color: #d0d0d0; }
            QPushButton:disabled { color: #aaa; }
        """)

        btn_layout.addWidget(self.select_btn)
        btn_layout.addWidget(self.clear_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        self.file_info = QLabel('')
        self.file_info.setStyleSheet("color: #333; font-size: 12px; padding: 4px 0;")
        layout.addWidget(self.file_info)

        self.setLayout(layout)

    def _on_file_dropped(self, file_path: str):
        self._set_file(file_path)

    def _on_select_clicked(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, '选择文件', '', SUPPORTED_FORMATS
        )
        if file_path:
            self._set_file(file_path)

    def _on_clear_clicked(self):
        self._current_file = ''
        self.file_info.setText('')
        self.drop_area.setText('\U0001f4e5 拖拽文件到此处\n\n或点击下方「选择文件」按钮')
        self.clear_btn.setEnabled(False)

    def _set_file(self, file_path: str):
        self._current_file = file_path
        fname = Path(file_path).name
        fsize = os.path.getsize(file_path)
        size_str = self._format_size(fsize)

        self.drop_area.setText(f'\u2705 已选择文件: {fname}')
        self.file_info.setText(f'\U0001f4c4 {fname}  ({size_str})')
        self.clear_btn.setEnabled(True)

        self.file_selected.emit(file_path)

    def _format_size(self, size: int) -> str:
        for unit in ['B', 'KB', 'MB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} GB"

    def get_file_path(self) -> str:
        return self._current_file
