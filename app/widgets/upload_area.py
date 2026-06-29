"""文件上传区域组件 - 支持多文件拖拽和选择"""
import os
from pathlib import Path
from typing import List

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QFileDialog, QListWidget, QListWidgetItem, QFrame, QAbstractItemView
)
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QDragEnterEvent, QDropEvent, QColor, QIcon


SUPPORTED_EXTENSIONS = (
    "PDF (*.pdf);;图片 (*.png *.jpg *.jpeg *.bmp *.tiff);;"
    "Word (*.docx);;Excel (*.xlsx);;"
    "所有支持文件 (*.pdf *.png *.jpg *.jpeg *.bmp *.tiff *.docx *.xlsx)"
)

FILE_TYPE_ICONS = {
    'pdf': '\U0001f4c4',
    'png': '\U0001f5bc\ufe0f',
    'jpg': '\U0001f5bc\ufe0f',
    'jpeg': '\U0001f5bc\ufe0f',
    'bmp': '\U0001f5bc\ufe0f',
    'tiff': '\U0001f5bc\ufe0f',
    'docx': '\U0001f4dd',
    'xlsx': '\U0001f4ca',
}


def get_file_icon(filename: str) -> str:
    ext = Path(filename).suffix.lower().lstrip('.')
    return FILE_TYPE_ICONS.get(ext, '\U0001f4c4')


def format_size(size: int) -> str:
    for unit in ['B', 'KB', 'MB']:
        if size < 1024:
            return f"{size:.1f} {unit}"
        size /= 1024
    return f"{size:.1f} GB"


class DropArea(QLabel):
    """拖拽区域"""
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignCenter)
        self.setText('\U0001f4e5 拖拽文件到此处\n\n支持多文件同时拖拽')
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 8px;
                padding: 24px;
                font-size: 14px;
                color: #666;
                background-color: #fafafa;
                min-height: 80px;
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
                    padding: 24px;
                    font-size: 14px;
                    color: #4a90d9;
                    background-color: #e8f0fe;
                    min-height: 80px;
                }
            """)

    def dragLeaveEvent(self, event):
        self._reset_style()

    def dropEvent(self, event: QDropEvent):
        self._reset_style()
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path and os.path.isfile(file_path):
                files.append(file_path)
        if files:
            self.files_dropped.emit(files)

    def _reset_style(self):
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 8px;
                padding: 24px;
                font-size: 14px;
                color: #666;
                background-color: #fafafa;
                min-height: 80px;
            }
        """)


class FileItemWidget(QFrame):
    """文件列表中的单个文件项"""
    remove_clicked = pyqtSignal(int)  # index

    def __init__(self, index: int, file_path: str, parent=None):
        super().__init__(parent)
        self.index = index
        self.file_path = file_path
        self._init_ui()

    def _init_ui(self):
        layout = QHBoxLayout()
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(8)

        name = Path(self.file_path).name
        size = os.path.getsize(self.file_path)
        icon = get_file_icon(self.file_path)

        info = QLabel(f'{icon} {name}  ({format_size(size)})')
        info.setStyleSheet("font-size: 12px; color: #333;")
        layout.addWidget(info, 1)

        remove_btn = QPushButton('\u2716')
        remove_btn.setFixedSize(24, 24)
        remove_btn.setStyleSheet("""
            QPushButton {
                border: none;
                border-radius: 12px;
                font-size: 12px;
                color: #999;
                background: transparent;
            }
            QPushButton:hover {
                color: #cc3333;
                background: #fce4ec;
            }
        """)
        remove_btn.clicked.connect(lambda: self.remove_clicked.emit(self.index))
        layout.addWidget(remove_btn)

        self.setLayout(layout)
        self.setStyleSheet("""
            FileItemWidget {
                background: #f8f9fa;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                margin: 2px 0;
            }
            FileItemWidget:hover {
                background: #eef1f5;
                border-color: #4a90d9;
            }
        """)


class UploadArea(QWidget):
    """多文件上传组件"""
    files_changed = pyqtSignal(list)  # 文件列表变化时发送

    def __init__(self, parent=None):
        super().__init__(parent)
        self._files: List[str] = []
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)

        # 拖拽区域
        self.drop_area = DropArea()
        self.drop_area.files_dropped.connect(self._on_files_dropped)
        layout.addWidget(self.drop_area)

        # 按钮区域
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

        self.clear_btn = QPushButton('\U0001f5d1\ufe0f 清空全部')
        self.clear_btn.clicked.connect(self._on_clear_all)
        self.clear_btn.setEnabled(False)
        self.clear_btn.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                font-size: 12px;
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

        self.file_count_label = QLabel('')
        self.file_count_label.setStyleSheet("color: #666; font-size: 12px;")
        btn_layout.addWidget(self.file_count_label)

        layout.addLayout(btn_layout)

        # 文件列表
        self.file_list = QListWidget()
        self.file_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                background: white;
                min-height: 100px;
                max-height: 200px;
            }
            QListWidget::item {
                border: none;
                padding: 0;
            }
        """)
        self.file_list.setVisible(False)
        layout.addWidget(self.file_list)

        self.setLayout(layout)

    def _on_files_dropped(self, files: List[str]):
        self._add_files(files)

    def _on_select_clicked(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, '选择文件（可多选）', '', SUPPORTED_EXTENSIONS
        )
        if files:
            self._add_files(files)

    def _add_files(self, new_files: List[str]):
        """添加文件（去重）"""
        existing = set(self._files)
        added = [f for f in new_files if f not in existing]

        if not added:
            return

        self._files.extend(added)
        self._refresh_list()
        self.files_changed.emit(self._files)

    def _remove_file(self, index: int):
        """移除指定文件"""
        if 0 <= index < len(self._files):
            self._files.pop(index)
            self._refresh_list()
            self.files_changed.emit(self._files)

    def _on_clear_all(self):
        self._files.clear()
        self._refresh_list()
        self.files_changed.emit(self._files)

    def _refresh_list(self):
        """刷新文件列表显示"""
        self.file_list.clear()

        if not self._files:
            self.file_list.setVisible(False)
            self.clear_btn.setEnabled(False)
            self.file_count_label.setText('')
            self.drop_area.setText('\U0001f4e5 拖拽文件到此处\n\n支持多文件同时拖拽')
            return

        self.file_list.setVisible(True)
        self.clear_btn.setEnabled(True)

        total_size = sum(os.path.getsize(f) for f in self._files if os.path.exists(f))
        self.file_count_label.setText(f'共 {len(self._files)} 个文件 ({format_size(total_size)})')
        self.drop_area.setText(f'\u2705 已选择 {len(self._files)} 个文件，继续添加可拖拽或选择')

        for idx, fpath in enumerate(self._files):
            item = QListWidgetItem(self.file_list)
            widget = FileItemWidget(idx, fpath)
            widget.remove_clicked.connect(self._remove_file)
            item.setSizeHint(widget.sizeHint())
            self.file_list.setItemWidget(item, widget)

    def get_files(self) -> List[str]:
        return self._files.copy()

    def has_files(self) -> bool:
        return len(self._files) > 0

    def clear(self):
        self._on_clear_all()
