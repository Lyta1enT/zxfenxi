"""处理日志面板"""
from datetime import datetime

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QTextEdit, QLabel, QHBoxLayout, QPushButton
)
from PyQt5.QtGui import QColor, QTextCharFormat


LOG_COLORS = {
    'INFO': QColor(0x33, 0x33, 0x33),
    'WARN': QColor(0xCC, 0x88, 0x00),
    'ERROR': QColor(0xCC, 0x33, 0x00),
    'SUCCESS': QColor(0x00, 0x88, 0x00),
}

LOG_PREFIX = {
    'INFO': ' \u2139\ufe0f',
    'WARN': ' \u26a0\ufe0f',
    'ERROR': ' \u274c',
    'SUCCESS': ' \u2705',
}


class LogPanel(QWidget):
    """日志显示面板"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._init_ui()

    def _init_ui(self):
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)

        header = QHBoxLayout()
        title = QLabel('\U0001f4dd 处理日志')
        title.setStyleSheet("font-size: 14px; font-weight: bold; color: #333;")
        header.addWidget(title)
        header.addStretch()

        clear_btn = QPushButton('清空')
        clear_btn.clicked.connect(self.clear_log)
        clear_btn.setStyleSheet("""
            QPushButton {
                padding: 2px 12px;
                font-size: 11px;
                background: #f0f0f0;
                border: 1px solid #ddd;
                border-radius: 3px;
            }
            QPushButton:hover { background: #e0e0e0; }
        """)
        header.addWidget(clear_btn)
        layout.addLayout(header)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                border: 1px solid #ddd;
                border-radius: 4px;
                font-family: monospace;
                font-size: 11px;
                background-color: #fafafa;
                padding: 6px;
            }
        """)
        layout.addWidget(self.log_area)

        self.setLayout(layout)

    def append_log(self, level: str, message: str):
        """添加日志"""
        timestamp = datetime.now().strftime('%H:%M:%S')
        color = LOG_COLORS.get(level, LOG_COLORS['INFO'])
        prefix = LOG_PREFIX.get(level, '')

        fmt = QTextCharFormat()
        fmt.setForeground(color)

        self.log_area.setCurrentCharFormat(fmt)
        self.log_area.append(f'[{timestamp}]{prefix} {message}')

        scrollbar = self.log_area.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def clear_log(self):
        self.log_area.clear()
