#!/usr/bin/env python3
"""征信报告OCR识别与生成工具 - 入口"""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from PyQt5.QtWidgets import QApplication
from PyQt5.QtGui import QFont

from app.main_window import MainWindow


def main():
    # 高DPI支持
    QApplication.setAttribute(0x10001)  # Qt.AA_EnableHighDpiScaling
    QApplication.setAttribute(0x10002)  # Qt.AA_UseHighDpiPixmaps

    app = QApplication(sys.argv)

    font = QFont('Microsoft YaHei', 9)
    app.setFont(font)

    window = MainWindow()
    window.show()

    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
