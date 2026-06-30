# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 如果有默认配置文件可以加在这里
    ],
    hiddenimports=[
        # PyQt5 相关
        'PyQt5.sip',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        # OCR 引擎
        'easyocr',
        'easyocr.detection',
        'easyocr.recognition',
        # PDF 处理
        'fitz',
        'fitz.utils',
        # Excel 输出
        'openpyxl',
        'openpyxl.cell._writer',
        # 图片处理
        'cv2',
        'PIL',
        'PIL._imaging',
        # 文档处理
        'docx',
        # 工具
        'numpy',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'scipy',
        'pandas',
        'notebook',
        'IPython',
        'PyQt5.QtWebEngine',
        'PyQt5.QtWebEngineWidgets',
        'PyQt5.QtQml',
        'PyQt5.QtQuick',
        'PyQt5.QtSvg',
        'PyQt5.QtTest',
        'PyQt5.QtXml',
        'PyQt5.QtNetwork',
        'PyQt5.QtBluetooth',
        'PyQt5.QtMultimedia',
        'PyQt5.QtPrintSupport',
        'PyQt5.QtSql',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='征信报告OCR工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',        # 如果需要图标，放一个 icon.ico 在项目目录
)
