# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file — WPS表格字段映射工具
在麒麟 Linux 上运行：pyinstaller wps_table_mapper.spec --clean --noconfirm
或直接执行 ./build.sh
"""

import sys
import os

block_cipher = None

# 尝试收集 PyQt5 所有组件（若失败则回退到手动 hiddenimports）
try:
    from PyInstaller.utils.hooks import collect_all
    pyqt5_datas, pyqt5_binaries, pyqt5_hidden = collect_all('PyQt5')
except Exception:
    pyqt5_datas, pyqt5_binaries, pyqt5_hidden = [], [], [
        'PyQt5', 'PyQt5.QtCore', 'PyQt5.QtGui', 'PyQt5.QtWidgets', 'PyQt5.sip',
    ]

a = Analysis(
    ['wps_table_mapper.py'],
    pathex=[],
    binaries=pyqt5_binaries,
    datas=pyqt5_datas,
    hiddenimports=pyqt5_hidden + [
        'xlrd',
        'openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'PyQt6',
        'PySide2',
        'PySide6',
        'cv2',
        'torch',
        'tensorflow',
    ],
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
    name='WPS表格字段映射工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    icon=None,
)