# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec 文件 —— WPS表格字段映射工具
在麒麟 Linux 上运行：pyinstaller wps_table_mapper.spec --clean --noconfirm
或直接执行 ./build.sh
"""

from PyInstaller.utils.hooks import collect_all

block_cipher = None

# 完整收集 PyQt5 的所有数据/二进制/隐藏导入（含 Qt 平台插件，防止运行时报错）
pyqt5_datas, pyqt5_binaries, pyqt5_hidden = collect_all('PyQt5')

a = Analysis(
    ['wps_table_mapper.py'],
    pathex=[],
    binaries=pyqt5_binaries,
    datas=pyqt5_datas,
    hiddenimports=pyqt5_hidden + [
        'xlrd',
        'openpyxl',
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.sip',
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
    [],
    exclude_binaries=True,
    name='WPS表格字段映射工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='WPS表格字段映射工具',
)
