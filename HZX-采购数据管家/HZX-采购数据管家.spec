# -*- mode: python ; coding: utf-8 -*-
"""
HZX-采购数据管家 V1.0.0
PyInstaller 打包配置文件
"""
import os, sys, platform

block_cipher = None

app_dir = os.path.dirname(os.path.abspath(__file__))

a = Analysis(
    ['app.py'],
    pathex=[app_dir],
    binaries=[],
    datas=[
        ('config.py', '.'),
        ('main_processing.py', '.'),
        ('pages\\__init__.py', 'pages'),
    ],
    hiddenimports=[
        # win32com
        'win32com',
        'win32com.client',
        'win32com.gen_py',
        # openpyxl
        'openpyxl',
        'openpyxl.cell',
        'openpyxl.styles',
        'openpyxl.worksheet',
        'openpyxl.reader',
        'openpyxl.writer',
        'openpyxl.writer.excel',
        'openpyxl.reader.excel',
        'openpyxl.formatting',
        # qfluentwidgets
        'qfluentwidgets',
        'qfluentwidgets.common',
        'qfluentwidgets.common.icon',
        'qfluentwidgets.components',
        'qfluentwidgets.window',
        'qfluentwidgets._rc',
        # PyQt5
        'PyQt5',
        'PyQt5.QtCore',
        'PyQt5.QtGui',
        'PyQt5.QtWidgets',
        'PyQt5.QtSvg',
        # darkdetect
        'darkdetect',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'tkinter',
        'matplotlib',
        'PIL',
        'scipy',
        'numpy',
        'cv2',
        'pandas',
        'requests',
        'cffi',
        'cryptography',
        'sqlite3',
        'ssl',
    ],
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
    name='HZX-采购数据管家',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=os.path.join(app_dir, 'app_icon.ico'),
)
