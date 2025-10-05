# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['multi_format_extractor.py'],  # 你的Python脚本文件名
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['comtypes.client', 'xlrd', 'openpyxl', 'pdfminer', 'bs4'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='多格式文件提取工具',  # 生成的EXE文件名
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',  # 图标文件路径（请将icon.ico放在与spec文件同目录）
)
    