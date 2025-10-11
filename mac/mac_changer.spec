# -*- mode: python ; coding: utf-8 -*-
import sys
from PyInstaller.utils.hooks import collect_data_files

# 确保 PyInstaller 能够找到 netifaces 模块的数据文件（如果需要的话）
# 推荐使用 hook-netifaces.py 文件，如果未使用，可以尝试此方法
# netifaces_datas = collect_data_files('netifaces')
# 如果您已经使用了 hook-netifaces.py 文件，则无需手动添加 netifaces_datas

block_cipher = None

a = Analysis(
    ['mac.py'],
    pathex=['.'],
    binaries=[],
    # ----------------------------------------------------------------------
    # !!! 关键修正：添加 mac.ico 文件到打包数据中 !!!
    # ('mac.ico', '.') 表示将当前目录下的 mac.ico 打包到临时解压目录的根 ('.')
    # ----------------------------------------------------------------------
    datas=[('mac.ico', '.')], 
    hiddenimports=['netifaces', 'winreg', 'subprocess', 'platform'],
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data,
             cipher=block_cipher)

exe = EXE(pyz,
          a.scripts,
          a.binaries,
          a.zipfiles,
          a.datas,
          [],
          name='MAC_Changer',
          debug=False,
          bootloader_ignore_signals=False,
          strip=False,
          upx=True,
          upx_exclude=[],
          runtime_tmpdir=None,
          console=False,  # 设置为 False，确保无命令行窗口
          # ----------------------------------------------------------------------
          # !!! 关键修正：同时设置 EXE 文件本身的图标 (在 PyInstaller 外部设置) !!!
          # 这行设置了 Windows 资源管理器中显示 EXE 文件的图标
          # ----------------------------------------------------------------------
          icon='mac.ico' 
)