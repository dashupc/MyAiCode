# hook-netifaces.py

from PyInstaller.utils.hooks import collect_all

# collect_all 会智能地在安装路径中查找 netifaces 的所有依赖项
# 包括隐藏的二进制文件 (DLLs/PYDs)，并确保它们被包含。
datas, binaries, hiddenimports = collect_all('netifaces')