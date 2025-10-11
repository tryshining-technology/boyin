# -*- mode: python ; coding: utf-8 -*-

# 这是一个修复了执行顺序问题的 .spec 文件

import os
from win32com.client import gencache
import win32com

# --- Block-based vars ---
# 将所有变量定义和预处理操作都放在 Analysis 块之前
# 这样 PyInstaller 在解析 Analysis 时就能找到它们

# --- 核心修复：在这里预先生成 COM 缓存 ---
print("--- Forcing generation of win32com cache for SAPI.SpVoice ---")
gencache.EnsureDispatch('SAPI.SpVoice')
print("--- COM Cache generated successfully. ---")

# --- 找到生成的缓存文件夹 (gen_py) 的路径 ---
gen_py_path = os.path.join(os.path.dirname(win32com.__file__), 'gen_py')
print(f"--- Found gen_py cache path: {gen_py_path} ---")

# --- 定义要打包的数据 ---
datas_to_bundle = [
    ('icon.ico', '.'),
    (gen_py_path, 'win32com/gen_py')
]

# --- PyInstaller 分析块 ---
a = Analysis(
    ['boyin.py'],
    pathex=[],
    binaries=[],
    datas=datas_to_bundle, # 使用上面定义好的变量
    hiddenimports=['win32com.gen_py', 'win32timezone'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='boyin',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False, # False = --windowed
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='boyin',
)
