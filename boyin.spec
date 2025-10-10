# -*- mode: python ; coding: utf-8 -*-

# 这个 .spec 文件是解决 pywin32 打包问题的最终方案。
# 它会在打包前，用 Python 代码强制生成 COM 缓存，并将其作为数据文件包含进来。

from win32com.client import gencache
import os
import win32com

# --- 核心修复：在分析代码前，强制生成 SAPI.SpVoice 的缓存 ---
print("--- Forcing generation of win32com cache for SAPI.SpVoice ---")
gencache.EnsureDispatch('SAPI.SpVoice')
print("--- COM Cache generated successfully. ---")

# --- 找到生成的缓存文件夹 (gen_py) 的路径 ---
gen_py_path = os.path.join(os.path.dirname(win32com.__file__), 'gen_py')
print(f"--- Found gen_py cache path: {gen_py_path} ---")


# --- PyInstaller 分析块 ---
a = Analysis(
    ['boyin.py'],
    pathex=[],
    binaries=[],
    # --- 关键：将图标和生成的 COM 缓存文件夹作为数据文件打包 ---
    datas=[
        ('icon.ico', '.'),
        (gen_py_path, 'win32com/gen_py')
    ],
    # --- 双重保险：保留 hidden imports ---
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
    console=False, # 设置为 False 来创建 --windowed 应用
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
