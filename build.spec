# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec.

ビルド:
    pyinstaller build.spec --clean --noconfirm

成果物は dist/KindleAutoCapture/ 以下に出力される（onedir方式）。
"""

from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

hiddenimports = []
hiddenimports += collect_submodules("PyQt6")
hiddenimports += [
    "mss",
    "mss.windows",
    "PIL._tkinter_finder",
    "win32api",
    "win32con",
    "win32gui",
    "win32process",
    "imagehash",
    "keyboard",
    "pyautogui",
]

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=[],
    datas=[],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter",
        "PySide6",
        "PyQt5",
        "matplotlib",
        "scipy",
        "pandas",
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="KindleAutoCapture",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
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
    name="KindleAutoCapture",
)
