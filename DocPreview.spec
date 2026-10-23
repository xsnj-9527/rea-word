# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller build spec for DocPreview.

Produces a single windowed executable with the icon and the icon resource
bundled. Built with `pyinstaller DocPreview.spec` from the repo root.
"""

import os

block_cipher = None

SRC = os.path.join(os.getcwd(), "src")
ICON = os.path.join(SRC, "icon.ico")

a = Analysis(
    [os.path.join(SRC, "app.py")],
    pathex=[SRC],
    binaries=[],
    datas=[(os.path.join(SRC, "icon.ico"), "."),
           (os.path.join(SRC, "icon.png"), ".")],
    hiddenimports=[
        # pywebview picks its backend at runtime; these must be bundled.
        "webview.platforms.winforms",
        "webview.platforms.edgechromium",
        "clr_loader",
        "pythonnet",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Only safe exclusions: these are large and genuinely unused by the
        # viewer. Do NOT add stdlib packages here — pkg_resources and pywebview
        # import email/http/etc. at startup, and excluding them aborts the
        # process with ModuleNotFoundError before any window appears.
        "tkinter", "unittest", "pydoc", "doctest",
        "numpy", "pandas", "matplotlib", "scipy",
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
    name="DocPreview",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # windowed: no console flashes on launch
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICON,
)
