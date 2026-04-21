# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['EVE_DPS.py'],
    pathex=[],
    binaries=[],
    datas=[('EVE_DPS_icon.ico', '.')],
    hiddenimports=[
        'datetime',
        'glob',
        'math',
        'struct',
        'tempfile',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.ttk',
        'wave',
        'winsound',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='EVE_DPS',
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
    icon=['EVE_DPS_icon.ico'],
)
