# -*- mode: python ; coding: utf-8 -*-
import certifi

a = Analysis(
    ['trackman_gui_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        (certifi.where(), 'certifi'),  # SSL CA bundle for requests/aiohttp in frozen build
    ],
    hiddenimports=['customtkinter', 'openpyxl', 'pandas', 'PIL', 'PIL._tkinter_finder', 'certifi',
                   'cryptography', 'cryptography.hazmat.primitives.ciphers.aead'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'scipy', 'IPython', 'notebook', 'pytest',
        'lib2to3', 'pydoc', 'doctest',
    ],
    noarchive=False,
    optimize=1,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],          # binaries moved to COLLECT – no temp extraction on launch
    name='TrackmanConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,   # disabled: saves decompression time at startup
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='TrackmanConverter',
)
