# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['C:/Users/Willi/Desktop/CS481/TARGEST.Final-1.0/main.py'],
    pathex=[],
    binaries=[],
    datas=[('C:/Users/Willi/Desktop/CS481/TARGEST.Final-1.0/itachiakatttt.png', '.'), ('C:/Users/Willi/Desktop/CS481/TARGEST.Final-1.0/TARGEST.png', '.'), ('C:/Users/Willi/Desktop/CS481/TARGEST.Final-1.0/TARGEST2.png', '.'), ('C:/Users/Willi/Desktop/CS481/TARGEST.Final-1.0/TARGEST3.ico', '.'), ('C:/Users/Willi/Desktop/CS481/TARGEST.Final-1.0/TARGEST3.png', '.')],
    hiddenimports=[],
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
    [],
    exclude_binaries=True,
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['C:\\Users\\Willi\\Desktop\\CS481\\TARGEST.Final-1.0\\TARGEST3.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
