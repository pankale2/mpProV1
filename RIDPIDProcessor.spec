# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['run.py'],
    pathex=[],
    binaries=[],
    datas=[('templates', 'templates'), ('static', 'static')],
    hiddenimports=[
        'openpyxl',
        'openpyxl.styles',
        'openpyxl.formatting',
        'openpyxl.formatting.rule',
        'xlsxwriter',
        'pandas',
        'pandas.io.excel'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],  # Simplified - no exclusions
    noarchive=False,
    optimize=0,  # No optimization for simpler debugging
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='RIDPIDProcessor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,  # Keep symbols for better error messages
    upx=False,    # Disable UPX compression
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['static\\favicon.ico'],
)