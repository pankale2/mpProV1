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
        'xlsxwriter.utility',
        'pandas',
        'pandas.io.excel',
        'pandas.io.formats.excel'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # Exclude unused pandas modules
        'pandas.tests',
        'pandas.io.clipboard',
        'pandas.plotting',
        'pandas.io.sql',
        'pandas.io.sas',
        'pandas.io.spss',
        'pandas.io.stata',
        'pandas.io.feather',
        'pandas.io.parquet',
        'pandas.io.orc',
        'pandas.io.gbq',
        
        # Exclude unused scientific libraries
        'scipy',
        'numpy.tests',
        'matplotlib',
        'plotly',
        'bokeh',
        'seaborn',
        
        # Exclude development tools
        'pytest',
        'IPython',
        'jupyter',
        'notebook',
        
        # Exclude unused standard library modules
        'tkinter',
        'turtle',
        'pydoc',
        'doctest',
        'xmlrpc',
        'http.server',
        'socketserver',
        'wsgiref',
        
        # Exclude unused networking
        'urllib3.contrib',
        'requests_oauthlib',
        'cryptography.hazmat.backends.commoncrypto',
        'cryptography.hazmat.backends.openssl'
    ],
    noarchive=False,
    optimize=2,  # Maximum optimization
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
    strip=True,  # Strip debugging symbols
    upx=True,    # Enable UPX compression
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
