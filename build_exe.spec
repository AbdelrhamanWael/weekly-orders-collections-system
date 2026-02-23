# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec: نظام ربط الطلبات والتحصيل — Single .exe
# Run: pyinstaller build_exe.spec

block_cipher = None

a = Analysis(
    ['app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('templates', 'templates'),
        ('static', 'static'),
    ],
    hiddenimports=[
        'flask', 'werkzeug', 'jinja2', 'markupsafe', 'itsdangerous', 'click',
        'pandas', 'numpy', 'openpyxl', 'openpyxl.styles', 'openpyxl.utils',
        'openpyxl.utils.get_column_letter', 'xlsxwriter', 'xlrd',
        'plotly', 'python_dateutil', 'pytz', 'kaleido',
        'process_data', 'process_collections', 'generate_report', 'init_db',
    ],
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
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='OrdersCollectionsSystem',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
