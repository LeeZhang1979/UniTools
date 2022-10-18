# -*- mode: python ; coding: utf-8 -*-


block_cipher = None

py_files = [
    'Main.py',
    'ui\\__init__.py',
    'ui\\Ui_MainWindow.py',
    'ui\\Ui_MDMForm.py',
    'src\\__init__.py',
    'src\\MainWindow.py',
    'src\\MDMForm.py',
    'src\\__init__.py',
    'conf\\AppConfigure.py' 
]

data_files = [    
    ('conf\\AppConfigs.xml','conf\\'),
    ('conf\\App.ini','conf\\'),
    ('res\\imgs\\*','res\\imgs\\'),
    ('db\\*.db','db\\'),
    ('docs\\*','docs\\')
]
a = Analysis(py_files,
    pathex=[],
    binaries=[],
    datas=data_files,
    hiddenimports=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    name='UniTools',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='res\\imgs\\UniTools.ico'
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='UniTools'
)