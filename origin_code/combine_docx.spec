# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py'],
    pathex=['.'],
    binaries=[],
    datas=[
	('docxcompose/templates/custom.xml','docxcompose/templates'),  # 包含 custom.xml
        ('clickme.xlsx', '.'),  # 包含 xlsx 文件
        ('fgx.png', '.'),  # 包含 png 文件
        ('input', 'input'),  # 包含 input 文件夹
        ('output', 'output')  # 包含 output 文件夹
],
    hiddenimports=[
	'docx',
        'docxcompose',
        'lxml',
        'combinemodule',
        'excel_catchmodule',
        'single_biaomodule',
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
    [],
    exclude_binaries=True,
    name='combine_docx',
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
    icon=['pic.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='combine_docx',
)
