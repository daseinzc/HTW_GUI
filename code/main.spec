# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['main.py','plugin_page.py'],
    pathex=[],
    binaries=[],
    datas=[
	 ('input/*', 'input'),
	 ('output/*', 'output'),
 	 ('extension/*', 'extension'),
         ('测试.json', '.'),
	 ('plugins.json', '.'),
	 ('plugin_manager.log', '.'),
	 ('program.log', '.'),
],
    hiddenimports=[],
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
    name='HSGUI',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='pic.ico',  # 添加图标文件
    onefile=True  # 启用 onefile 模式
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='main',
)
