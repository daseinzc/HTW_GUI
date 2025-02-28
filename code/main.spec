# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from PyInstaller.utils.hooks import collect_data_files, collect_all

# 使用SPECPATH变量获取spec文件所在目录作为基准点
# 这个变量由PyInstaller自动提供
base_path = SPECPATH
print(f"Spec文件目录: {base_path}")

# 所有路径都相对于spec文件所在目录
image_dir = os.path.join('image')
icon_path = os.path.join(image_dir, 'pic.ico')

print(f"图标相对路径: {icon_path}")
print(f"解析后的图标路径: {os.path.abspath(os.path.join(base_path, icon_path))}")
print(f"图标文件存在: {os.path.exists(os.path.join(base_path, icon_path))}")

# 创建一个包含所有图像文件的数据列表
image_files = []
full_image_dir = os.path.join(base_path, image_dir)
if os.path.exists(full_image_dir):
    print(f"图像目录存在: {full_image_dir}")
    for file in os.listdir(full_image_dir):
        if file.endswith(('.svg', '.ico', '.png')):
            # 源文件是相对于spec文件的完整路径
            source_file = os.path.join(image_dir, file)
            # 目标保持相对路径结构
            target_path = image_dir
            image_files.append((os.path.join(base_path, source_file), target_path))
            print(f"添加图像文件: {source_file} -> {target_path}")
else:
    print(f"警告：图像目录不存在：{full_image_dir}")

# 收集qt_material的资源文件
qt_material_data = collect_data_files('qt_material')

# 收集numpy和pandas的数据文件和依赖
np_datas, np_binaries, np_imports = collect_all('numpy')
pd_datas, pd_binaries, pd_imports = collect_all('pandas')

# 组合所有数据文件和二进制文件
all_datas = image_files + qt_material_data + np_datas + pd_datas
all_binaries = np_binaries + pd_binaries

# 组合所有隐藏导入
all_hidden_imports = ['qt_material', 'numpy', 'pandas']

a = Analysis(
    ['main.py'],  # 使用相对路径
    pathex=[base_path],  # 使用spec文件目录作为基准路径
    binaries=all_binaries,
    datas=all_datas,
    hiddenimports=all_hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='团费整理系统',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,  
    icon=os.path.join(base_path, icon_path),  # 使用相对于spec文件的路径
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='团费整理系统',
)