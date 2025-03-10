
# 基团团费管理系统
一个简单高效的团费数据处理和管理工具，提供Excel式操作界面与插件扩展功能。

## 功能特点

- **直观的用户界面**：提供Excel风格的表格操作，无需使用命令行或配置其他环境
- **插件扩展系统**：支持通过插件管理器添加和运行外部功能模块
- **数据处理功能**：读取Excel表格，生成三联表Word文档，支持多种数据格式化操作
- **进度保存与恢复**：支持将工作进度保存为JSON文件，方便后续继续编辑
- **跨设备兼容**：支持在不同电脑或虚拟机上安装使用，保持目录结构即可正常运行

## 安装说明

### 方法一：使用安装向导

1. 运行`Setup`安装向导
2. 按照提示选择安装位置
3. 可选择是否创建桌面快捷方式和开始菜单项

### 方法二：直接解压

1. 下载[最新发布版本](https://github.com/daseinzc/HTW_GUI/releases/tag/v2.0.0)
2. 解压压缩包到任意位置
3. 不需要安装，可直接运行

## 使用指南

### 基本操作

1. 双击`HSGUI.exe`或桌面快捷方式启动程序
   - **注意**：如遇启动问题，可尝试右键点击并选择"以管理员身份运行"

2. 主界面功能：
   - 添加/删除行：选中行后使用底部按钮操作
   - 排序：按序号对表格进行排序
   - 保存进度（.json）：点击"保存进度"将当前工作保存为JSON文件
   - 导入进度(.json)/打开excel表：点击"导入进度"恢复之前保存的工作
   - 输出Excel：点击"输出为Excel"按钮生成`input/data.xlsx`文件

### 插件使用

1. 点击左侧导航栏中的"插件管理"图标
2. 首次使用需添加插件：
   - 点击"添加插件"按钮
   - 导航至`extension/TTGC/combine_docx.exe`并选择添加
3. 运行插件：
   - 找到已添加的`combine_docx.exe`插件
   - 点击对应的"运行"按钮
   - 系统将在`input/word_data`目录生成单个Word文档
   - 最终处理结果保存在`output/`文件夹

## 目录结构

```
├─ _internal/               # 内部程序文件
├─ HSGUI.exe                # 主程序可执行文件
├─ input/                   # 输入数据目录
│  └─ data.xlsx             # 导出的Excel数据文件
│  └─ word_data/            # 生成的Word文档
├─ output/                  # 处理结果输出目录
├─ extension/               # 插件目录
│  └─ TTGC/                 # 插件子目录
│     └─ combine_docx.exe   # 文档整合插件
└─ 测试.json                # 示例数据文件
```

## 常见问题

1. **无法启动程序**
   - 确保以管理员权限运行程序
   - 检查是否有防病毒软件阻止运行
   - 验证文件路径中是否包含特殊字符

2. **找不到文件**
   - 请确保关闭杀毒软件或将程序文件夹添加到安全名单
   - 检查目录结构是否完整
   - 确认文件访问权限是否正确

3. **插件无法运行**
   - 确保插件路径正确添加
   - 检查输入文件是否正确生成
   - 尝试重新添加插件

## 技术说明

* 本程序使用Python语言开发，采用PyQt5构建用户界面
* 支持Excel数据处理与Word文档生成
* 采用插件化设计，便于功能扩展

## 许可证

本项目采用MIT许可证开源，详情请参阅LICENSE文件。

```
MIT License
```

## 其他说明

* 此程序为轻量级应用，持续维护和改进中
* 仅供学习和参考使用
* 如需修改或扩展，可查看源代码进行自定义开发
* 测试环境：Windows 11 操作系统

---

© 2025 基团团费管理系统 | [GitHub项目主页](https://github.com/daseinzc/HTW_GUI)
