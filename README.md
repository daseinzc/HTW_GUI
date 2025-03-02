# HTW_GUI
用于进行数据的简单处理和三联表生成。 含**整合功能插件** 插件和 **UI界面**。

**https://github.com/daseinzc/HTW_GUI/releases/tag/v2.0.0**
---

## 主要功能

1. **插件程序**：可以执行读取excel表，生成单个word三联表并进行整合操作。  
2. **UI 界面**：提供简单的图形界面，无需使用命令行或配置其他环境，拥有类excel功能。  
3. 支持在多台电脑或虚拟机上安装，只要保证目录结构不变，即可正常读取文件。
4. 左上角“文件”中“保存进度”会在本地保存一个后缀为`.json`文件，“导入进度”即选择`.json`文件导入。

## 使用说明

- **安装**：  
  1. 使用 `Setup` 安装向导进行安装或直接解压压缩包。  
  2. 若需要桌面快捷方式等，可在安装过程中勾选选项。  

- **运行**：  
  1. 双击 `HSGUI.exe`或桌面快捷方式启动主页面。**如果存在无法打开或长时间不响应的问题，可以右键以管理员权限打开**。  
  2. 在界面下方可以点击“添加行”、“删除行”、“排序”和“输出excel文档”按钮，输出的excel文档位于同一目录下`input/`，命名为`data.xlsx`。
  3. 输出`data.xlsx`后，点击左侧导航栏中的“插件管理”，找到`combine_docx.exe`插件，点击右侧`运行`（第一次运行需要手动添加插件。点击“添加插件”，在同一目录下找到`extension/TTGC/combine_docx.exe`并添加。）。运行后，在同一目录下`input/word_data`生成单个word文档，同时最终处理结果保存在 `output/` 文件夹下。

## 目录结构

```
├─ _internal/               
├─ HSGUI.exe                    # UI相关可执行文件
├─ input/                 # 需要读取的输入文件
├─ output/                # 处理后的输出结果
├─ 测试.json              # 测试文件
└─ extension              # 插件相关可执行文件等文件
```

## 常见问题

1. **找不到文件**  
   - 请确保关闭杀毒软件或将文件夹加入安全名单。

## 许可证

```
MIT License
```

## 其他说明

- 此为小型程序，由python编写完成，非专业，仍在维护和改进，仅供参考和学习。  
- 若修改或扩展此程序，可以查看源码进行更改。  

---
运行测试环境为win11系统
