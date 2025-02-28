import logging
import json
import os
import subprocess
import ctypes
import sys
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QFont, QIcon, QColor, QPalette, QPixmap
from PyQt5.QtWidgets import (
    QHBoxLayout, QPushButton, QListWidgetItem, QWidget, QFileDialog, QTextEdit,
    QListWidget, QLabel, QVBoxLayout, QProgressDialog, QMessageBox, QFrame,
    QScrollArea, QSplitter, QApplication, QStyle
)


class PluginItem(QWidget):
    """自定义插件列表项组件"""

    def __init__(self, plugin_info, parent=None, run_callback=None, delete_callback=None):
        super().__init__(parent)
        self.plugin_info = plugin_info
        self.run_callback = run_callback
        self.delete_callback = delete_callback

        self.init_ui()

    def init_ui(self):
        # 创建水平布局
        layout = QHBoxLayout()

        # 创建左侧图标
        icon_label = QLabel()
        icon = QApplication.style().standardIcon(QStyle.SP_FileIcon)
        pixmap = icon.pixmap(QSize(48, 48))
        icon_label.setPixmap(pixmap)
        icon_label.setFixedSize(48, 48)
        layout.addWidget(icon_label)

        # 创建中间信息部分（垂直布局）
        info_layout = QVBoxLayout()

        # 插件名称（大字体，粗体）
        filename = self.plugin_info.get('filename', '')
        name_label = QLabel(os.path.splitext(filename)[0])
        name_label.setStyleSheet("font-size: 18px; font-weight: bold; color: #2c3e50;")
        info_layout.addWidget(name_label)

        # 插件路径（小字体，灰色）
        path_label = QLabel(self.plugin_info.get('path', ''))
        path_label.setStyleSheet("font-size: 12px; color: #7f8c8d;")
        path_label.setWordWrap(True)
        info_layout.addWidget(path_label)

        # 将信息布局添加到主布局
        layout.addLayout(info_layout, 1)  # 占据所有可用空间

        # 创建右侧按钮区域
        buttons_layout = QHBoxLayout()

        # 运行按钮 - 使用渐变背景
        self.run_btn = QPushButton("运行")
        self.run_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #4CAF50, stop:1 #388E3C);
                color: white;
                border: none;
                padding: 8px 16px;
                font-size: 14px;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #66BB6A, stop:1 #43A047);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #388E3C, stop:1 #2E7D32);
            }
        """)
        self.run_btn.clicked.connect(self.on_run)
        buttons_layout.addWidget(self.run_btn)

        # 删除按钮 - 使用渐变背景
        self.delete_btn = QPushButton("删除")
        self.delete_btn.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #F44336, stop:1 #D32F2F);
                color: white;
                border: none;
                padding: 8px 16px;
                font-size: 14px;
                border-radius: 4px;
                min-width: 80px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #EF5350, stop:1 #E53935);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #D32F2F, stop:1 #C62828);
            }
        """)
        self.delete_btn.clicked.connect(self.on_delete)
        buttons_layout.addWidget(self.delete_btn)

        layout.addLayout(buttons_layout)

        # 应用主布局到组件
        self.setLayout(layout)

        # 设置组件样式
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                border-radius: 6px;
                padding: 10px;
            }
        """)

        # 设置固定高度
        self.setMinimumHeight(90)

    def on_run(self):
        if self.run_callback:
            self.run_callback(self.plugin_info)

    def on_delete(self):
        if self.delete_callback:
            self.delete_callback(self.plugin_info)


class PluginPage(QWidget):
    def __init__(self, main_page=None, stacked_widget=None):
        super().__init__()

        # 设置日志配置
        self.setup_logging()

        # 获取应用程序路径
        self.app_path = self.get_application_path()

        # 配置插件数据文件路径
        self.plugin_data_file = os.path.join(self.app_path, "plugins.json")
        self.log_file = os.path.join(self.app_path, "plugin_manager.log")

        self.main_page = main_page  # 用于返回主页面
        self.stacked_widget = stacked_widget  # 传递 stacked_widget
        self.loaded_plugins = []  # 内存中缓存插件数据

        # 确保extension文件夹存在
        self.extension_path = self.ensure_extension_folder()

        self.init_plugin_page()

        # 日志记录初始化信息
        self.log_message(f"PluginPage初始化完成，插件数据文件：{self.plugin_data_file}")
        self.log_message(f"扩展文件夹路径：{self.extension_path}")

    def get_application_path(self):
        """获取应用程序的实际路径，兼容开发环境和PyInstaller打包后环境"""
        if getattr(sys, 'frozen', False):
            # 打包后的情况
            application_path = os.path.dirname(sys.executable)
        else:
            # 开发环境
            application_path = os.path.dirname(os.path.abspath(__file__))

        return application_path

    def ensure_extension_folder(self):
        """确保extension文件夹存在，如果不存在则创建"""
        extension_path = os.path.join(self.app_path, "extension")

        if not os.path.exists(extension_path):
            try:
                os.makedirs(extension_path)
                self.log_message(f"创建extension文件夹: {extension_path}")
            except Exception as e:
                self.log_message(f"创建extension文件夹失败: {str(e)}", level="error")

        return extension_path

    def setup_logging(self):
        """配置日志功能"""
        self.logger = logging.getLogger(__name__)

    def log_message(self, message, level="info"):
        """记录日志并显示在 UI 中"""
        if level == "info":
            logging.info(message)
        elif level == "error":
            logging.error(message)
        elif level == "warning":
            logging.warning(message)

        # 在UI中显示
        if hasattr(self, 'log_area'):
            self.log_area.append(message)

    def init_plugin_page(self):
        """初始化插件页内容"""
        # 设置整体样式
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f7fa;
                font-family: 'Segoe UI', 'Microsoft YaHei', sans-serif;
            }
            QLabel {
                color: #2c3e50;
            }
            QTextEdit {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                padding: 8px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 13px;
                color: #34495e;
            }
            QListWidget {
                background-color: transparent;
                border: none;
                outline: none;
            }
            QListWidget::item {
                background-color: transparent;
                padding: 5px;
            }
        """)

        # 创建主布局
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # 顶部区域 - 返回按钮和标题
        top_layout = QHBoxLayout()

        # 创建返回按钮 - 使用箭头图标
        self.back_button = QPushButton(" 返回")
        icon = QApplication.style().standardIcon(QStyle.SP_ArrowBack)
        self.back_button.setIcon(icon)
        self.back_button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                padding: 8px 16px;
                font-size: 14px;
                border-radius: 4px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #1a5276;
            }
        """)
        self.back_button.clicked.connect(self.go_back_to_main)
        top_layout.addWidget(self.back_button)

        # 创建页面标题
        self.header_label = QLabel("插件管理中心")
        self.header_label.setStyleSheet("""
            font-size: 28px; 
            font-weight: bold; 
            color: #2c3e50;
            padding: 0 15px;
        """)
        top_layout.addWidget(self.header_label, 1)  # 1表示伸缩因子，标题会占据更多空间

        # 创建 "添加插件" 按钮
        self.add_plugin_button = QPushButton(" 添加插件")
        icon = QApplication.style().standardIcon(QStyle.SP_FileDialogNewFolder)
        self.add_plugin_button.setIcon(icon)
        self.add_plugin_button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #3498db, stop:1 #2980b9);
                color: white;
                border: none;
                padding: 8px 20px;
                font-size: 14px;
                border-radius: 4px;
                min-width: 130px;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #5dade2, stop:1 #3498db);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 #2980b9, stop:1 #21618c);
            }
        """)
        self.add_plugin_button.clicked.connect(self.add_plugin)
        top_layout.addWidget(self.add_plugin_button)

        main_layout.addLayout(top_layout)

        # 添加分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setStyleSheet("background-color: #e0e0e0; max-height: 1px;")
        main_layout.addWidget(separator)

        # 插件列表区域 - 带有说明标签
        plugin_section = QVBoxLayout()

        # 增加说明标签
        info_label = QLabel("已安装的插件:")
        info_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #34495e; margin: 5px 0;")
        plugin_section.addWidget(info_label)

        # 创建插件列表 - 使用QListWidget
        self.plugin_list = QListWidget()
        self.plugin_list.setSpacing(10)  # 设置列表项之间的间距
        self.plugin_list.setStyleSheet("""
            QListWidget {
                background-color: transparent;
                border: none;
            }
            QListWidget::item {
                background-color: transparent;
                padding: 5px 0;
            }
        """)
        plugin_section.addWidget(self.plugin_list)

        # 空状态提示
        self.empty_label = QLabel("没有安装任何插件。点击按钮来添加第一个插件。")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet("""
            font-size: 14px; 
            color: #95a5a6; 
            background-color: #ecf0f1;
            padding: 30px;
            border-radius: 5px;
            margin: 20px 0;
        """)
        plugin_section.addWidget(self.empty_label)
        self.empty_label.hide()  # 默认隐藏

        main_layout.addLayout(plugin_section, 1)  # 插件列表占用更多垂直空间

        # 创建日志区域 - 带标题和清晰边框
        log_section = QVBoxLayout()

        # 日志标题
        log_label = QLabel("操作日志:")
        log_label.setStyleSheet("font-size: 16px; font-weight: bold; color: #34495e; margin: 5px 0;")
        log_section.addWidget(log_label)

        # 创建一个带有样式的文本编辑框用于显示日志信息
        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setStyleSheet("""
            QTextEdit {
                background-color: #ffffff;
                border: 1px solid #e0e0e0;
                border-radius: 4px;
                padding: 10px;
                font-family: 'Consolas', 'Courier New', monospace;
                font-size: 13px;
                color: #34495e;
                max-height: 150px;
            }
        """)
        log_section.addWidget(self.log_area)

        main_layout.addLayout(log_section)

        # 设置布局比例
        main_layout.setStretch(0, 0)  # 顶部区域不伸缩
        main_layout.setStretch(1, 0)  # 分隔线不伸缩
        main_layout.setStretch(2, 3)  # 插件列表占用更多空间
        main_layout.setStretch(3, 1)  # 日志区域占用较少空间

        # 加载已保存的插件
        self.load_plugins()

        # 根据是否有插件来显示空状态提示
        self.update_empty_state()

    def update_empty_state(self):
        """根据是否有插件来更新空状态显示"""
        if self.plugin_list.count() == 0:
            self.empty_label.show()
        else:
            self.empty_label.hide()

    def go_back_to_main(self):
        """返回到主页面"""
        if self.stacked_widget is None or self.main_page is None:
            self.log_message("Error: stacked_widget or main_page is None", level="error")
            return

        # 直接切换页面
        self.stacked_widget.setCurrentWidget(self.main_page)

    def add_plugin(self):
        """打开文件对话框选择插件（.exe文件）"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择插件", "", "可执行文件 (*.exe);;所有文件 (*)", options=options
        )

        if not file_path:
            return

        self.log_message(f"选择的插件：{file_path}")

        # 验证是否是有效的可执行文件
        if not os.path.isfile(file_path) or not file_path.lower().endswith('.exe'):
            self.log_message(f"选择的文件不是有效的可执行文件：{file_path}", level="error")
            QMessageBox.warning(self, "无效文件", "请选择有效的.exe可执行文件。")
            return

        # 获取文件的名称（不包含路径）
        exe_name = os.path.basename(file_path)

        # 检查是否已存在相同路径的插件
        existing_plugins = self.load_plugins_data()
        for plugin in existing_plugins:
            if plugin.get('path') == file_path:
                self.log_message(f"插件已存在：{exe_name}")
                QMessageBox.information(self, "插件已存在", f"插件 '{exe_name}' 已在列表中。")
                return

        # 添加插件信息
        plugin_info = {
            'name': os.path.splitext(exe_name)[0],
            'path': file_path,
            'filename': exe_name
        }

        # 保存到配置文件
        self.save_plugin(plugin_info)

        # 更新插件列表显示
        self.add_plugin_to_list(plugin_info)

        # 更新空状态
        self.update_empty_state()

        QMessageBox.information(self, "添加成功", f"插件 '{exe_name}' 已成功添加。")

    def add_plugin_to_list(self, plugin_info):
        """将插件添加到插件列表"""
        # 创建自定义列表项组件
        plugin_item = PluginItem(
            plugin_info,
            self,
            run_callback=self.run_plugin,
            delete_callback=self.delete_plugin
        )

        # 创建列表项
        item = QListWidgetItem(self.plugin_list)

        # 设置项目大小
        item.setSizeHint(plugin_item.sizeHint())

        # 将自定义组件添加到列表项
        self.plugin_list.addItem(item)
        self.plugin_list.setItemWidget(item, plugin_item)

    def run_plugin(self, plugin_info):
        """运行选中的插件（.exe文件）"""
        plugin_path = plugin_info.get('path', '')
        if not os.path.exists(plugin_path):
            self.log_message(f"插件文件不存在：{plugin_path}", level="error")
            QMessageBox.warning(self, "文件不存在", f"插件文件 '{plugin_path}' 不存在或无法访问。")
            return

        self.log_message(f"正在运行插件: {plugin_path}")

        # 显示进度对话框
        progress_dialog = QProgressDialog("正在启动插件...", "取消", 0, 100, self)
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setValue(0)
        progress_dialog.setStyleSheet("""
            QProgressDialog {
                background-color: #ffffff;
                border-radius: 5px;
            }
            QProgressBar {
                border: 1px solid #e0e0e0;
                border-radius: 3px;
                background-color: #f5f5f5;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                width: 10px;
                margin: 0.5px;
            }
        """)
        progress_dialog.show()

        try:
            # 获取插件所在目录
            plugin_dir = os.path.dirname(os.path.abspath(plugin_path))
            self.log_message(f"插件工作目录: {plugin_dir}")

            # 使用 subprocess.Popen 运行插件，并设置工作目录为插件所在目录
            process = subprocess.Popen(
                plugin_path,
                cwd=plugin_dir,  # 设置工作目录为插件所在目录
                shell=True,  # 使用shell运行，解决某些权限问题
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )

            # 设置进度对话框为50%表示启动成功
            progress_dialog.setValue(50)

            # 等待进程完成，最多等待1秒以避免界面卡死
            try:
                stdout, stderr = process.communicate(timeout=1)
                progress_dialog.setValue(100)
                self.log_message(f"插件启动成功：{os.path.basename(plugin_path)}")
                if stdout:
                    self.log_message(f"输出：\n{stdout}")
            except subprocess.TimeoutExpired:
                # 如果超时，说明插件可能是长时间运行的程序
                progress_dialog.setValue(100)
                self.log_message(f"插件正在运行中：{os.path.basename(plugin_path)}")
                # 不终止进程，让它继续在后台运行

        except Exception as e:
            progress_dialog.setValue(100)
            self.log_message(f"运行插件时出错: {str(e)}", level="error")
            QMessageBox.critical(self, "运行错误", f"运行插件时发生错误：{str(e)}")

    def delete_plugin(self, plugin_info):
        """从列表中删除插件（不删除实际文件）"""
        plugin_path = plugin_info.get('path', '')
        plugin_name = plugin_info.get('name', os.path.basename(plugin_path))

        self.log_message(f"准备删除插件: {plugin_name}")

        # 确认是否删除
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定要从列表中删除插件 '{plugin_name}' 吗？\n\n注意：这不会删除实际的插件文件。",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )

        if reply != QMessageBox.Yes:
            return

        # 删除插件记录
        plugins = self.load_plugins_data()
        updated_plugins = [p for p in plugins if p.get('path') != plugin_path]

        # 将更新后的插件列表保存回文件
        try:
            with open(self.plugin_data_file, "w", encoding='utf-8') as f:
                json.dump(updated_plugins, f, ensure_ascii=False, indent=4)

            # 更新内存缓存
            self.loaded_plugins = updated_plugins

            # 从UI中删除所有项并重新加载
            self.plugin_list.clear()
            for plugin in updated_plugins:
                self.add_plugin_to_list(plugin)

            # 更新空状态显示
            self.update_empty_state()

            self.log_message(f"插件已从列表中删除: {plugin_name}")

        except Exception as e:
            self.log_message(f"删除插件时出错: {str(e)}", level="error")
            QMessageBox.critical(self, "错误", f"删除插件时发生错误：{str(e)}")

    def save_plugin(self, plugin_info):
        """保存插件信息到文件"""
        plugins = self.load_plugins_data()  # 获取已保存的插件列表

        # 检查是否已存在
        for plugin in plugins:
            if plugin.get('path') == plugin_info.get('path'):
                return  # 已存在则不添加

        plugins.append(plugin_info)  # 添加新的插件信息
        self.loaded_plugins = plugins  # 更新内存缓存

        # 将插件列表保存到 JSON 文件中
        try:
            with open(self.plugin_data_file, "w", encoding='utf-8') as f:
                json.dump(plugins, f, ensure_ascii=False, indent=4)
            self.log_message(f"插件信息已保存到：{self.plugin_data_file}")
        except Exception as e:
            self.log_message(f"保存插件信息失败：{str(e)}", level="error")

    def load_plugins(self):
        """加载插件并显示在插件列表中"""
        self.loaded_plugins = self.load_plugins_data()

        # 清空当前列表
        self.plugin_list.clear()

        # 添加到UI
        for plugin_info in self.loaded_plugins:
            self.add_plugin_to_list(plugin_info)

        # 更新空状态显示
        self.update_empty_state()

        self.log_message(f"已加载 {len(self.loaded_plugins)} 个插件")

    def load_plugins_data(self):
        """从文件中加载已保存的插件信息"""
        if self.loaded_plugins:  # 优先使用内存中的缓存数据
            return self.loaded_plugins

        if os.path.exists(self.plugin_data_file):
            try:
                with open(self.plugin_data_file, "r", encoding='utf-8') as f:
                    try:
                        file_content = f.read()
                        self.log_message(f"读取插件数据文件，大小：{len(file_content)} 字节")

                        if not file_content.strip():
                            self.log_message("插件数据文件为空，使用空列表")
                            return []

                        plugins = json.loads(file_content)

                        # 兼容旧格式（如果是字符串列表则转为对象列表）
                        if plugins and isinstance(plugins[0], str):
                            self.log_message("检测到旧格式的插件数据，正在转换...")
                            new_plugins = []
                            for path in plugins:
                                new_plugins.append({
                                    'name': os.path.splitext(os.path.basename(path))[0],
                                    'path': path,
                                    'filename': os.path.basename(path)
                                })
                            plugins = new_plugins

                            # 保存转换后的格式
                            with open(self.plugin_data_file, "w", encoding='utf-8') as fw:
                                json.dump(plugins, fw, ensure_ascii=False, indent=4)

                        self.loaded_plugins = plugins
                        self.log_message(f"成功加载 {len(plugins)} 个插件信息")

                    except json.JSONDecodeError as e:
                        self.log_message(f"JSON 解码错误: {str(e)}", level="error")
                        self.show_error_message("插件数据格式错误", "插件配置文件格式不正确，已恢复为默认状态。")
                        self.loaded_plugins = []  # 清空插件列表
                        # 重新保存一个空的插件列表
                        with open(self.plugin_data_file, "w", encoding='utf-8') as fw:
                            json.dump(self.loaded_plugins, fw, ensure_ascii=False)
            except Exception as e:
                self.log_message(f"加载插件数据时出错: {str(e)}", level="error")
                self.loaded_plugins = []

            return self.loaded_plugins
        else:
            self.log_message(f"插件数据文件不存在，将创建: {self.plugin_data_file}")
            # 创建空文件
            try:
                with open(self.plugin_data_file, "w", encoding='utf-8') as f:
                    json.dump([], f)
            except Exception as e:
                self.log_message(f"创建插件数据文件失败: {str(e)}", level="error")

            return []  # 返回空列表

    def show_error_message(self, title, message):
        """显示错误提示框"""
        QMessageBox.critical(self, title, message)