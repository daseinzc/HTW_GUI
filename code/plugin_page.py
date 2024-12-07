import logging
import json
import os
import subprocess
import ctypes
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QHBoxLayout, QPushButton, QListWidgetItem, QWidget, QFileDialog, QTextEdit, QListWidget, \
    QLabel, QVBoxLayout, QProgressDialog


class PluginPage(QWidget):
    def __init__(self, main_page=None, stacked_widget=None):
        super().__init__()

        # 设置日志配置
        self.setup_logging()

        self.plugin_data_file = "plugins.json"  # 存储插件路径的文件
        self.main_page = main_page  # 用于返回主页面
        self.stacked_widget = stacked_widget  # 传递 stacked_widget
        self.loaded_plugins = []  # 内存中缓存插件数据
        self.init_plugin_page()

        self.log_file = "plugin_manager.log"  # 日志文件路径

    def setup_logging(self):
        """配置日志功能"""
        self.log_file = "plugin_manager.log"  # 日志文件路径
        logging.basicConfig(
            filename=self.log_file,
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )
        self.logger = logging.getLogger()

        # 在日志中记录一些初始化信息
        self.logger.info("PluginPage 初始化")


    def init_plugin_page(self):
        """初始化插件页内容"""

        self.setWindowTitle("插件管理")

        # 创建一个垂直布局
        self.layout = QVBoxLayout()


        # 创建返回按钮
        self.back_button = QPushButton("返回", self)
        self.back_button.setStyleSheet("background-color: #e67e22; color: white; padding: 10px; font-size: 16px; border-radius: 5px;")
        self.back_button.clicked.connect(self.go_back_to_main)

        # 创建一个水平布局来放置右上角的 "添加插件" 按钮
        top_layout = QHBoxLayout()

        # 创建页面标题
        self.header_label = QLabel("插件管理", self)
        self.header_label.setStyleSheet("font-size: 30px; font-weight: bold; padding: 10px; color: #2C3E50;")

        # 创建 "添加插件" 按钮
        self.add_plugin_button = QPushButton("添加插件", self)
        self.add_plugin_button.setStyleSheet("background-color: #3498DB; color: white; padding: 15px; font-size: 18px; border-radius: 8px;")
        self.add_plugin_button.clicked.connect(self.add_plugin)

        # 将按钮添加到布局
        top_layout.addWidget(self.header_label)
        top_layout.addWidget(self.add_plugin_button)

        # 创建显示插件的列表
        self.plugin_list = QListWidget(self)

        # 创建一个文本编辑框用于显示日志信息
        self.log_area = QTextEdit(self)
        self.log_area.setReadOnly(True)  # 设置为只读
        self.log_area.setStyleSheet(
            "font-size: 14px; background-color: #f1f1f1; border: 1px solid #ddd; padding: 10px;")

        # 将控件添加到主布局
        self.layout.addWidget(self.back_button)  # 添加返回按钮
        self.layout.addLayout(top_layout)  # 添加标题和按钮的布局
        self.layout.addWidget(self.plugin_list)  # 添加插件列表
        self.layout.addWidget(self.log_area)  # 添加日志区域

        # 设置插件页的主布局
        self.setLayout(self.layout)

        # 加载已保存的插件
        self.load_plugins()

    def go_back_to_main(self):
        """返回到主页面"""

        if self.stacked_widget is None or self.main_page is None:
            print("Error: stacked_widget or main_page is None")
            return

        # 直接切换页面
        self.stacked_widget.setCurrentWidget(self.main_page)

        def setup_logging(self):
            """设置日志配置"""
            logging.basicConfig(
                filename=self.log_file,  # 指定日志文件
                level=logging.DEBUG,  # 日志记录级别为 DEBUG，记录所有级别的日志
                format="%(asctime)s - %(levelname)s - %(message)s",  # 日志格式
            )

        def log_message(self, message):
            """记录日志并显示在 UI 中"""
            # 记录到日志文件
            logging.info(message)

            # 同时在 log_area 中显示
            self.log_area.append(message)

    def add_plugin(self):
        """打开文件对话框选择插件（.exe文件）"""
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "选择插件", "", "Executable Files (*.exe);;All Files (*)",
                                                   options=options)
        self.log_area.append(f"当前工作目录：{os.getcwd()}")  # 打印当前工作目录
        self.log_area.append(f"插件路径：{file_path}")

        if file_path:
            # 获取文件的名称（不包含路径）
            exe_name = os.path.basename(file_path)

            # 防止重复添加插件路径
            existing_plugins = self.load_plugins_data()
            if file_path in existing_plugins:
                self.log_area.append(f"插件已存在：{exe_name}")
                return

            # 添加插件路径到列表
            self.save_plugin(file_path)

            # 更新插件列表显示
            self.add_plugin_to_list(file_path)

    def add_plugin_to_list(self, file_path):
        """将插件添加到插件列表"""
        exe_name = os.path.basename(file_path)

        # 创建一个 QWidget 来包装列表项及其按钮
        widget = QWidget(self)  # 这里创建了一个 QWidget，作为列表项的容器

        # 创建一个水平布局来放置插件名称和按钮
        button_layout = QHBoxLayout()
        label = QLabel(exe_name)  # 显示插件名称
        label.setStyleSheet("font-size: 24px; font-weight: bold;")  # 设置字体大小和加粗样式
        button_layout.addWidget(label)

        # 创建 "运行" 按钮
        run_button = QPushButton("运行", self)
        run_button.setStyleSheet(
            "background-color: #81C784; color: white; padding: 8px; font-size: 12px; border-radius: 8px;")
        run_button.clicked.connect(lambda checked, exe=file_path: self.run_plugin(exe))  # 运行插件
        button_layout.addWidget(run_button)

        # 创建 "删除" 按钮
        del_button = QPushButton("删除", self)
        del_button.setStyleSheet(
            "background-color: #e57373; color: white; padding: 8px; font-size: 12px; border-radius: 8px;")
        del_button.clicked.connect(lambda checked, exe=file_path: self.delete_plugin(exe))  # 删除插件
        button_layout.addWidget(del_button)

        # 将按钮布局应用到 QWidget 中
        widget.setLayout(button_layout)

        # 调整高度和按钮宽度
        widget.setMinimumHeight(70)  # 增加容器的最小高度
        run_button.setMinimumWidth(70)  # 限制“运行”按钮的宽度
        del_button.setMinimumWidth(70)  # 限制“删除”按钮的宽度

        # 创建一个列表项
        list_item = QListWidgetItem(self.plugin_list)
        self.plugin_list.addItem(list_item)

        # 将创建的 QWidget 添加到该列表项的控件中
        self.plugin_list.setItemWidget(list_item, widget)

    def run_plugin(self, exe_path):
        """运行选中的插件（.exe文件）"""
        if not os.path.exists(exe_path):
            self.log_area.append(f"插件文件不存在：{exe_path}")
            return

        self.log_area.append(f"正在运行插件: {exe_path}")

        # 显示进度对话框
        progress_dialog = QProgressDialog("运行插件...", "取消", 0, 100, self)
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setValue(0)
        progress_dialog.show()

        try:
            # 获取插件所在目录
            plugin_dir = os.path.dirname(os.path.abspath(exe_path))

            # 使用 subprocess.run 运行插件，并设置工作目录为插件所在目录
            result = subprocess.run(
                exe_path,
                check=True,
                capture_output=True,
                text=True,
                cwd=plugin_dir  # 设置工作目录为插件所在目录
            )

            progress_dialog.setValue(100)  # 运行完毕，设置为100%
            self.log_area.append(f"运行成功：{exe_path}")
            self.log_area.append(f"输出：\n{result.stdout}")

        except subprocess.CalledProcessError as e:
            progress_dialog.setValue(100)
            self.log_area.append(f"运行出错：{exe_path}")
            self.log_area.append(f"错误：\n{e.stderr}")
        except FileNotFoundError as e:
            progress_dialog.setValue(100)
            self.log_area.append(f"未找到文件：{exe_path}")
            self.log_area.append(f"错误：{str(e)}")
        except PermissionError as e:
            progress_dialog.setValue(100)
            self.log_area.append(f"权限错误：{exe_path}")
            self.log_area.append(f"错误：{str(e)}")
        except Exception as e:
            progress_dialog.setValue(100)
            self.log_area.append(f"未知错误：{str(e)}")

    def delete_plugin(self, exe_path):
        """删除选中的插件"""
        self.log_area.append(f"删除插件: {exe_path}")

        # 删除插件路径
        plugins = self.load_plugins_data()  # 获取当前已加载的插件列表
        if exe_path in plugins:
            plugins.remove(exe_path)  # 移除插件路径

            # 将更新后的插件路径保存回文件
            with open(self.plugin_data_file, "w") as f:
                json.dump(plugins, f)

            # 从 UI 中移除对应的插件项
            for i in range(self.plugin_list.count()):
                item = self.plugin_list.item(i)
                widget_item = self.plugin_list.itemWidget(item)
                label = widget_item.findChild(QLabel)  # 获取标签
                if label and label.text() == os.path.basename(exe_path):
                    self.plugin_list.takeItem(i)  # 移除该项
                    break  # 停止遍历

            self.log_area.append(f"插件已从列表中删除: {exe_path}")
        else:
            self.log_area.append(f"插件路径未找到: {exe_path}")

    def save_plugin(self, plugin_path):
        """保存插件路径到文件"""
        plugins = self.load_plugins_data()  # 获取已保存的插件路径
        plugins.append(plugin_path)  # 添加新的插件路径

        # 将插件路径列表保存到 JSON 文件中
        with open(self.plugin_data_file, "w") as f:
            json.dump(plugins, f)

    def load_plugins(self):
        """加载插件路径并显示在插件列表中"""
        self.loaded_plugins = self.load_plugins_data()  # 只在初始化时加载
        for plugin_path in self.loaded_plugins:
            self.add_plugin_to_list(plugin_path)

    def load_plugins_data(self):
        """从文件中加载已保存的插件路径"""
        if self.loaded_plugins:  # 优先使用内存中的缓存数据
            return self.loaded_plugins
        if os.path.exists(self.plugin_data_file):
            with open(self.plugin_data_file, "r") as f:
                try:
                    file_content = f.read()
                    print(f"读取的文件内容: {file_content}")  # 日志记录读取的内容
                    self.loaded_plugins = json.loads(file_content)
                except json.JSONDecodeError as e:
                    print(f"JSON 解码错误: {e}")  # 打印错误信息
                    print(f"错误发生在: {file_content}")  # 打印出错时的文件内容
                    self.show_error_message("插件数据格式错误", "插件配置文件格式不正确，已恢复为默认状态。")
                    self.loaded_plugins = []  # 清空插件列表
                    # 重新保存一个空的插件列表
                    with open(self.plugin_data_file, "w") as f:
                        json.dump(self.loaded_plugins, f)
                    return self.loaded_plugins
            return self.loaded_plugins
        return []  # 如果没有文件，返回空列表

    def show_error_message(self, title, message):
        """显示错误提示框"""
        from PyQt5.QtWidgets import QMessageBox
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(message)
        msg.setWindowTitle(title)
        msg.exec_()



    def is_admin(self):
        """检查是否是管理员权限"""
        try:
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
