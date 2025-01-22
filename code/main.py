import sys
import os
import json
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem, QVBoxLayout,
    QPushButton, QWidget, QHBoxLayout, QAbstractItemView, QFileDialog, QMessageBox,
    QLineEdit, QStackedWidget, QLabel, QHeaderView, QAction
)
from PyQt5.QtCore import Qt, QSize
from PyQt5.QtGui import QIcon, QFont
from qt_material import apply_stylesheet
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from plugin_page import PluginPage  # 导入 PluginPage 类

# 自定义支持数字排序
class NumericTableWidgetItem(QTableWidgetItem):
    def __lt__(self, other):
        try:
            self_val = float(self.text())
        except ValueError:
            self_val = 0.0
        try:
            other_val = float(other.text())
        except ValueError:
            other_val = 0.0
        return self_val < other_val

class ModernTableApp(QMainWindow):
    def __init__(self):
        super().__init__()

        # 1) 全局字体
        font = QFont("Roboto", 12)
        font.setStyleHint(QFont.SansSerif)
        font.setWeight(QFont.Normal)
        app.setFont(font)

        self.setWindowTitle("基团团费整理系统")
        self.setGeometry(100, 100, 1400, 800)
        self.setWindowFlags(Qt.Window | Qt.WindowTitleHint | Qt.WindowMinMaxButtonsHint | Qt.WindowCloseButtonHint)

        # 设置图标（可选）
        base_path = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_path, "../pic.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        # 2) 中央布局：左侧导航(窄) + 右侧StackedWidget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)
        main_layout.setContentsMargins(0,0,0,0)

        # 左侧导航栏(更窄)
        self.nav_widget = QWidget()
        self.nav_widget.setFixedWidth(60)  # 仅容纳图标
        self.nav_widget.setStyleSheet("""
            QWidget {
                background-color: #f5f5f5; 
                border-right: 1px solid #e0e0e0;
            }
        """)
        self.nav_layout = QVBoxLayout(self.nav_widget)
        self.nav_layout.setContentsMargins(0,10,0,10)
        self.nav_layout.setSpacing(10)

        # 右侧 StackedWidget
        self.stacked_widget = QStackedWidget()
        self.main_page = QWidget(self)
        self.option_page = QWidget(self)
        self.plugin_page = PluginPage(main_page=self.main_page, stacked_widget=self.stacked_widget)

        self.stacked_widget.addWidget(self.main_page)
        self.stacked_widget.addWidget(self.option_page)
        self.stacked_widget.addWidget(self.plugin_page)
        self.stacked_widget.setCurrentWidget(self.main_page)

        main_layout.addWidget(self.nav_widget)
        main_layout.addWidget(self.stacked_widget)

        # ========== 导航栏按钮(纯图标) ==========
        # 3) 加载图标(黑色线条),你可用 home_black.svg/settings_black.svg/plugin_black.svg
        icon_home = QIcon(os.path.join(base_path, "home.svg"))
        icon_settings = QIcon(os.path.join(base_path, "settings.svg"))
        icon_plugin = QIcon(os.path.join(base_path, "plugin.svg"))

        self.btn_go_main = QPushButton()
        self.btn_go_main.setIcon(icon_home)
        self.btn_go_main.setIconSize(QSize(24,24))
        self.btn_go_main.setToolTip("首页")

        self.btn_go_option = QPushButton()
        self.btn_go_option.setIcon(icon_settings)
        self.btn_go_option.setIconSize(QSize(24,24))
        self.btn_go_option.setToolTip("选项页")

        self.btn_go_plugin = QPushButton()
        self.btn_go_plugin.setIcon(icon_plugin)
        self.btn_go_plugin.setIconSize(QSize(24,24))
        self.btn_go_plugin.setToolTip("插件管理")

        # 4) QSS: 透明默认背景 + 悬浮/按下显示灰色, 让图标有焦点反馈
        #    并保持按钮区域小, padding: 8px, 这样鼠标点击范围大些
        btn_style = """
            QPushButton {
                background-color: transparent; 
                border: none;
                padding: 8px; 
                margin: 0px;
            }
            QPushButton:hover {
                background-color: #e0e0e0;
            }
            QPushButton:pressed {
                background-color: #cccccc;
            }
            QPushButton:focus {
                outline: none;
                border: none;
            }
        """
        for btn in [self.btn_go_main, self.btn_go_option, self.btn_go_plugin]:
            btn.setStyleSheet(btn_style)

        # 5) 点击事件切换页面
        self.btn_go_main.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.main_page))
        self.btn_go_option.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.option_page))
        self.btn_go_plugin.clicked.connect(lambda: self.stacked_widget.setCurrentWidget(self.plugin_page))

        # 加入左侧布局
        self.nav_layout.addWidget(self.btn_go_main)
        self.nav_layout.addWidget(self.btn_go_option)
        self.nav_layout.addWidget(self.btn_go_plugin)
        self.nav_layout.addStretch(1)

        # 初始化三个页面 + 菜单
        self.init_main_page()
        self.init_option_page()
        self.init_menu()

    def init_menu(self):
        menubar = self.menuBar()
        menubar.setStyleSheet("font-size: 16px; background-color: #ffffff; color: #333333;")
        file_menu = menubar.addMenu("文件")

        save_action = QAction("保存进度", self)
        save_action.triggered.connect(self.save_progress)
        file_menu.addAction(save_action)

        load_action = QAction("导入进度", self)
        load_action.triggered.connect(self.load_progress)
        file_menu.addAction(load_action)

    def init_main_page(self):
        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["序号", "学院", "财务金额", "是否补交"])
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setStyleSheet("font-size: 24px; selection-background-color: #a5d6a7;")
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)

        # 填充学院数据
        self.schools = [
            "社会学院", "网络空间安全学院", "光学与电子信息学院", "人文学院", "未来技术学院", "机械科学与工程学院",
            "土木与水利工程学院", "马克思主义学院", "口腔医学院", "能源与动力工程学院", "外国语学院", "护理学院",
            "化学与化工学院", "新闻与信息传播学院", "生命科学与技术学院", "电子信息与通信学院", "法学院", "哲学学院",
            "航空航天学院", "人工智能与自动化学院", "医药卫生管理学院", "物理学院", "管理学院", "环境科学与工程学院",
            "公共管理学院", "第一临床学院", "计算机科学与技术学院", "艺术学院", "法医学系", "数学与统计学院",
            "第二临床学院", "电气与电子工程学院", "经济学院", "集成电路学院", "药学院", "公共卫生学院", "基础医学院",
            "材料科学与工程学院", "建筑与城市规划学院", "船舶与海洋工程学院", "武汉光电国家研究中心", "体育学院",
            "教育科学研究院", "生殖健康研究所", "软件学院"
        ]
        for i, school in enumerate(self.schools):
            self.table.insertRow(i)
            seq_item = NumericTableWidgetItem(str(i + 1))
            self.table.setItem(i, 0, seq_item)
            self.table.setItem(i, 1, QTableWidgetItem(school))
            self.table.setItem(i, 2, QTableWidgetItem(""))
            self.table.setItem(i, 3, QTableWidgetItem(""))

        # 底部按钮
        self.add_row_btn = self.create_button("添加行", "#4caf50")
        self.del_row_btn = self.create_button("删除行", "#e57373")
        self.sort_btn = self.create_button("按序号排序", "#64b5f6")
        self.output_btn = self.create_button("输出为Excel", "#81c784")

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(self.add_row_btn)
        btn_layout.addWidget(self.del_row_btn)
        btn_layout.addWidget(self.sort_btn)
        btn_layout.addWidget(self.output_btn)

        main_layout = QVBoxLayout(self.main_page)
        header_label = QLabel("基团团费整理系统", self)
        header_label.setStyleSheet("font-size: 28px; font-weight: bold; padding: 15px;")
        header_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(header_label)
        main_layout.addWidget(self.table)
        main_layout.addLayout(btn_layout)

        # 信号
        self.add_row_btn.clicked.connect(self.add_row)
        self.del_row_btn.clicked.connect(self.delete_row)
        self.sort_btn.clicked.connect(self.sort_by_index)
        self.output_btn.clicked.connect(self.output_to_excel)

    def init_option_page(self):
        self.year_input = QLineEdit(self)
        self.month_input = QLineEdit(self)
        self.day_input = QLineEdit(self)

        self.year_input.setPlaceholderText("请输入团费年份（仅数字）")
        self.month_input.setPlaceholderText("请输入团费月份（仅数字）")
        self.day_input.setPlaceholderText("请输入落款日期（x年x月x日）")

        self.year_input.setStyleSheet("color: black; background-color: #e8f5e9; padding: 10px; font-size: 16px;")
        self.month_input.setStyleSheet("color: black; background-color: #e8f5e9; padding: 10px; font-size: 16px;")
        self.day_input.setStyleSheet("color: black; background-color: #e8f5e9; padding: 10px; font-size: 16px;")

        option_layout = QVBoxLayout(self.option_page)
        header_label = QLabel("选项设置", self)
        header_label.setStyleSheet("font-size: 28px; font-weight: bold; padding: 15px;")
        header_label.setAlignment(Qt.AlignCenter)
        option_layout.addWidget(header_label)

        input_layout = QHBoxLayout()
        input_layout.addWidget(self.year_input)
        input_layout.addWidget(self.month_input)
        input_layout.addWidget(self.day_input)
        option_layout.addLayout(input_layout)

    def create_button(self, text, color):
        btn = QPushButton(text)
        btn.setStyleSheet(f"background-color: {color}; color: white; padding: 10px; font-size: 16px;")
        return btn

    # ============ 业务逻辑 ============

    def add_row(self):
        current_row = self.table.currentRow()
        if current_row == -1:
            current_row = self.table.rowCount()
        self.table.insertRow(current_row)
        seq_item = NumericTableWidgetItem(str(current_row + 1))
        self.table.setItem(current_row, 0, seq_item)
        self.table.setItem(current_row, 1, QTableWidgetItem(""))
        self.table.setItem(current_row, 2, QTableWidgetItem(""))
        self.table.setItem(current_row, 3, QTableWidgetItem(""))

    def delete_row(self):
        current_row = self.table.currentRow()
        if current_row >= 0:
            self.table.removeRow(current_row)
        else:
            QMessageBox.warning(self, "警告", "请选择要删除的行！")

    def sort_by_index(self):
        self.table.sortItems(0, Qt.AscendingOrder)

    def output_to_excel(self):
        row_count = self.table.rowCount()
        data = []
        for row in range(row_count):
            row_data = []
            for col in range(3):
                item = self.table.item(row, col)
                cell_text = item.text() if item else ""
                row_data.append(cell_text)

            month_value = self.month_input.text()
            year_value = self.year_input.text()
            day_value = self.day_input.text()

            supplement_item = self.table.item(row, 3)
            if supplement_item and "补" in supplement_item.text():
                month_value = supplement_item.text()

            row_data.extend([month_value, year_value, day_value])
            data.append(row_data)

        base_path = os.path.dirname(os.path.abspath(__file__))
        save_dir = os.path.join(base_path, "../input")
        if not os.path.exists(save_dir):
            os.makedirs(save_dir)

        file_path = os.path.join(save_dir, "data.xlsx")

        try:
            columns = ["序号", "学院", "财务金额", "团费月份", "团费年份", "落款日期"]
            df = pd.DataFrame(data, columns=columns)
            df.to_excel(file_path, index=False)

            workbook = load_workbook(file_path)
            worksheet = workbook.active
            for column in worksheet.columns:
                max_length = 0
                column_letter = get_column_letter(column[0].column)
                for cell in column:
                    try:
                        max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = (max_length + 2) * 1.2
                worksheet.column_dimensions[column_letter].width = adjusted_width
            workbook.save(file_path)

            QMessageBox.information(self, "成功", "文件已成功保存！")
        except Exception as e:
            QMessageBox.critical(self, "错误", f"保存文件时发生错误: {str(e)}")

    def save_progress(self):
        row_count = self.table.rowCount()
        data = []
        for row in range(row_count):
            row_data = []
            for col in range(4):
                item = self.table.item(row, col)
                cell_text = item.text() if item else ""
                row_data.append(cell_text)
            data.append(row_data)

        options_data = {
            "year": self.year_input.text(),
            "month": self.month_input.text(),
            "day": self.day_input.text()
        }

        save_data = {
            "table_data": data,
            "options_data": options_data
        }

        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "保存进度", "", "JSON 文件 (*.json)", options=options)
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(save_data, f, ensure_ascii=False, indent=4)
                QMessageBox.information(self, "成功", "进度已成功保存！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"保存进度时发生错误: {str(e)}")

    def load_progress(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "导入进度", "", "JSON 文件 (*.json)", options=options)
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    load_data = json.load(f)

                self.table.setRowCount(0)
                for row_data in load_data["table_data"]:
                    current_row = self.table.rowCount()
                    self.table.insertRow(current_row)

                    numeric_item = NumericTableWidgetItem(row_data[0])
                    self.table.setItem(current_row, 0, numeric_item)
                    for col in range(1, 4):
                        self.table.setItem(current_row, col, QTableWidgetItem(row_data[col]))

                self.year_input.setText(load_data["options_data"]["year"])
                self.month_input.setText(load_data["options_data"]["month"])
                self.day_input.setText(load_data["options_data"]["day"])

                QMessageBox.information(self, "成功", "进度已成功导入！")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导入进度时发生错误: {str(e)}")

    def update_row_numbers(self):
        for row in range(self.table.rowCount()):
            self.table.setItem(row, 0, NumericTableWidgetItem(str(row + 1)))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    # 浅色主题(light_blue.xml), 若要暗色可换 'dark_teal.xml' 或其它
    apply_stylesheet(app, theme='light_blue.xml')

    window = ModernTableApp()
    window.show()
    sys.exit(app.exec_())
