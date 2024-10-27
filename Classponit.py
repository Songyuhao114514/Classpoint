import pandas as pd
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtWidgets import QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QInputDialog
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
import sys
import os
import json
import boto3
from botocore.exceptions import NoCredentialsError, EndpointConnectionError
from datetime import datetime

# 初始化数据
data = {
    "学号": [],
    "姓名": [],
    "出勤": [],
    "仪容": [],
    "晨读": [],
    "课堂": [],
    "作业": [],
    "两操": [],
    "午休": [],
    "自习": [],
    "卫生": [],
    "总分": []
}
df = pd.DataFrame(data)

# 读取或初始化标题
def load_title():
    if os.path.exists("title.txt"):
        with open("title.txt", "r", encoding="utf-8") as file:
            return file.read().strip()
    return "班级积分管理系统"  # 默认标题

def save_title(title):
    with open("title.txt", "w", encoding="utf-8") as file:
        file.write(title)

# 读取所有学生数据
def load_all_students_data():
    if os.path.exists("all_students_data.json"):
        with open("all_students_data.json", "r", encoding="utf-8") as file:
            return json.load(file)
    return []

class ScoreManagementApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("班级积分管理系统")
        self.setGeometry(100, 100, 400, 300)

        # 设置样式
        self.setStyleSheet("""
            QWidget {
                background-color: #f0f0f0;
                font-family: '钉钉进步体'; 
                font-size: 12px; 
                font-weight: normal;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QLineEdit {
                border: 1px solid #ccc;
                border-radius: 5px;
                padding: 5px;
            }
            QLabel {
                font-weight: normal;
            }
        """)

        # 加载标题
        self.title_text = load_title()

        # 加载所有学生数据
        self.students_data = load_all_students_data()
        self.students_dict = {str(student["学号"]): student["姓名"] for student in self.students_data}

        # 创建主布局
        layout = QVBoxLayout()
        layout.setContentsMargins(3, 3, 3, 3)  # 更小的边距
        layout.setSpacing(3)  # 更小的间距
        self.setFixedSize(500, 150)
        # 学号和姓名的水平布局
        id_name_layout = QHBoxLayout()
        id_name_layout.setSpacing(3)

        # 学号输入
        self.student_id_label = QLabel("学号：")
        self.student_id_label.setFont(QtGui.QFont('钉钉进步体', 14))
        self.student_id_input = QLineEdit()
        self.student_id_input.setFixedHeight(25)  # 调整输入框高度
        self.student_id_input.editingFinished.connect(self.auto_fill_name)  # 自动填充姓名
        id_name_layout.addWidget(self.student_id_label)
        id_name_layout.addWidget(self.student_id_input)

        # 姓名输入
        self.student_name_label = QLabel("姓名：")
        self.student_name_label.setFont(QtGui.QFont('钉钉进步体', 14))
        self.student_name_input = QLineEdit()
        self.student_name_input.setFixedHeight(25)  # 调整输入框高度
        id_name_layout.addWidget(self.student_name_label)
        id_name_layout.addWidget(self.student_name_input)

        # 将水平布局添加到主布局中
        layout.addLayout(id_name_layout)

        # 类别
        self.categories = ["出勤", "仪容", "晨读", "课堂", "作业", "两操", "午休", "自习", "卫生"]

        # 按钮的水平布局
        button_layout = QHBoxLayout()
        button_layout.setSpacing(5)  # 更小的间距

        # 录入积分按钮
        self.entry_button = QPushButton("录入积分")
        self.entry_button.clicked.connect(self.input_scores)
        button_layout.addWidget(self.entry_button)

        # 设置标题按钮
        self.set_title_button = QPushButton("设置标题")
        self.set_title_button.clicked.connect(self.set_title)
        button_layout.addWidget(self.set_title_button)

        # 导出到Excel按钮
        self.export_button = QPushButton("导出到Excel")
        self.export_button.clicked.connect(self.export_to_excel)
        button_layout.addWidget(self.export_button)

        # 保存到JSON按钮
        self.save_json_button = QPushButton("保存到JSON")
        self.save_json_button.clicked.connect(self.save_all_to_json)
        button_layout.addWidget(self.save_json_button)

        # 加载JSON按钮
        self.load_json_button = QPushButton("加载JSON")
        self.load_json_button.clicked.connect(self.load_all_from_json)
        button_layout.addWidget(self.load_json_button)

        # 关于按钮
        self.about_button = QPushButton("关于")
        self.about_button.clicked.connect(self.show_about)
        button_layout.addWidget(self.about_button)

        # 添加按钮布局到主布局
        layout.addLayout(button_layout)

        # 设置窗口布局
        self.setLayout(layout)

    def show_message_box(self, title, message, icon=QMessageBox.Information):
        msg_box = QMessageBox()
        msg_box.setWindowTitle(title)
        msg_box.setText(message)
        msg_box.setIcon(icon)  # 设置图标
        font = QtGui.QFont("钉钉进步体", 12, QtGui.QFont.Normal)  # 正常字体
        msg_box.setFont(font)
        msg_box.exec_()

    def set_title(self):
        title, ok = QInputDialog.getText(self, "设置标题", "请输入Excel顶部标题：", text=self.title_text)
        if ok and title.strip():
            self.title_text = title.strip()
            save_title(self.title_text)  # 保存标题到本地文件
            self.show_message_box("标题设置", f"标题已设置为：{self.title_text}")

    def auto_fill_name(self):
        student_id = self.student_id_input.text().strip()
        if student_id in self.students_dict:
            self.student_name_input.setText(self.students_dict[student_id])
        else:
            self.student_name_input.clear()

    def input_scores(self):
        global df
        student_id = self.student_id_input.text()
        student_name = self.student_name_input.text()

        if not student_id or not student_name:
            self.show_message_box("输入错误", "学号和姓名不能为空", QMessageBox.Warning)
            return

        if student_id not in df['学号'].values:
            new_row = [student_id, student_name] + [0] * len(self.categories) + [0]
            df.loc[len(df)] = new_row

        for category in self.categories:
            score, ok = QInputDialog.getInt(self, f"{category} 积分录入", f"请输入 {category} 的加/扣分（正数为加分，负数为扣分）")
            if ok:
                df.loc[df['学号'] == student_id, category] += score

        df["总分"] = df[self.categories].sum(axis=1)

        self.show_message_box("保存成功", f"{student_name} 的积分已更新，总分为 {df.loc[df['学号'] == student_id, '总分'].values[0]}")

        self.student_id_input.clear()
        self.student_name_input.clear()

    def export_to_excel(self):
        global df
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"{self.title_text}_{current_time}.xlsx"
        writer = pd.ExcelWriter(filename, engine='openpyxl')
        df.to_excel(writer, index=False, startrow=2)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        # 设置标题样式
        worksheet.merge_cells('A1:L1')
        worksheet['A1'] = self.title_text
        worksheet['A1'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet['A1'].font = Font(size=14, bold=True)

        max_row = len(df) + 4
        worksheet.merge_cells(f'A{max_row}:L{max_row}')
        worksheet[f'A{max_row}'] = 'Powered By Songyuhao'
        worksheet[f'A{max_row}'].alignment = Alignment(horizontal='center', vertical='center')
        worksheet[f'A{max_row}'].font = Font(size=10, italic=True, color="808080")

        writer.close()  # 使用 close() 方法来保存和关闭
        self.show_message_box("导出成功", f"积分信息已成功导出到 {filename}")

        # 上传到 S3
        self.upload_to_s3(filename)

    def upload_to_s3(self, file_path):
        # S3 连接配置
        access_key = "" # 替换为你的 access_key
        secret_key = "" # 替换为你的secret_key
        bucket_name = ""  # 替换为你的存储桶名称
        endpoint_url = "" # 替换为你的API

        # 创建 S3 客户端
        s3 = boto3.client(
            's3',
            aws_access_key_id=access_key,
            aws_secret_access_key=secret_key,
            endpoint_url=endpoint_url
        )

        # 构建文件名（带时间戳）
        now = datetime.now().strftime("%Y%m%d_%H%M%S")
        s3_file_name = f"班级积分_{now}.xlsx"

        try:
            # 上传文件
            s3.upload_file(file_path, bucket_name, s3_file_name)
            self.show_message_box("上传成功", f"文件已成功上传到 S3 存储：{s3_file_name}")
        except FileNotFoundError:
            self.show_message_box("上传失败", "错误：文件未找到，请检查路径", QMessageBox.Warning)
        except NoCredentialsError:
            self.show_message_box("上传失败", "错误：凭证错误，请检查Access Key和Secret Key", QMessageBox.Warning)
        except EndpointConnectionError:
            self.show_message_box("上传失败", "错误：无法连接到指定端点，请检查网络或端点配置", QMessageBox.Warning)

    def save_all_to_json(self):
        if df.empty:
            self.show_message_box("保存错误", "请先记录数据", QMessageBox.Warning)
            return

        filename = "all_students_data.json"
        all_data = df.to_dict(orient='records')
        with open(filename, "w", encoding="utf-8") as file:
            json.dump(all_data, file, ensure_ascii=False, indent=4)

        self.show_message_box("保存成功", f"所有学生数据已保存到 {filename}")

    def load_all_from_json(self):
        if os.path.exists("all_students_data.json"):
            with open("all_students_data.json", "r", encoding="utf-8") as file:
                all_data = json.load(file)
                global df
                df = pd.DataFrame(all_data)
                self.students_data = load_all_students_data()
                self.students_dict = {str(student["学号"]): student["姓名"] for student in self.students_data}
                self.show_message_box("加载成功", "学生数据已成功加载！")
        else:
            self.show_message_box("加载失败", "未找到JSON文件，请确认文件是否存在。", QMessageBox.Warning)

    def show_about(self):
        about_box = QMessageBox(self)
        about_box.setWindowTitle("关于")
        about_box.setIcon(QMessageBox.NoIcon)  # 删除旁边的信息图标
        about_box.setText("班级积分管理系统\n版本 Beta 1.0\n作者 Songyuhao")
        about_box.setStandardButtons(QMessageBox.Ok)

        # 设置图标
        about_box.setIconPixmap(QtGui.QPixmap("icon.ico"))

        about_box.exec_()


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = ScoreManagementApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
