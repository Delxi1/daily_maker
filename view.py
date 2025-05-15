import sys
import threading
import time

import keyboard
from PyQt5.QtWidgets import (
    QApplication, QMessageBox, QWidget, QVBoxLayout, QHBoxLayout, QLabel, QDateEdit, QCheckBox, QPushButton, QTimeEdit,
    QTextEdit, QGridLayout
)
from PyQt5.QtCore import QDate, QTime
import subprocess


class DateSelectionWindow(QWidget):
    start = None
    end = None

    def __init__(self):
        super().__init__()
        self.process = None
        self.initUI()

        self.output_text = QTextEdit(self)
        self.output_text.setReadOnly(True)
        self.layout().addWidget(self.output_text)

    def kill_process(self):
        keyboard.wait('esc')  # 监听ESC键
        self.on_stop()
        self.output_text.append("程序被快捷键终止")

    def initUI(self):
        # 创建主垂直布局
        main_layout = QVBoxLayout()

        # 创建日期和时间选择布局
        date_time_layout = QGridLayout()

        # 创建开始日期标签和输入框
        start_date_label = QLabel('开始日期:')
        self.start_date_edit = QDateEdit(self)
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        date_time_layout.addWidget(start_date_label, 0, 0)
        date_time_layout.addWidget(self.start_date_edit, 0, 1)

        # 创建结束日期标签和输入框
        end_date_label = QLabel('结束日期:')
        self.end_date_edit = QDateEdit(self)
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        date_time_layout.addWidget(end_date_label, 1, 0)
        date_time_layout.addWidget(self.end_date_edit, 1, 1)

        # 创建定时时间标签和输入框
        timing_time_label = QLabel('定时时间:')
        self.timing_time_edit = QTimeEdit(self)
        self.timing_time_edit.setTime(QTime.currentTime())
        self.timing_time_edit.setCalendarPopup(True)
        date_time_layout.addWidget(timing_time_label, 2, 0)
        date_time_layout.addWidget(self.timing_time_edit, 2, 1)

        main_layout.addLayout(date_time_layout)

        # 创建日报勾选框
        self.daily_report_checkbox = QCheckBox('日报', self, checked=True)
        main_layout.addWidget(self.daily_report_checkbox)

        # 创建按钮布局
        button_layout = QHBoxLayout()

        # 创建开始按钮
        start_button = QPushButton('开始', self)
        start_button.clicked.connect(self.on_start)
        button_layout.addWidget(start_button)

        # 创建停止按钮
        stop_button = QPushButton('停止', self)
        stop_button.clicked.connect(self.on_stop)
        button_layout.addWidget(stop_button)

        # 创建清空日志按钮
        clear_log_button = QPushButton('清空日志', self)
        clear_log_button.clicked.connect(self.on_clear_log)
        button_layout.addWidget(clear_log_button)

        main_layout.addLayout(button_layout)

        # 设置布局
        self.setLayout(main_layout)

        # 设置窗口属性
        self.setWindowTitle('自动日报机')
        self.setGeometry(300, 300, 300, 300)
        self.setStyleSheet("""
            QLabel {
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton {
                font-size: 14px;
                padding: 5px 10px;
            }
            QDateEdit, QTimeEdit {
                font-size: 14px;
            }
            QCheckBox {
                font-size: 14px;
            }
        """)
        self.show()

    def on_start(self):
        self.start = time.time()
        self.stop_flag = False

        kill_thread = threading.Thread(target=self.kill_process)
        kill_thread.daemon = True
        kill_thread.start()

        # 开始日期
        start_date = self.start_date_edit.date().toString('yyyy-MM-dd')
        # 结束日期
        end_date = self.end_date_edit.date().toString('yyyy-MM-dd')
        # 日报按钮
        is_daily_report = self.daily_report_checkbox.isChecked()
        # 定时时间
        timing_time = self.timing_time_edit.time().toString('HH:mm')

        # 判断开始日期小于等于结束日期
        if start_date > end_date:
            self.output_text.append("开始日期不能大于结束日期")
            # 报错窗口
            QMessageBox.critical(self, "错误", "开始日期不能大于结束日期")
            return

        self.output_text.append(f"开始日期: {start_date}, 结束日期: {end_date}, 定时时间: {timing_time}, 日报: {is_daily_report}")
        # 启动DingDing.py进程
        try:
            self.process = subprocess.Popen(
                ['python', 'DingDing.py', start_date, end_date, timing_time, str(is_daily_report)],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, encoding='utf-8'
            )
            threading.Thread(target=self.handle_process_output, daemon=True).start()
        except Exception as e:
            self.output_text.append(f"启动进程时出错: {e}")

    stop_flag = False
    def on_stop(self):
        self.stop_flag = True
        if self.process:
            self.process.terminate()
            self.output_text.append("DingDing.py进程已停止")
            self.process = None

            self.end = time.time()
            # self.output_text.append(f"运行时间: {self.end - self.start}秒")
            # self.output_text.append("===================")


    def handle_process_output(self):
        if not self.process:
            return

        while True:
            if self.stop_flag:
                break

            output = self.process.stdout.readline()
            if output == '' and self.process != None and self.process.poll() is not None:
                break
            if output:
                output = output.strip()
                self.output_text.append(output)
                if output == "DingDing.py 任务完成":  # 检查通知消息
                    self.output_text.append("日报任务 完成通知")

        if self.process:
            returncode = self.process.wait()
            if returncode == 0:  # 进程正常结束
                self.output_text.append("进程正常结束")
            else:
                stderr = self.process.stderr.read()
                self.output_text.append(f"进程异常结束 (返回码: {returncode})\n{stderr}")

        self.process = None  # 重置进程对象
        self.end = time.time()
        self.output_text.append(f"运行时间: {self.end - self.start}秒")
        self.output_text.append("===================")

    def on_clear_log(self):
        self.output_text.clear()

    def closeEvent(self, event):
        """窗口关闭时终止DingDing进程"""
        if self.process:
            self.process.terminate()
            self.output_text.append("窗口关闭时终止了DingDing.py进程")

        event.accept()  # 接受关闭事件


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DateSelectionWindow()
    sys.exit(app.exec_())