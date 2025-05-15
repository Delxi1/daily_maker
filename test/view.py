import sys
import threading

import keyboard
from PyQt5.QtWidgets import QApplication, QMessageBox, QWidget, QVBoxLayout, QLabel, QDateEdit, QCheckBox, QPushButton, \
    QTimeEdit
from PyQt5.QtCore import QDate, QTime
import multiprocessing
import subprocess


class DateSelectionWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.process = None
        self.initUI()

        kill_thread = threading.Thread(target=self.kill_process)
        kill_thread.daemon = True
        kill_thread.start()

    def kill_process(self):
        keyboard.wait('esc')  # 监听 ESC 键
        self.on_stop()
        print("程序被快捷键终止")

    def initUI(self):
        # 创建垂直布局
        layout = QVBoxLayout()

        # 创建开始日期标签和输入框
        start_date_label = QLabel('开始日期:')
        self.start_date_edit = QDateEdit(self)
        self.start_date_edit.setDate(QDate.currentDate())
        self.start_date_edit.setCalendarPopup(True)
        layout.addWidget(start_date_label)
        layout.addWidget(self.start_date_edit)

        # 创建结束日期标签和输入框
        end_date_label = QLabel('结束日期:')
        self.end_date_edit = QDateEdit(self)
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.setCalendarPopup(True)
        layout.addWidget(end_date_label)
        layout.addWidget(self.end_date_edit)

        # 创建定时时间标签和输入框
        timing_time_label = QLabel('定时时间:')
        self.timing_time_edit = QTimeEdit(self)
        self.timing_time_edit.setTime(QTime.currentTime())
        self.timing_time_edit.setCalendarPopup(True)
        layout.addWidget(timing_time_label)
        layout.addWidget(self.timing_time_edit)

        # 创建日报勾选框
        self.daily_report_checkbox = QCheckBox('日报', self, checked=True)
        layout.addWidget(self.daily_report_checkbox)

        # 创建开始按钮
        start_button = QPushButton('开始', self)
        start_button.clicked.connect(self.on_start)
        layout.addWidget(start_button)

        # 创建停止按钮
        stop_button = QPushButton('停止', self)
        stop_button.clicked.connect(self.on_stop)
        layout.addWidget(stop_button)

        # 设置布局
        self.setLayout(layout)

        # 设置窗口属性
        self.setWindowTitle('自动日报机')
        self.setGeometry(300, 300, 300, 250)
        self.show()

    def on_start(self):
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
            print("开始日期不能大于结束日期")
            # 报错窗口
            QMessageBox.critical(self, "错误", "开始日期不能大于结束日期")
            print(f"开始日期: {start_date}, 结束日期: {end_date}, 定时时间: {timing_time}, 日报: {is_daily_report}")
            return

        print(f"开始日期: {start_date}, 结束日期: {end_date}, 定时时间: {timing_time}, 日报: {is_daily_report}")
        # 启动 DingDing.py 进程
        try:
            self.process = subprocess.Popen(
                ['python', 'DingDing.py', start_date, end_date, timing_time, str(is_daily_report)])
        except Exception as e:
            print(f"启动进程时出错: {e}")

    def on_stop(self):
        if self.process:
            self.process.terminate()
            print("DingDing.py 进程已停止")
            self.process = None

    def handle_process_output(self):
        if not self.process:
            return

        # 读取标准输出
        stdout, stderr = self.process.communicate()

        # 等待进程结束
        returncode = self.process.wait()

        if returncode == 0:  # 进程正常结束
            # 提取结果部分
            result_start = stdout.find("===== 执行结果 =====")
            result_end = stdout.find("===== 结果结束 =====")

            if result_start != -1 and result_end != -1:
                result_start += len("===== 执行结果 =====")
                result = stdout[result_start:result_end].strip()
                self.result = result
                self.result_label.setText(result)
            else:
                self.result_label.setText("未找到结果标记")
        else:
            self.result_label.setText(f"进程异常结束 (返回码: {returncode})\n{stderr}")

        self.process = None  # 重置进程对象


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = DateSelectionWindow()
    sys.exit(app.exec_())
