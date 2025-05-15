import datetime
import threading
import comtypes
import keyboard
from PyQt5.QtWidgets import QMessageBox
from utils import *
import sys
import time


def coreFun(start_date, timing_time):
    timingdate = start_date + " " + timing_time
    print(f"开始定{timingdate}的日志")

    # 打开钉钉,并点击工作台和日志
    open_write_log_tag()

    # 点击写日志
    clickTemplate("./pic/写日志.PNG")

    # 导入上篇
    clickTemplate("./pic/导入上篇.PNG")

    # 导入提示
    clickTemplate("./pic/导入提示.PNG")

    # 导入提示-确定
    clickTemplate("./pic/确定.PNG")

    # 向下翻页
    pyautogui.scroll(-9999)

    # 点击定时发送
    clickTemplate("./pic/定时发送1.PNG", isShow=True)

    # 选择发送时间
    clickTemplate("./pic/选择发送时间.PNG")

    # 点击时间输入框
    # clickTemplate("./pic/时间输入框.PNG", clickType=0)

    clickTemplate("./pic/选择时间.PNG", isClick=False)

    sleep(3)
    # 输入日期
    pyautogui.hotkey('ctrl', 'a')
    sleep(3)
    pyautogui.write(timingdate, interval=0.1)
    sleep(3)
    pyautogui.press('enter')

    # 保存
    clickTemplate("./pic/保存.PNG")

    # 关闭当前日志tag
    pyautogui.hotkey('ctrl', 'w')


def startFun(start_date='2025-05-13', end_date='2025-05-13', timing_time=None):
    comtypes.CoInitialize()
    try:
        i = 1
        while True:
            print(f"第{i}次")
            coreFun(start_date, timing_time)
            sleep(1)

            i = i + 1
            start_date = ((datetime.datetime.strptime(start_date, "%Y-%m-%d").date()
                           + datetime.timedelta(days=1))
                          .strftime("%Y-%m-%d"))
            if start_date > end_date:
                break
    finally:
        comtypes.CoUninitialize()  # 释放 COM 库资源


if __name__ == "__main__":
    start_date = sys.argv[1]
    end_date = sys.argv[2]
    timing_time = sys.argv[3]
    is_daily_report = sys.argv[4]
    # startFun(start_date, end_date, timing_time)
    sleep(10)
    print("DingDing.py 任务完成")  # 打印通知消息