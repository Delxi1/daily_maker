from time import sleep
from tkinter import messagebox
from scipy.ndimage import maximum_filter
import cv2
import numpy as np
import pyautogui
import uiautomation as uia
import win32con
import win32gui


def cv_imread(filePath, flag=-1):
    cv_img = cv2.imdecode(np.fromfile(filePath, dtype=np.uint8), flag)
    return cv_img


def matchPicByPath(imgPath, templatePath, threshold=0.8):
    # 加载目标图像和模板图像
    img = cv_imread(imgPath)
    template = cv_imread(templatePath)

    return matchPic(img, template, threshold)


def matchPic(img, template, threshold=0.8, clickType=-1,isShow=False):
    # 将图像转换为灰度图，因为模板匹配通常在灰度图上进行
    gray_image = cv2.cvtColor(np.array(img), cv2.COLOR_BGR2GRAY)
    gray_template = cv2.cvtColor(np.array(template), cv2.COLOR_BGR2GRAY)

    # 获取模板图像的尺寸
    w, h = gray_template.shape[::-1]

    # 进行模板匹配
    res = cv2.matchTemplate(gray_image, gray_template, cv2.TM_CCOEFF_NORMED)

    # 设置匹配阈值
    # threshold = 0.9

    # 找到匹配位置
    loc = np.where(res >= threshold)

    # 使用NMS去重
    # max_res = maximum_filter(res, size=3)  # 定义3x3邻域
    # mask = (res == max_res) & (res >= threshold)  # 保留局部最大值
    # loc = np.where(mask)

    # 在目标图像中标记匹配位置
    if isShow:
        for pt in zip(*loc[::-1]):
            cv2.rectangle(gray_image, pt, (pt[0] + w, pt[1] + h), (0, 255, 0), 2)
        # 显示结果图像
        cv2.imshow('Matched Image', gray_image)
        cv2.waitKey(0)
        cv2.destroyAllWindows()

    # 判断img中是否找到template
    if len(loc[0]) > 0:
        # print("找到模板图")
        # img中template的中心点的坐标位置
        xyArr = []
        for pt in zip(*loc[::-1]):
            x, y = pt[0], pt[1]
            if clickType == -1:
                x = x + w / 2
                y = y + h / 2
            elif clickType == 0:
                x, y = x, y
            xyArr.append([int(x), int(y)])
        return sorted(xyArr, key=lambda pt: pt[0])
    else:
        # print("未找到模板图")
        return None


def open_write_log_tag():
    # 激活钉钉窗口
    # 使用 win32gui.FindWindow 函数查找钉钉窗口的句柄
    hwnd = win32gui.FindWindow('StandardFrame_DingTalk', None)
    # 显示钉钉窗口，参数 1 表示正常显示
    win32gui.ShowWindow(hwnd, 1)
    # 将钉钉窗口置于前台
    win32gui.SetForegroundWindow(hwnd)
    # 最大化窗口（核心 API 调用）
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
    # 可选：强制刷新窗口状态
    win32gui.UpdateWindow(hwnd)

    # 定位日志功能入口
    # 创建一个 WindowControl 对象，定位到钉钉主窗口
    main_window = uia.WindowControl(ClassName='StandardFrame_DingTalk', Name='钉钉')

    # 判断日志tag是否打开
    # children = main_window.GetChildren()
    # for child in children:
    #     print(f"名称: {child.Name}, 类名: {child.ClassName},句柄: {child.NativeWindowHandle}")
    # log_tag_handle = 395140
    # log_tag = main_window.WindowControl(hwnd=log_tag_handle)
    log_tag = main_window.WindowControl(Name='日志')
    if log_tag.Exists():
        print("日志tag已打开")
        return
    # else:
    #     print("日志tag未打开")

    # 在主窗口中查找名为 '工作台' 的按钮控件
    workbench_button = main_window.ButtonControl(Name='工作台')
    # 点击 '工作台' 按钮
    workbench_button.Click()

    # 进入日志填写页面
    # 在主窗口中查找名为 '日志' 的文本控件
    log_entry = main_window.TextControl(Name='日志')
    # 点击 '日志' 控件，进入日志填写页面
    log_entry.Click()


def matchPicCore(templatePath, threshold=0.8, clickType=-1,isShow=False):
    template = cv_imread(templatePath)
    screenshot = pyautogui.screenshot()

    return matchPic(screenshot, template, threshold, clickType,isShow)


# def clickTemplateCore(templatePath, threshold=0.8, clickType=-1):
#     xy = matchPicCore(templatePath, threshold, clickType)
#     if xy is None:
#         return False
#     else:
#         pyautogui.click(xy.get("x"), xy.get("y"))
#         return True


def clickTemplate(templatePath, threshold=0.8, clickType=-1, isClick=True, clickIndex=0,isShow=False):
    sleep(3)
    xy = None
    for i in range(20):
        # print(f"第{i + 1}次尝试点击模板图")
        xy = matchPicCore(templatePath, threshold, clickType, isShow)
        if xy is not None:
            break
        sleep(0.5)
    if xy is None:
        templateName = templatePath.split("/")[-1].split(".")[0].split("-")[0]
        response = messagebox.askyesno("警告", f"未点击到[{templateName}]按钮,请接管")
        exit(1)
    else:
        pyautogui.moveTo(xy[clickIndex][0], xy[clickIndex][1], duration=0.5)
        if isClick:
            # 点击目标位置
            pyautogui.click()


    print(f"已点击:{templatePath.split("/")[-1].split(".")[0]}")
    return xy


if __name__ == '__main__':
    # 点击时间输入框
    clickTemplate("./pic/日志tag.PNG",threshold=0.9, isClick=False)
