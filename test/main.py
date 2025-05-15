import pyautogui
from utils import *


if __name__ == '__main__':
    templatePath = './pic/写日志.PNG'
    template = cv_imread(templatePath)
    screenshot = pyautogui.screenshot()

    xy = matchPic(screenshot, template)
    print(xy.get("x"), xy.get("y"))

    if xy is None:
        print("未找到模板图")
    else:
        pyautogui.click(xy.get("x"), xy.get("y"))
        print("已点击定时日志")

    # pyautogui.click(1341,682)