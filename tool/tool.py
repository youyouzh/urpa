import time

import pyautogui


def print_current_cursor(space_seconds=0.5):
    while True:
        x, y = pyautogui.position()
        print("当前鼠标游标坐标：", x, y)
        time.sleep(space_seconds)


if __name__ == '__main__':
    print_current_cursor()
