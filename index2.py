import win32gui, win32con
import win32api
import pyautogui
import time
import os
import random
import json
import cv2 as cv
import pyperclip


# 读取配置文件
f = open("./config.json",encoding='utf-8')
config = f.read()
f.close()
print(config)
config = json.loads(config)

# pip install opencv-python -i https://pypi.tuna.tsinghua.edu.cn/simple
# pip install opencv-contrib-python -i https://pypi.tuna.tsinghua.edu.cn/simple

key_map = {
    "0": 49, "1": 50, "2": 51, "3": 52, "4": 53, "5": 54, "6": 55, "7": 56, "8": 57, "9": 58,
    'F1': 112, 'F2': 113, 'F3': 114, 'F4': 115, 'F5': 116, 'F6': 117, 'F7': 118, 'F8': 119,
    'F9': 120, 'F10': 121, 'F11': 122, 'F12': 123, 'F13': 124, 'F14': 125, 'F15': 126, 'F16': 127,
    "A": 65, "B": 66, "C": 67, "D": 68, "E": 69, "F": 70, "G": 71, "H": 72, "I": 73, "J": 74,
    "K": 75, "L": 76, "M": 77, "N": 78, "O": 79, "P": 80, "Q": 81, "R": 82, "S": 83, "T": 84,
    "U": 85, "V": 86, "W": 87, "X": 88, "Y": 89, "Z": 90,
    'BACKSPACE': 8, 'TAB': 9, 'TABLE': 9, 'CLEAR': 12,
    'ENTER': 13, 'SHIFT': 16, 'CTRL': 17,
    'CONTROL': 17, 'ALT': 18, 'ALTER': 18, 'PAUSE': 19, 'BREAK': 19, 'CAPSLK': 20, 'CAPSLOCK': 20, 'ESC': 27,
    'SPACE': 32, 'SPACEBAR': 32, 'PGUP': 33, 'PAGEUP': 33, 'PGDN': 34, 'PAGEDOWN': 34, 'END': 35, 'HOME': 36,
    'LEFT': 37, 'UP': 38, 'RIGHT': 39, 'DOWN': 40, 'SELECT': 41, 'PRTSC': 42, 'PRINTSCREEN': 42, 'SYSRQ': 42,
    'SYSTEMREQUEST': 42, 'EXECUTE': 43, 'SNAPSHOT': 44, 'INSERT': 45, 'DELETE': 46, 'HELP': 47, 'WIN': 91,
    'WINDOWS': 91, 'NMLK': 144,
    'NUMLK': 144, 'NUMLOCK': 144, 'SCRLK': 145,
    '[': 219, ']': 221, '+': 107, '-': 109
}

def release_key(key_code):
    """
        函数功能：抬起按键
        参   数：key:按键值
    """
    win32api.keybd_event(key_code, win32api.MapVirtualKey(key_code, 0), win32con.KEYEVENTF_KEYUP, 0)
 
 
def press_key(key_code):
    """
        函数功能：按下按键
        参   数：key:按键值
    """
    win32api.keybd_event(key_code, win32api.MapVirtualKey(key_code, 0), 0, 0)
 
 
def press_and_release_key(key_code):
    """
        按一下按键
    :param key_code: 按键值，如91,代表WIN windows系统的系统按键，弹出开始菜单
    :return:
    """
    press_key(key_code)
    release_key(key_code)

def press_keys(*args):
    """
    按下组合键 支持多个字符，数组，元组类型
    :param args: 例如： ALT，TAB
    :return:
    """
    for i in args:
        if isinstance(i, str):
            press_key(key_map.get(i))
            time.sleep(0.3)
        elif isinstance(i, list):
            [press_keys(n) for n in i]
        elif isinstance(i, tuple):
            [press_keys(n) for n in i]
 
 
def release_keys(*args):
    """
    松开组合键 支持多个字符，数组，元组类型
    :param args: 例如：ALT，TAB
    :return:
    """
    for i in args:
        if isinstance(i, str):
            release_key(key_map.get(i))
        elif isinstance(i, list):
            [release_keys(n) for n in i]
        elif isinstance(i, tuple):
            [release_keys(n) for n in i]
 
 
def press_release_keys(*args):
    """
    按下停顿一秒然后松开组合键, 支持多个字符，数组，元组类型
    :param args: 例如：ALT，TAB
    :return:
    """
    press_keys(*args)
    time.sleep(0.5)
    release_keys(*args)

def match_windows(win_title):
    """
    查找指定窗口
    :param win_title: 窗口名称
    :return: 句柄列表
    """
    def callback(hwnd, hwnds):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            win_text = win32gui.GetWindowText(hwnd)
            # 模糊匹配
            if win_text.find(win_title) > -1:
                print(hwnd, win_text)
                hwnds.append(hwnd)
        return True
    hwnds = []
    win32gui.EnumWindows(callback, hwnds)  # 列出所有顶级窗口，并传递它们的指针给callback函数
    return hwnds

def win_active(winID):
    """
    激活指定窗口
    :param winID: 窗口ID
    :return:
    """
    win32gui.ShowWindow(winID, win32con.SW_SHOWNORMAL)  # SW_SHOWNORMAL 默认大小，SW_SHOWMAXIMIZED 最大化显示
    win32gui.SetForegroundWindow(winID)
    win32gui.SetActiveWindow(winID)

def clickWindows (x, y, winP):
    win32api.SetCursorPos([x + winP[0], y + winP[1]])
    #右键单击
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

def clickScreen (x, y):
    win32api.SetCursorPos([x, y])
    #右键单击
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)

# 图片匹配
def findkPic (img, confidence, xc = 0, yc = 0):
    imgPos = pyautogui.locateCenterOnScreen(img,confidence=confidence)
    if (imgPos):
        win32api.SetCursorPos([imgPos[0] + xc, imgPos[1] + yc])
        return True
    return False

def clickPic (img, confidence, xc = 0, yc = 0):
    imgPos = pyautogui.locateCenterOnScreen(img,confidence=confidence)
    if (imgPos):
        clickScreen(imgPos[0] + xc, imgPos[1] + yc)
        return True
    return False

def wheelFindImg(wheelNum, img, confidence, xc = 0, yc = 0):
    while not clickPic (img, confidence, xc, yc):
        print('寻找元素:' + img)
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL,0,0, wheelNum)
        time.sleep(0.5)

temp = match_windows("AdsPower")
print(temp)
handle = temp[0]
# 获取窗口位置
left, top, right, bottom = win32gui.GetWindowRect(handle)

winP = [left, top, right, bottom]

win_active(handle)

if (not clickPic('./s1.png', 0.9)):
    clickPic('./s2.png', 0.9)
time.sleep(1)
clickPic('./s3.png', 0.9)
time.sleep(0.5)
clickPic('./s4.png', 0.9, 10, 200)
time.sleep(0.5)
clickPic('./s5.png', 0.9)
time.sleep(2)
clickPic('./s6.png', 0.9)
time.sleep(0.5)