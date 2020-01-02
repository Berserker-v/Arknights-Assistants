'''
Arknights.py
Author: DY
Data: 2019.08.28
- 2020.1.2 修改体力药检测坐标
'''

import time
import random
import win32api
import win32con
import win32gui
from PIL import Image
from PIL import ImageGrab

# import orc_num

# 鼠标移动
def mouse_move(x, y, delay=0):
    win32api.SetCursorPos((x, y))

    if delay == 0 :
        time.sleep(random.random()*5 + 1)
    else:
        time.sleep(random.random()*5 + delay)

    mouse_click()
# end mouse_move

# 左键单击
def mouse_click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN |
                         win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
# end mouse_click

# 获取屏幕分辨率
def resolution():
    return win32api.GetSystemMetrics(0), win32api.GetSystemMetrics(1)
# end resolution

# 检测窗口是否打开
def check_window():
    win_name = u'明日方舟 - MuMu模拟器'
    handle = win32gui.FindWindow(0, win_name)
    if handle:
        return False
    else:
        print("未检测到模拟器...")
        return True
# end check_window

# 获取界面信息
def get_window_info():
    win_name = u'明日方舟 - MuMu模拟器'
    handle = win32gui.FindWindow(0, win_name)
    return win32gui.GetWindowRect(handle)
# end get_window_info

# 计算坐标
def cal_pos_x(x):
    x = float(x)
    return int(x*width)+origin_x
# end cal_pos_x

def cal_pos_y(y):
    y = float(y)
    return int(y*heigth)+origin_y
# end cal_pos_y

def cal_pos(x,y):
    return [cal_pos_x(x), cal_pos_y(y)]
# end cal_pos

### Start
# 注释给出的坐标皆为显示器分辨率为1920x1080窗口大小为1440x899下计算出的数值
print("Arknights Assist Start...")

while check_window():
    time.sleep(5)
    
win_pos = get_window_info()

print("Win Coordinates: {}".format(win_pos))

global width, heigth, origin_x, origin_y
origin_x = win_pos[0]
origin_y = win_pos[1]
width = win_pos[2] - win_pos[0]
heigth = win_pos[3] - win_pos[1]
print("Win Size: {} x {}".format(width, heigth))

while True:
    # 输入刷图次数
    epoch = input("刷图次数(按q退出): ")
    if epoch == 'q':
        print("程序退出")
        break

    isdrag = False
    # 输入是否吃药/碎石
    drag = input("是否吃药/碎石(y/n): ")
    if drag == 'y':
        isdrag = True

    # 输入是否退出
    close = input("作战结束是否退出(y/n): ")
    isclose = False
    if close == 'y':
        isclose = True

    # intellect = 0
    # intellect_last = 0
    # cost = 0

    for i in range(0,int(epoch)):

        ### 开始行动
        # Relative Coordinates (1285,776)
        mouse_move(cal_pos_x(0.892), cal_pos_y(0.863), 2)
        time.sleep(2)

        # 检测是否需要吃药
        # Relative Coordinates (27,376,40,390)
        sanity_img = ImageGrab.grab((cal_pos_x(0.029),cal_pos_y(0.528),cal_pos_x(0.029)+1,cal_pos_y(0.528)+1))
        print(sanity_img.getcolors())
        sanity_rbg = sanity_img.getcolors()[0]
        if isdrag and (sanity_rbg[1] == (176,176,176)):
            # Relative Coordinates (1227,689)
            mouse_move(cal_pos_x(0.852), cal_pos_y(0.766), 2)
            time.sleep(2)
            # Relative Coordinates (1285,776)
            mouse_move(cal_pos_x(0.892), cal_pos_y(0.863), 2)
            time.sleep(2)
        else:
            pass

        # Relative Coordinates (1236,621)
        mouse_move(cal_pos_x(0.858), cal_pos_y(0.691), 2)
        time.sleep(30)

        ### 检测敌人数旁Logo判定结束
        count = 0
        while True:
            time.sleep(5)
            # Relative Coordinates (782,81,783,82)
            img = ImageGrab.grab((cal_pos_x(0.543),cal_pos_y(0.090),cal_pos_x(0.543)+1,cal_pos_y(0.090)+1))
            img_rbg = img.getcolors()[0]
            # print(img_rbg)
            if img_rbg[1] != (255, 255, 255):
                print("作战结束 完成({}/{})".format(i+1, epoch))
                time.sleep(10)
                # 检测是否升级
                # Relative Coordinates (1293,418)
                img = ImageGrab.grab((cal_pos_x(0.898),cal_pos_y(0.465),cal_pos_x(0.898)+1,cal_pos_y(0.465)+1))
                img_rbg = img.getcolors()[0]
                # print(img_rbg)
                if img_rbg[1] == (45,44,43):
                    print("Level Up")
                    # Relative Coordinates (40,814)
                    mouse_move(cal_pos_x(0.028), cal_pos_y(0.905), 3)
                # Relative Coordinates (40,814)
                mouse_move(cal_pos_x(0.028), cal_pos_y(0.905), 2)
                time.sleep(4)
                break
            else:
                if count%3 == 0:
                    print('Running.  '+'\b'*10, end='')
                elif count%3 == 1:
                    print('Running.. '+'\b'*10, end='')
                elif count%3 == 2:
                    print('Running...'+'\b'*10, end='')
                if count > 20:
                    count = 0     
                time.sleep(5)
                count += 1

    if isclose:
        # Relative Coordinates (1422,17)
        mouse_move(cal_pos_x(0.988), cal_pos_y(0.019),5)
        # Relative Coordinates (793,567)
        mouse_move(cal_pos_x(0.551), cal_pos_y(0.631),2)
        print("程序退出")
        break


