'''
Arknights.py
Author: DY
Data: 2019.08.28
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

    for i in range(0,int(epoch)):
        ### 开始行动
        # Relative Coordinates (1285,776)
        mouse_move(cal_pos_x(0.892), cal_pos_y(0.863), 2)
        time.sleep(2)

        # 检测是否需要吃药
        # Relative Coordinates (127,538,144,556)
        sanity_img = ImageGrab.grab((cal_pos_x(0.088),cal_pos_y(0.598),cal_pos_x(0.100),cal_pos_y(0.620)))
        sanity_rbg = sanity_img.getcolors()[0]
        # print(sanity_rbg)
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

        ### 检测信赖判定结束
        count = 0
        while True:
            # Relative Coordinates (1105,576,1106,577)
            trust_img = ImageGrab.grab((cal_pos_x(0.767),cal_pos_y(0.641),cal_pos_x(0.768),cal_pos_y(0.642)))
            trust_rgb = trust_img.getcolors()[0]
            print(trust_rgb)
            if trust_rgb[1] == (255, 150, 2):
                print("信赖提升...战斗结束 完成({}/{})".format(i+1, epoch))
                time.sleep(5)
                # Relative Coordinates (40,814)
                mouse_move(cal_pos_x(0.028), cal_pos_y(0.905), 2)
                time.sleep(2)
                break
            else:
                if count%3 == 0:
                    print('Running.  '+'\b'*10, end='')
                elif count%3 == 1:
                    print('Running.. '+'\b'*10, end='')
                elif count%3 == 2:
                    print('Running...'+'\b'*10, end='')
                if count > 60:
                    count = 0
                    # Relative Coordinates (40,814)
                    mouse_move(cal_pos_x(0.028), cal_pos_y(0.905))      
                time.sleep(5)
                count += 1


