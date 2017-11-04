''' 
    此前，由于订阅了node的更新，邮箱里有2万多邮件，客户端用不了导致这个邮箱几乎废了。

    由于163邮箱每页只能最多显示100个邮件，所以打算用我最省力的办法，写一个按键精灵来趁中饭的时间删除（其实几分钟就删完了）

    1. 打开Chrome，打开邮箱，到收件箱界面
    2. 打开spy++ （visual studio 附赠，或直接上网下载），获取chrome窗口的类名和窗口名，在我这里是 'Chrome_WidgetWin_1' 和 '网易邮箱6.0版 - Google Chrome'
    3. 获得窗口的句柄(spy++其实也可以直接获取句柄，通用起见，不要直接使用句柄)
    4. 循环向窗口的全选和删除按钮位置发送左键按下和弹起信息。（由于chrome内的按钮并不是windows本地窗口，所以MK_LBUTTONCLICK是不能起到预期的作用的）

    本工具依赖win32gui，这个模块相比其他如win32api的好处是可以直接pip安装。
    win32gui相比win32api没有提供setcursorpos的接口，其实是很好的，因为模拟鼠标按下并不需要真的去移动屏幕上的鼠标，而只需要在传递信息时带上预期的应点击位置即可。
'''
import win32gui
import win32con
import time

if __name__ == '__main__':
    CHROMEWND = win32gui.FindWindowEx(None,None,'Chrome_WidgetWin_1','网易邮箱6.0版 - Google Chrome')
    # CHROMEWND = 0x000b0382
    # (left, top, right, bottom) = win32gui.GetWindowRect(CHROMEWND)
    # print('window at position: ({},{},{},{})'.format(left, top, right, bottom))
    LPARAMSELECT = 440 << 15 | 228
    LPARAMDELETE = 440 << 15 | 289
    LPARAMCONFIRM = 950 << 15 | 738
    while True:
        win32gui.SendMessage(CHROMEWND,win32con.WM_LBUTTONDBLCLK,win32con.MK_LBUTTON, LPARAMSELECT)
        time.sleep(0.1)
        win32gui.SendMessage(CHROMEWND,win32con.WM_LBUTTONUP,win32con.MK_LBUTTON, LPARAMSELECT)
        time.sleep(0.1)
        win32gui.SendMessage(CHROMEWND,win32con.WM_LBUTTONDBLCLK,win32con.MK_LBUTTON, LPARAMDELETE)
        time.sleep(0.1)
        win32gui.SendMessage(CHROMEWND,win32con.WM_LBUTTONUP,win32con.MK_LBUTTON, LPARAMDELETE)
        # time.sleep(0.1)
        # win32gui.SendMessage(CHROMEWND,win32con.WM_LBUTTONDBLCLK,win32con.MK_LBUTTON, LPARAMCONFIRM)
        # time.sleep(0.1)
        # win32gui.SendMessage(CHROMEWND,win32con.WM_LBUTTONUP,win32con.MK_LBUTTON, LPARAMCONFIRM)
        time.sleep(3)