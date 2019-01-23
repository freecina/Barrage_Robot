import win32api
import win32con
import win32gui
import win32clipboard 
from ctypes import *
import time
import os,sys
import random
from agandanmu import Ui_MainWindow
from PyQt5 import QtCore,QtWidgets,QtGui,sip
from PyQt5.QtCore import QTimer,QCoreApplication,Qt
from PyQt5.QtGui import QPalette
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
class Barrage(object):
    """docstring for Barrage"""
    def __init__(self):
        super(Barrage, self).__init__()
        app = QtWidgets.QApplication(sys.argv)
        self.ui = Ui_MainWindow()
        self.timer = QTimer()
        self.red = QPalette()
        self.red.setColor(QPalette.WindowText,Qt.red)
        self.ui.pushButton.clicked.connect(self.start)
        self.ui.pushButton_2.clicked.connect(self.stop)
        self.paste_key = ['ctrl','v']
        self.VK_CODE = {
            'backspace':0x08,
            'tab':0x09,
            'clear':0x0C,
            'enter':0x0D,
            'shift':0x10,
            'ctrl':0x11,
            'alt':0x12,
            'pause':0x13,
            'caps_lock':0x14,
            'esc':0x1B,
            'spacebar':0x20,
            'page_up':0x21,
            'page_down':0x22,
            'end':0x23,
            'home':0x24,
            'left_arrow':0x25,
            'up_arrow':0x26,
            'right_arrow':0x27,
            'down_arrow':0x28,
            'select':0x29,
            'print':0x2A,
            'execute':0x2B,
            'print_screen':0x2C,
            'ins':0x2D,
            'del':0x2E,
            'help':0x2F,
            '0':0x30,
            '1':0x31,
            '2':0x32,
            '3':0x33,
            '4':0x34,
            '5':0x35,
            '6':0x36,
            '7':0x37,
            '8':0x38,
            '9':0x39,
            'a':0x41,
            'b':0x42,
            'c':0x43,
            'd':0x44,
            'e':0x45,
            'f':0x46,
            'g':0x47,
            'h':0x48,
            'i':0x49,
            'j':0x4A,
            'k':0x4B,
            'l':0x4C,
            'm':0x4D,
            'n':0x4E,
            'o':0x4F,
            'p':0x50,
            'q':0x51,
            'r':0x52,
            's':0x53,
            't':0x54,
            'u':0x55,
            'v':0x56,
            'w':0x57,
            'x':0x58,
            'y':0x59,
            'z':0x5A,
            'numpad_0':0x60,
            'numpad_1':0x61,
            'numpad_2':0x62,
            'numpad_3':0x63,
            'numpad_4':0x64,
            'numpad_5':0x65,
            'numpad_6':0x66,
            'numpad_7':0x67,
            'numpad_8':0x68,
            'numpad_9':0x69,
            'multiply_key':0x6A,
            'add_key':0x6B,
            'separator_key':0x6C,
            'subtract_key':0x6D,
            'decimal_key':0x6E,
            'divide_key':0x6F,
            'F1':0x70,
            'F2':0x71,
            'F3':0x72,
            'F4':0x73,
            'F5':0x74,
            'F6':0x75,
            'F7':0x76,
            'F8':0x77,
            'F9':0x78,
            'F10':0x79,
            'F11':0x7A,
            'F12':0x7B,
            'F13':0x7C,
            'F14':0x7D,
            'F15':0x7E,
            'F16':0x7F,
            'F17':0x80,
            'F18':0x81,
            'F19':0x82,
            'F20':0x83,
            'F21':0x84,
            'F22':0x85,
            'F23':0x86,
            'F24':0x87,
            'num_lock':0x90,
            'scroll_lock':0x91,
            'left_shift':0xA0,
            'right_shift ':0xA1,
            'left_control':0xA2,
            'right_control':0xA3,
            'left_menu':0xA4,
            'right_menu':0xA5,
            'browser_back':0xA6,
            'browser_forward':0xA7,
            'browser_refresh':0xA8,
            'browser_stop':0xA9,
            'browser_search':0xAA,
            'browser_favorites':0xAB,
            'browser_start_and_home':0xAC,
            'volume_mute':0xAD,
            'volume_Down':0xAE,
            'volume_up':0xAF,
            'next_track':0xB0,
            'previous_track':0xB1,
            'stop_media':0xB2,
            'play/pause_media':0xB3,
            'start_mail':0xB4,
            'select_media':0xB5,
            'start_application_1':0xB6,
            'start_application_2':0xB7,
            'attn_key':0xF6,
            'crsel_key':0xF7,
            'exsel_key':0xF8,
            'play_key':0xFA,
            'zoom_key':0xFB,
            'clear_key':0xFE,
            '+':0xBB,
            ',':0xBC,
            '-':0xBD,
            '.':0xBE,
            '/':0xBF,
            '`':0xC0,
            ';':0xBA,
            '[':0xDB,
            '\\':0xDC,
            ']':0xDD,
            "'":0xDE,
            '`':0xC0}
        self.ui.show()
        sys.exit(app.exec_())
    def content_dict01(self):
        dict01 = []
        f = open("content01.txt")
        line = f.readline()
        while line:
            if line.strip():
                dict01.append(line.strip())
            line = f.readline()
        f.close()
        # print(dict01)
        # print(len(dict01))
        return dict01
    def content_dict02(self):
        dict02 = []
        f = open("content02.txt")
        line = f.readline()
        while line:
            if line.strip():
                dict02.append(line.strip()+' ')
            line = f.readline()
        f.close()
        # print(dict02)
        # print(len(dict02))
        return dict02
    def setcur(self,data):
    		win32api.SetCursorPos(data)
    def d_click(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
        # time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    def click(self):
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP | win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    def get_key_point(self):
        try:
            x,y = int(self.ui.x),int(self.ui.y)
            self.set_position = [x,y]
        except Exception as e:
            self.set_position = 0
        
    def key_press(self,key):
        if key!='' and isinstance(key,list):
            if key.__len__()==1:
                key = key[0]
                win32api.keybd_event(self.VK_CODE[key],0,0,0)
            else:
                for x in key:
                    win32api.keybd_event(self.VK_CODE[x],0,0,0)

    def key_up(self,key):
        if key!='' and isinstance(key,list):
            if key.__len__()==1:
                key = key[0]
                win32api.keybd_event(self.VK_CODE[key],0,win32con.KEYEVENTF_KEYUP,0)
            else:
                for x in key:
                    win32api.keybd_event(self.VK_CODE[x],0,win32con.KEYEVENTF_KEYUP,0)

    def paste(self):
        self.key_press(self.paste_key)
        # time.sleep(0.1)
        self.key_up(self.paste_key)
    def copy(self,text):
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text)
        win32clipboard.CloseClipboard()
    def submit(self):
        win32api.keybd_event(13,0,0,0)
        win32api.keybd_event(13,0,win32con.KEYEVENTF_KEYUP,0)
    def start(self):
        self.thread = MyThread(self.ui.horizontalSlider.value(),self.ui.verticalSlider.value())
        self.thread.sec_changed_signal.connect(self.message)
        self.thread.start()
        self.ui.listWidget.addItem('程序已开始运行')
    def stop(self):
        self.thread.terminate()
        self.ui.listWidget.addItem('程序已停止运行')
    def message(self,pause):
        self.get_key_point()
        if isinstance(self.set_position,list) and len(self.set_position) == 2:
            dict01 = self.content_dict01()
            dict02 = self.content_dict02()
            sample = random.sample(dict01,k=5)
            text01 = random.choice(sample) 
            TEXT = text01 + random.choice(dict02) * random.randint(0,3)
            self.copy(TEXT)
            self.setcur(self.set_position)
            self.click()
            self.paste()
            self.submit()
            Time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
            log_text = str(Time) + ' : ' + TEXT + ' pause time: ' + str(pause)
            self.ui.listWidget.addItem(log_text)
            self.ui.listWidget.setCurrentRow(self.ui.listWidget.count()-1)
        
            
            # self.ui.listWidget.addItem('发送时间随机范围：' + '[' + str(l_time) + '-' + str(h_time) + ']')
            
        
        
        else:
            self.ui.label.setText('请按F3获取输入坐标')
            self.ui.label.setPalette(self.red)
            self.thread.terminate()

    
    # def run(self):
        
class MyThread(QThread):  
  
    sec_changed_signal = pyqtSignal(int) # 信号类型：int
  
    def __init__(self, hor=100,ver=20, parent=None):  
        super().__init__(parent)
        self.ver = ver
        self.hor = hor 
  
    def run(self):
        while 1:
            l_time = int(self.hor*(1-self.ver*0.01))

            h_time = int(self.hor*(1+self.ver*0.01))
            pause = random.randint(l_time,h_time)
            self.sec_changed_signal.emit(pause)  #发射信号
            time.sleep(pause)
if __name__ == "__main__":
    robot = Barrage()
    # time.sleep(200)
    robot.run()
