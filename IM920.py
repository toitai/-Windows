#!/usr/bin/env python
# -*- coding: utf-8 -*-

#pip install pyserial
#pip install pyserial

import serial
import signal
import sys
import platform
from serial.tools import list_ports



class IM920WinClass:
    def __init__(self):
        if platform.system() == 'Windows':  #windows用
            ports = list_ports.comports()
            portnumber = None
            for port in ports:
                if (port.vid == 1027) and (port.pid == 24597): #プロダクトIDとベンダーIDが一致したら接続
                    portnumber = port.device
                    print("connect to " + portnumber)
                    break

            if portnumber == None:
                print("not connetc to im920!")
                sys.exit(1)
        elif platform.system() == 'Linux': #Linux用
            portnumber = '/dev/ttyUSB0'

        self.com = serial.Serial(portnumber,19200)

        self.SM = 0

        #bufferクリア
        self.com.flushInput()
        self.com.flushOutput()

    '''
    ctrl+cの命令
    '''
    def signal_handler(self, signal, frame):
        print('exit')
        sys.exit()

    '''
    IMコマンド操作
    '''
    def Write(self, cmd):
        self.SM = 1
        self.com.flushInput()
        self.com.write(cmd.encode('utf-8') + b'\r\n')
        self.com.flushOutput()
        res = self.com.readline().strip().decode('utf-8')
        print(res)
        self.SM = 0
        return res

    '''
    固有IDの読み出し
    '''
    def Rdnn(self):
        self.com.flushInput()
        self.com.write(b'RDNN'+b'\r\n')
        self.com.flushOutput()
        print(self.com.readline().strip().decode('utf-8'))
        

    '''
    windows立ち上げ時
    '''
    def startWindows(self):
        self.SM = 1
        self.com.flushInput()
        self.com.write(b'TXDU0001,star' + b'\r\n')    #文字列をバイト型に変換
        self.com.flushOutput()
        res = self.com.readline().strip().decode('utf-8')
        print(res)
        self.SM = 0

    '''
    TXDU送信
    '''
    def Send(self,id,state):
        self.SM = 1
        self.com.flushInput()
        self.com.write(b'TXDU0001,'+id.encode('utf-8') + state.encode('utf-8') + b'\r\n')    #文字列をバイト型に変換
        self.com.flushOutput()
        res = self.com.readline().strip().decode('utf-8')
        print(res)
        self.SM = 0
        return res
       
    def Close(self):
        self.com.close()

    '''
    受信
    '''
    def Read(self):
        while True:
            if self.SM == 0:
                self.com.flushInput()

                text = ""
                try:
                    text = self.com.readline().strip().decode('utf-8')  #受信　バイト型を文字列変数に変換 改行文字を除外

                except Exception:
                    pass

                return text
            

    



