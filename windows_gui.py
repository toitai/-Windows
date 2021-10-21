#!/usr/bin/env python
# -*- coding: utf-8 -*-


#コマンド入力
#IM920.Write("RDNN")

#固有ID
#IM920.Rdnn()

#文字列受信
#print(IM920.Read())

#文字列送信
#IM920.Send("0002","hogre")

#0がON
#1がOFF

import sys
import tkinter as tk
import IM920
import threading
import edit_exceldata as eed
import subprocess
import sound

IM920 = IM920.IM920WinClass()

IM920.Rdnn()
IM920.Write('ECIO')


class App(tk.Tk):
    def __init__(self,*args,**kwargs):
        tk.Tk.__init__(self,*args,**kwargs)
        # ウインドウのタイトルを定義する
        self.title(u'警備システム')
        # ウインドウサイズを定義する
        self.geometry('460x350')
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.create_widgets()
        
    def create_widgets(self):
#---------------------------------------main_frame-----------------------------------------------------------------------------------------------------------
        #メインページフレーム作成
        self.main_frame = tk.Frame()
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        # ラベルを使って文字を画面上に出す
        self.Static1 = tk.Label(self.main_frame, text=u'▼　建物の選択　▼')
        self.Static1.pack()

        self.Static2 = tk.Label(self.main_frame, text=u'8号館(0002)')
        self.Static2.place(x=50, y=20)

        self.Static3 = tk.Label(self.main_frame, text=u'新2号館(0004)')
        self.Static3.place(x=50, y=60)

        self.Static4 = tk.Label(self.main_frame, text=u'一括制御ボタン')
        self.Static4.place(x=50, y=300)

       
        # Buttonを設置する
        self.Button1 = tk.Button(self.main_frame, text=u"ON", width =15,command = lambda : [self.SendData('0002','2'), self.switchButtonState('1','ON')])
        self.Button1.place(x=150, y=20)

        self.Button2 = tk.Button(self.main_frame, text=u'OFF',width=15,command = lambda : [self.SendData('0002','1'), self.switchButtonState('1','OFF')])
        self.Button2.place(x=265, y=20)

        self.Button3 = tk.Button(self.main_frame, text=u'ON', width=15,command = lambda : [self.SendData('0004','2'), self.switchButtonState('2','ON')])
        self.Button3.place(x=150, y=60)

        self.Button4 = tk.Button(self.main_frame, text=u'OFF',width=15,command = lambda : [self.SendData('0004','1'), self.switchButtonState('2','OFF')])
        self.Button4.place(x=265, y=60)

        self.ButtonON = tk.Button(self.main_frame, text=u'ON',width=15,command = lambda : [self.AllSendData('2'),self.allsendButton('ON')])
        self.ButtonON.place(x=150,y=300)

        self.ButtonOFF = tk.Button(self.main_frame, text=u'OFF',width=15,command = lambda : [self.AllSendData('1'),self.allsendButton('OFF')])
        self.ButtonOFF.place(x=265,y=300)

        self.history_button1 = tk.Button(self.main_frame, text=u"履歴",bg = "#00ff7f", width=5,command=lambda : [self.openExcel()])
        self.history_button1.place(x=0,y=0)
#-------------------------------------------------------------------------------------------------------------------------------------------------------------
#---------------------------------------frame1----------------------------------------------------------------------------------------------------------------
        #移動先フレーム作成
        self.frame1 = tk.Frame()
        self.frame1.grid(row=0, column=0, sticky="nsew")

        #履歴ファイルにジャンプするボタン
        self.history_button2 = tk.Button(self.frame1, text=u"履歴", bg ="#00ff7f", command=lambda : [self.openExcel()])
        self.history_button2.pack()

        #フレーム1からmainフレームに戻るボタン
        self.back_button = tk.Button(self.frame1, text=u"Back", command=lambda : [self.changePage(self.main_frame),
                                                                                  self.Label1.destroy()])
        self.back_button.pack()
#-------------------------------------------------------------------------------------------------------------------------------------------------------------
        #main_frameを一番上に表示

        self.main_frame.tkraise()

        #マルチスレッドで受信する
        self.thread1 = threading.Thread(target = self.Read920)
        self.thread1.start()

#---------------------------------------------------------------------------------------------------------------------------------------------------------------

    #画面遷移用の関数
    def changePage(self, page):
        page.tkraise()
    
    #Excelファイルを開く関数
    def openExcel(self):
        eed.openExcel()

    def SendData(self,id, text):
    #個別送信
        self.res = IM920.Send(id,text)
        if self.res == 'OK':
            if text == '2':
                state = 'ON'
            else:
                state = 'OFF'
            
            id = 'B'+id
            print(id)
            RoomData = eed.pickoutdata1(id)
            eed.writedata1(RoomData,state)

            senddata = RoomData+'に'+state+'を送信しました'
            print(senddata)

        else:
            print("送信失敗")

    def AllSendData(self,text):
    #一斉送信
        send = 'TXDA'+text
        self.res = IM920.Write(send)
        if self.res == 'OK':
            if text == '2':
                state = 'ON'
            else:
                state = 'OFF'

            eed.writedata1('全体',state)

            senddata = '全体に'+state+'を送信しました'
            print(senddata)
        else:
            print("送信失敗")

    #ボタンを押したときの見た目を変更する関数
    def switchButtonState(self,select,state):
        if self.res == 'OK':
            if select == '1':
                if state == 'ON':
                    self.Button1['state'] = tk.DISABLED
                    self.Button1['relief'] = tk.SUNKEN
                    self.Button1['bg'] ='#00ff7f'
                    self.Button2['state'] = tk.NORMAL
                    self.Button2['relief'] = tk.RAISED
                    self.Button2['bg'] = '#f0f0f0'
                else:
                    self.Button2['state'] = tk.DISABLED
                    self.Button2['relief'] = tk.SUNKEN
                    self.Button2['bg'] = '#ff4747'
                    self.Button1['state'] = tk.NORMAL
                    self.Button1['relief'] = tk.RAISED
                    self.Button1['bg'] = '#f0f0f0'
            elif select == '2':
                if state == 'ON':
                    self.Button3['state'] = tk.DISABLED
                    self.Button3['relief'] = tk.SUNKEN
                    self.Button3['bg'] ='#00ff7f'
                    self.Button4['state'] = tk.NORMAL
                    self.Button4['relief'] = tk.RAISED
                    self.Button4['bg'] = '#f0f0f0'
                else:
                    self.Button4['state'] = tk.DISABLED
                    self.Button4['relief'] = tk.SUNKEN
                    self.Button4['bg'] = '#ff4747'
                    self.Button3['state'] = tk.NORMAL
                    self.Button3['relief'] = tk.RAISED
                    self.Button3['bg'] = '#f0f0f0'
        else:
            pass
    
    
    def allsendButton(self,state):
        if state == 'ON':
            self.ButtonON['relief'] = tk.SUNKEN
            self.ButtonON['bg'] = '#00ff7f'
            self.ButtonOFF['relief'] = tk.RAISED
            self.ButtonOFF['bg'] = '#f0f0f0'
            self.switchButtonState('1','ON')
            self.switchButtonState('2','ON')
        else:
            self.ButtonOFF['relief'] = tk.SUNKEN
            self.ButtonOFF['bg'] = '#ff4747'
            self.ButtonON['relief'] = tk.RAISED
            self.ButtonON['bg'] = '#f0f0f0'
            self.switchButtonState('1','OFF')
            self.switchButtonState('2','OFF')
 
     #IM920からのデータを受信してページを切り替える関数
    def Read920(self):
        while True:
            try:
                rx_data = IM920.Read()                        # 受信処理           
                if rx_data is not None:                          # 11は受信データのノード番号+RSSI等の長さ
                    print(rx_data)
                    if (rx_data[2]==',' and    
                        rx_data[7]==',' and rx_data[10]==':'):
                        rx_xbeeid = rx_data[11:14]
                        
                        xbee = 'R'+rx_xbeeid        

                        RoomData = eed.pickoutdata(xbee)

                        eed.writedata(RoomData)

                        self.changePage(self.frame1)
                        self.Label1 = tk.Label(self.frame1, text= RoomData +"で人物検知しました" , font=('Helvetica','20'))
                        self.Label1.pack(anchor='center', expand=True)

                        sound.Sound()
                    else:
                        self.main_frame.tkraise()

            except Exception:
                pass
                

if __name__ == "__main__":
    app = App()         #Appクラスをインスタンス化
    app.mainloop()      #メインループ開始
    IM920.Close()
