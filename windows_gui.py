#!/usr/bin/env python
# -*- coding: utf-8 -*-

'''
プログラムを動かすためには
pip install pyserial
pip install pandas
pip install openpyxl
pip install xrld
pip install pygame
'''

'''
建物の状態は1の時sleep,0の時監視開始とする
センサ反応時はセンサ番号を受信する　例b 
初回起動時にセンサ反応の有無を確認する
サーバーに送信時"建物番号(3桁)"+'0'or'1'とする
'''

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
#import sound

IM920 = IM920.IM920WinClass()

IM920.Rdnn()
IM920.Write('ecio')
IM920.startWindows()


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

        self.Static2 = tk.Label(self.main_frame, text=u'1号館')
        self.Static2.place(x=50, y=20)

        self.Static3 = tk.Label(self.main_frame, text=u'新2号館')
        self.Static3.place(x=50, y=60)

        self.Static3 = tk.Label(self.main_frame, text=u'5号館')
        self.Static3.place(x=50, y=100)

        self.Static3 = tk.Label(self.main_frame, text=u'8号館')
        self.Static3.place(x=50, y=140)

        self.Static3 = tk.Label(self.main_frame, text=u'9号館')
        self.Static3.place(x=50, y=180)

        self.Static3 = tk.Label(self.main_frame, text=u'10号館')
        self.Static3.place(x=50, y=220)


        self.Static4 = tk.Label(self.main_frame, text=u'一括制御ボタン')
        self.Static4.place(x=50, y=300)

       
        # Buttonを設置する
        self.Button1 = tk.Button(self.main_frame, text=u"監視開始", width =15,command = lambda : [self.SendData('001','0'), self.switchButtonState('1','0')])
        self.Button1.place(x=150, y=20)

        self.Button2 = tk.Button(self.main_frame, text=u'監視停止',width=15,command = lambda : [self.SendData('001','1'), self.switchButtonState('1','1')])
        self.Button2.place(x=265, y=20)

        self.Button3 = tk.Button(self.main_frame, text=u'監視開始', width=15,command = lambda : [self.SendData('002','0'), self.switchButtonState('2','0')])
        self.Button3.place(x=150, y=60)

        self.Button4 = tk.Button(self.main_frame, text=u'監視停止',width=15,command = lambda : [self.SendData('002','1'), self.switchButtonState('2','1')])
        self.Button4.place(x=265, y=60)

        self.Button5 = tk.Button(self.main_frame, text=u'監視開始', width=15,command = lambda : [self.SendData('003','0'), self.switchButtonState('3','0')])
        self.Button5.place(x=150, y=100)

        self.Button6 = tk.Button(self.main_frame, text=u'監視停止',width=15,command = lambda : [self.SendData('003','1'), self.switchButtonState('3','1')])
        self.Button6.place(x=265, y=100)

        self.Button7 = tk.Button(self.main_frame, text=u'監視開始', width=15,command = lambda : [self.SendData('004','0'), self.switchButtonState('4','0')])
        self.Button7.place(x=150, y=140)

        self.Button8 = tk.Button(self.main_frame, text=u'監視停止',width=15,command = lambda : [self.SendData('004','1'), self.switchButtonState('4','1')])
        self.Button8.place(x=265, y=140)

        self.Button9 = tk.Button(self.main_frame, text=u'監視開始', width=15,command = lambda : [self.SendData('005','0'), self.switchButtonState('5','0')])
        self.Button9.place(x=150, y=180)

        self.Button10 = tk.Button(self.main_frame, text=u'監視停止',width=15,command = lambda : [self.SendData('005','1'), self.switchButtonState('5','1')])
        self.Button10.place(x=265, y=180)

        self.Button11 = tk.Button(self.main_frame, text=u'監視開始', width=15,command = lambda : [self.SendData('006','0'), self.switchButtonState('6','0')])
        self.Button11.place(x=150, y=220)

        self.Button12 = tk.Button(self.main_frame, text=u'監視停止',width=15,command = lambda : [self.SendData('006','1'), self.switchButtonState('6','1')])
        self.Button12.place(x=265, y=220)

        self.ButtonON = tk.Button(self.main_frame, text=u'監視開始',width=15,command = lambda : [self.AllSendData('0'),self.allsendButton('0')])
        self.ButtonON.place(x=150,y=300)

        self.ButtonOFF = tk.Button(self.main_frame, text=u'監視停止',width=15,command = lambda : [self.AllSendData('1'),self.allsendButton('1')])
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
        self.back_button = tk.Button(self.frame1, text=u"管理画面", command=lambda : [self.frame1.destroy(),self.newflame(),
                                                                                  self.changePage(self.main_frame)])
        self.back_button.pack()
#-------------------------------------------------------------------------------------------------------------------------------------------------------------
        #main_frameを一番上に表示

        self.main_frame.tkraise()

        #マルチスレッドで受信する
        self.thread1 = threading.Thread(target = self.Read920)
        self.thread1.start()

#---------------------------------------------------------------------------------------------------------------------------------------------------------------
    
    #新しいframeを生成する関数
    def newflame(self):
        self.frame1 = tk.Frame()
        self.frame1.grid(row=0, column=0, sticky="nsew")

        #履歴ファイルにジャンプするボタン
        self.history_button2 = tk.Button(self.frame1, text=u"履歴", bg ="#00ff7f", command=lambda : [self.openExcel()])
        self.history_button2.pack()

        #フレーム1からmainフレームに戻るボタン
        self.back_button = tk.Button(self.frame1, text=u"管理画面", command=lambda : [self.frame1.destroy(),self.newflame(),
                                                                                  self.changePage(self.main_frame)])
        self.back_button.pack()

    #画面遷移用の関数
    def changePage(self, page):
        page.tkraise()
    
    #Excelファイルを開く関数
    def openExcel(self):
        eed.openExcel()

    def SendData(self,bldnumber, mode):
    #サーバへ個別送信
        self.res = IM920.Send(bldnumber,mode)
        if self.res == 'OK':
            if mode == '0':
                state = '監視開始'
            else:
                state = '監視停止'
            
            bldnumber = 'B'+bldnumber
            #print(id)
            bldname = eed.ExtractBldName(bldnumber)
            eed.writeSendData(bldname,state)

            senddata = bldname+'は'+state+'しました'
            print(senddata)

        else:
            print("送信失敗")

    def AllSendData(self,mode):
    #一斉送信
        send = 'TXDU0001,000'+mode
        self.res = IM920.Write(send)
        if self.res == 'OK':
            if mode == '0':
                state = '監視開始'
            else:
                state = '監視停止'

            eed.writeSendData('全体',state)

            senddata = '全体に'+state+'中です'
            print(senddata)
        else:
            print("送信失敗")

    #ボタンを押したときの見た目を変更する関数
    def switchButtonState(self,select,state):
        if self.res == 'OK':
            if select == '1':
                if state == '0':         #監視開始時
                    self.Button1['state'] = tk.DISABLED
                    self.Button1['relief'] = tk.SUNKEN
                    self.Button1['bg'] ='#00ff7f'
                    self.Button1['text'] = '監視中'
                    self.Button2['state'] = tk.NORMAL
                    self.Button2['relief'] = tk.RAISED
                    self.Button2['bg'] = '#f0f0f0'
                    self.Button2['text'] = '監視停止'
                else:                        #監視停止時
                    self.Button2['state'] = tk.DISABLED
                    self.Button2['relief'] = tk.SUNKEN
                    self.Button2['bg'] = '#ff4747'
                    self.Button2['text'] = '監視停止中'
                    self.Button1['state'] = tk.NORMAL
                    self.Button1['relief'] = tk.RAISED
                    self.Button1['bg'] = '#f0f0f0'
                    self.Button1['text'] = '監視開始'
            elif select == '2':
                if state == '0':       #監視開始時
                    self.Button3['state'] = tk.DISABLED
                    self.Button3['relief'] = tk.SUNKEN
                    self.Button3['bg'] ='#00ff7f'
                    self.Button3['text'] = '監視中'
                    self.Button4['state'] = tk.NORMAL
                    self.Button4['relief'] = tk.RAISED
                    self.Button4['bg'] = '#f0f0f0'
                    self.Button4['text'] = '監視停止'
                else:                     #監視停止時
                    self.Button4['state'] = tk.DISABLED
                    self.Button4['relief'] = tk.SUNKEN
                    self.Button4['bg'] = '#ff4747'
                    self.Button4['text'] = '監視停止中'
                    self.Button3['state'] = tk.NORMAL
                    self.Button3['relief'] = tk.RAISED
                    self.Button3['bg'] = '#f0f0f0'
                    self.Button3['text'] = '監視開始'
            elif select == '3':
                if state == '0':       #監視開始時
                    self.Button5['state'] = tk.DISABLED
                    self.Button5['relief'] = tk.SUNKEN
                    self.Button5['bg'] ='#00ff7f'
                    self.Button5['text'] = '監視中'
                    self.Button6['state'] = tk.NORMAL
                    self.Button6['relief'] = tk.RAISED
                    self.Button6['bg'] = '#f0f0f0'
                    self.Button6['text'] = '監視停止'
                else:                     #監視停止時
                    self.Button6['state'] = tk.DISABLED
                    self.Button6['relief'] = tk.SUNKEN
                    self.Button6['bg'] = '#ff4747'
                    self.Button6['text'] = '監視停止中'
                    self.Button5['state'] = tk.NORMAL
                    self.Button5['relief'] = tk.RAISED
                    self.Button5['bg'] = '#f0f0f0'
                    self.Button5['text'] = '監視開始'
            elif select == '4':
                if state == '0':       #監視開始時
                    self.Button7['state'] = tk.DISABLED
                    self.Button7['relief'] = tk.SUNKEN
                    self.Button7['bg'] ='#00ff7f'
                    self.Button7['text'] = '監視中'
                    self.Button8['state'] = tk.NORMAL
                    self.Button8['relief'] = tk.RAISED
                    self.Button8['bg'] = '#f0f0f0'
                    self.Button8['text'] = '監視停止'
                else:                     #監視停止時
                    self.Button8['state'] = tk.DISABLED
                    self.Button8['relief'] = tk.SUNKEN
                    self.Button8['bg'] = '#ff4747'
                    self.Button8['text'] = '監視停止中'
                    self.Button7['state'] = tk.NORMAL
                    self.Button7['relief'] = tk.RAISED
                    self.Button7['bg'] = '#f0f0f0'
                    self.Button7['text'] = '監視開始'
            elif select == '5':
                if state == '0':       #監視開始時
                    self.Button9['state'] = tk.DISABLED
                    self.Button9['relief'] = tk.SUNKEN
                    self.Button9['bg'] ='#00ff7f'
                    self.Button9['text'] = '監視中'
                    self.Button10['state'] = tk.NORMAL
                    self.Button10['relief'] = tk.RAISED
                    self.Button10['bg'] = '#f0f0f0'
                    self.Button10['text'] = '監視停止'
                else:                     #監視停止時
                    self.Button10['state'] = tk.DISABLED
                    self.Button10['relief'] = tk.SUNKEN
                    self.Button10['bg'] = '#ff4747'
                    self.Button10['text'] = '監視停止中'
                    self.Button9['state'] = tk.NORMAL
                    self.Button9['relief'] = tk.RAISED
                    self.Button9['bg'] = '#f0f0f0'
                    self.Button9['text'] = '監視開始'
            elif select == '6':
                if state == '0':       #監視開始時
                    self.Button11['state'] = tk.DISABLED
                    self.Button11['relief'] = tk.SUNKEN
                    self.Button11['bg'] ='#00ff7f'
                    self.Button11['text'] = '監視中'
                    self.Button12['state'] = tk.NORMAL
                    self.Button12['relief'] = tk.RAISED
                    self.Button12['bg'] = '#f0f0f0'
                    self.Button12['text'] = '監視停止'
                else:                     #監視停止時
                    self.Button12['state'] = tk.DISABLED
                    self.Button12['relief'] = tk.SUNKEN
                    self.Button12['bg'] = '#ff4747'
                    self.Button12['text'] = '監視停止中'
                    self.Button11['state'] = tk.NORMAL
                    self.Button11['relief'] = tk.RAISED
                    self.Button11['bg'] = '#f0f0f0'
                    self.Button11['text'] = '監視開始'
        else:
            pass
    
    
    def allsendButton(self,state):
        if state == '0':                  #監視開始時
            self.ButtonON['relief'] = tk.SUNKEN
            self.ButtonON['bg'] = '#00ff7f'
            self.ButtonON['text'] = '監視中'
            self.ButtonOFF['relief'] = tk.RAISED
            self.ButtonOFF['bg'] = '#f0f0f0'
            self.ButtonOFF['text'] ='監視停止'
            self.switchButtonState('1','0')
            self.switchButtonState('2','0')
            self.switchButtonState('3','0')
            self.switchButtonState('4','0')
            self.switchButtonState('5','0')
            self.switchButtonState('6','0')
        else:                          #監視停止時
            self.ButtonOFF['relief'] = tk.SUNKEN
            self.ButtonOFF['bg'] = '#ff4747'
            self.ButtonOFF['text'] = '監視停止中'
            self.ButtonON['relief'] = tk.RAISED
            self.ButtonON['bg'] = '#f0f0f0'
            self.ButtonON['text'] = '監視開始'
            self.switchButtonState('1','1')
            self.switchButtonState('2','1')
            self.switchButtonState('3','1')
            self.switchButtonState('4','1')
            self.switchButtonState('5','1')
            self.switchButtonState('6','1')
 
     #IM920からのデータを受信してページを切り替える関数
    def Read920(self):
        while True:
            try:
                rx_data = IM920.Read()                        # 受信処理           
                if rx_data is not None:                          # 11は受信データのノード番号+RSSI等の長さ
                    print(rx_data)
                    if (rx_data[2]==',' and    
                        rx_data[7]==',' and rx_data[10]==':'):
                        rx_message = rx_data[11:15]
                        startmessage = rx_data[13:]
                        print(rx_data)

                        if rx_message[0] == 'S':

                            BldNumber = eed.ExtractBldNumber(rx_message)
                            bldnumber = BldNumber[1:]
                            bld = BldNumber[3]
                            self.SendData(bldnumber,'1')
                            self.switchButtonState(bld,'1')
                            RoomName = eed.ExtractRoomName(rx_message)

                            eed.writeReceiveData(RoomName,rx_message)

                            self.changePage(self.frame1)
                            self.Label1 = tk.Label(self.frame1, text= "『"+RoomName +"』で人物検知しました" , font=('Helvetica','20'))
                            self.Label1.pack(anchor='center', expand=True)

                            sound.Sound()

                        elif rx_message[:2] == 'st':
                            self.res = 'OK'
                            print(startmessage)
                            count = len(startmessage)
                            for num in range(count):
                                self.switchButtonState(str(num+1),startmessage[num])

                        else:
                            self.res = 'OK'
                            bld = rx_message[2]
                            state = rx_message[3]
                            if bld == '0':
                                self.allsendButton(state)
                            else:
                                self.switchButtonState(bld,state)
                    else:
                        self.main_frame.tkraise()

            except Exception:
                pass
                

if __name__ == "__main__":
    app = App()         #Appクラスをインスタンス化
    app.mainloop()      #メインループ開始
    IM920.Close()
