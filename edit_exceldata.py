
#pip install xlrdをする
#pip install openpyxlをする


import pandas as pd
from datetime import datetime
import openpyxl
import subprocess

#Excelファイルから特定の文字列を抽出する(部屋番号)
def pickoutdata(a):
	df  = pd.read_excel('datasheet.xlsx',sheet_name = 1)
	#print(df)
	df1 = df.query('部屋番号 == @a')
	#print(df1)
	data = df1.iat[0,1]
	#print(data)
	return data

#Excelファイルから特定の文字列を抽出する(建物番号)
def pickoutdata1(a):
	df  = pd.read_excel('datasheet.xlsx',sheet_name = 0)
	#print(df)
	df1 = df.query('建物番号 == @a')
	#print(df1)
	data = df1.iat[0,1]
	#print(data)
	return data

#Excelファイルの最終行に時間と場所を記述する(受信時)
def writedata(RoomData):
	today = datetime.today().strftime('%Y/%m/%d %H:%M:%S')
	wb = openpyxl.load_workbook('datasheet.xlsx')
	ws = wb.worksheets[2]
	maxRow = ws.max_row + 1
	#print(maxRow)
	ws.cell(row=maxRow,column=1).value= today
	ws.cell(row=maxRow,column=2).value= RoomData
	wb.save('datasheet.xlsx')


#Excelファイルの最終行に時間と場所と状態を記述する(送信時)
def writedata1(RoomData,state):
	today = datetime.today().strftime('%Y/%m/%d %H:%M:%S')
	wb = openpyxl.load_workbook('datasheet.xlsx')
	ws = wb.worksheets[3]
	maxRow = ws.max_row + 1
	#print(maxRow)
	ws.cell(row=maxRow,column=1).value= today
	ws.cell(row=maxRow,column=2).value= RoomData
	ws.cell(row=maxRow,column=3).value= state
	wb.save('datasheet.xlsx')

#Excelファイルを起動する
def openExcel():
	subprocess.Popen(['start',"datasheet.xlsx"], shell=True)


'''
if __name__ == '__main__':
	today = datetime.today().strftime('%Y/%m/%d %H:%M:%S')
	print(today)
'''