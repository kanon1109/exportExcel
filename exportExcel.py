#! /usr/bin/env python
#coding=utf-8
import win32com.client
import os

def exportExcel(path, macroName):
    nameStr = ""
    filesList = os.listdir(path)
    for file in filesList:
        filename = str(file)
        #忽略.开头的文件 ~开头的文件
        if filename.startswith(".") or filename.startswith("~") :
            continue
        #只取.xlsm文件
        if filename.endswith("xlsm") == False:
            continue
        print "filename: " + path + filename
        xlApp = win32com.client.Dispatch('Excel.Application')
        xlApp.visible = 0 # 此行设置打开的Excel表格为可见状态；忽略则Excel表格默认不可见
        xlBook = xlApp.Workbooks.Open(path + filename)  #打开文件
        strPara = xlBook.Name + '!' + macroName #文件名+宏名
        status = xlApp.ExecuteExcel4Macro(strPara) #调用export宏
        print 'status:', status
        xlBook.Close(SaveChanges=False)#关闭


#获取当前目录
d = os.getcwd() + '\\cfg\\'
print "run in document: " + d 
exportExcel(str(d), 'export()')

