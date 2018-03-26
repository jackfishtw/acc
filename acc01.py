#!/bin/env python
# -*- encoding: utf-8 -*-
#-------------------------------------------------------------------------------
# Purpose:     transfor SCB account txt to excel
# Author:      Jack Fish
# Created:     2018-03-26
# update:      2018-03-26
#-------------------------------------------------------------------------------
from datetime import datetime
import time
import os
import sys
import xlwt #需要的模块

def txt2xls(filename,xlsname):  #文本转换成xls的函数，filename 表示一个要被转换的txt文本，xlsname 表示转换后的文件名
    f = open(filename, encoding = 'utf8')   #打开txt文本进行读取
    x = 1                #在excel开始写的位置（y）
    y = 0                #在excel开始写的位置（x）
    xls=xlwt.Workbook()
    sheet = xls.add_sheet('sheet1',cell_overwrite_ok=True) #生成excel的方法，声明excel
    sheet.write(0,0,'交易日')
    sheet.write(0,1,'序號')
    sheet.write(0,2,'時間')
    sheet.write(0,3,'票號')
    sheet.write(0,4,'支出金額')
    sheet.write(0,5,'收入金額')
    sheet.write(0,6,'餘額')
    sheet.write(0,7,'摘要')
    sheet.write(0,8,'股票代號')
    sheet.write(0,9,'摘要明細')
    while True:  #循环，读取文本里面的所有内容
        line = f.readline() #一行一行读取
        if not line:  #如果没有内容，则退出循环
            break        
        if line[0:2] == '20':
            sheet.write(x,0,line[0:8].strip())      #交易日
     #       sheet.write(x,0,datetime.strptime((line[0:8].strip()),'%Y%m%d'))      #交易日
            sheet.write(x,1,line[8:12].strip())     #序號
            sheet.write(x,2,line[12:16].strip())    #時間
            sheet.write(x,3,line[16:23].strip())    #票號
            sheet.write(x,4,line[23:39].strip())    #支出金額
            sheet.write(x,5,line[39:55].strip())    #收入金額
            sheet.write(x,6,line[55:71].strip())    #餘額
            sheet.write(x,7,line[71:79].strip())    #摘要
            sheet.write(x,8,line[79:85].strip())    #股票代號
            sheet.write(x,9,line[85:].strip())      #摘要明細
            x += 1 #另起一行
    f.close()
    xls.save(xlsname+'.xls') #保存

if __name__ == "__main__":
    filename = sys.argv[1]
    xlsname  = sys.argv[2]
    txt2xls(filename,xlsname)