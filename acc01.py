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
import xlwt #用這個來寫Excel

def txt2xls(filename,xlsname):  #txt轉excel的函式，filename 是要被轉的txt，xlsname是要存的excel
    f = open(filename, encoding = 'utf8')   #用utf8的編碼打開下載的文字檔
    x = 1                #excel從第一行開始寫，第零行放標題
    xls=xlwt.Workbook() #宣告一個excel 工作區
    sheet = xls.add_sheet('sheet1',cell_overwrite_ok=True) #產生一個scheet 取名sheet1
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
    for line in f:
        if line[0:2] == '20':                       #因為有用資料列的開始是西元年，所以開頭是20的才轉換
            sheet.write(x,0,line[0:8].strip())      #交易日
            sheet.write(x,1,line[8:12].strip())     #序號
            sheet.write(x,2,line[12:16].strip())    #時間
            sheet.write(x,3,line[16:23].strip())    #票號
            sheet.write(x,4,line[23:39].strip())    #支出金額
            sheet.write(x,5,line[39:55].strip())    #收入金額
            sheet.write(x,6,line[55:71].strip())    #餘額
            sheet.write(x,7,line[71:79].strip())    #摘要
            sheet.write(x,8,line[79:85].strip())    #股票代號
            sheet.write(x,9,line[85:].strip())      #摘要明細
            x += 1 #excel 要寫的行號加1
    f.close()
    xls.save(xlsname+'.xls') #Excel 存檔

if __name__ == "__main__":
    filename = sys.argv[1]
    xlsname  = sys.argv[2]
    txt2xls(filename,xlsname)