# -*- coding: utf-8 -*-
"""
Created on Mon May 28 17:01:08 2018

@author: Adu
"""
import pandas as pd
import numpy as np
import win32com
from win32com.client import *
# 新建一个基于COM对象的应用
xlApp = win32com.client.Dispatch("Excel.Application")
# 设置应用可见
xlApp.Visible = False
# 新增一个工作簿
#xlBook = xlApp.Workbooks.Add()
# 保存并关闭工作簿
#xlBook.SaveAs("F:\StockData\沪深A股列表1.xls")
#xlBook.Close()
# 打开已有的应用
xlBook = xlApp.Workbooks.Open(r"F:\StockData\欧奈尔系统.xlsm")

sh=xlBook.Worksheets[3]#第4个表

df=pd.DataFrame(sh.Range("A1:T3572"))
#sh.cells(1,3).Value="SyntaxError: can't assign to function call"
# 不保存，直接退出
xlBook.Close(SaveChanges=0)
# 关闭应用
xlApp.Quit()
print(df.type)