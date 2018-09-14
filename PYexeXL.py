# -*- coding: utf-8 -*-
"""
Created on Thu May 17 11:24:08 2018

pd.read_excel(io, sheetname=0,header=0,skiprows=None,index_col=None,names=None,
                arse_cols=None,date_parser=None,na_values=None,thousands=None, 
                convert_float=True,has_index_names=None,converters=None,dtype=None,
                true_values=None,false_values=None,engine=None,squeeze=False,**kwds)

@author: Adu
"""
import os
import win32com.client
import pandas as pd
#import numpy as np
import random

def updatexls():#更新Excel主力资金流入数据
    ts=int(random.random()*100000000)
    
    #print (ts)
    xls=win32com.client.Dispatch("Excel.Application")
    xls.visible = 0
    cwd = os.getcwd()
    Filename = os.path.join(cwd,"主力资金当日.xlsm")#cwd + "\主力资金当日.xlsm")#'F:\baiduyundownload\主力资金当日.xlsm'
    xlbook=xls.Workbooks.Open(Filename)
    xlbook.Application.Run("test",ts)#执行excel更新数据宏命令
        #xls.Save
    xlbook.Close(SaveChanges=True)#保存数据并关闭Excel
    
def updateBK():#更新当日概念板块主力资金流入数据

    #print (ts)
    xls2=win32com.client.Dispatch("Excel.Application")
    xls2.visible = 0
    cwd = os.getcwd()
    Filename = os.path.join(cwd,"概念资金流入.xlsm")#cwd + "\主力资金当日.xlsm")#'F:\baiduyundownload\主力资金当日.xlsm'
    xlbook2=xls2.Workbooks.Open(Filename)
    xlbook2.Application.Run("update0day")#执行excel更新数据宏命令
        #xls.Save
    xlbook2.Close(SaveChanges=True)#保存数据并关闭Excel  
    
    
def BKZJ():
    cwd = os.getcwd()
    excel_path = os.path.join(cwd,"概念资金流入.xlsm")#cwd + "BKZJ\主力资金当日.xlsm")#'F:\baiduyundownload\主力资金当日.xlsm'
    #excel_path = r'F:\baiduyundownload\主力资金当日.xlsm'
    df = pd.read_excel(excel_path)
    #df.head()
    #print(type(df))
    #df['今日涨跌'].astype(float)
    #df['代码'].astype(str)
    #"{:0>6d}".format(df['代码'])
    df2=(df.loc[0:9,["板块代码","板块名称","今日涨跌","主力净占比","净流入最大个股"]])
    
    #print("主力净占比小于40%，大于25%的股票")
    #print (df)
    #print(df["主力净占比"].min())
    print("实时概念板块净占比排名")
    #print(df2.sort_values(by=u'主力净占比'))#按降序排序，升序去掉ascending=False即可 True为升序
    print(df2.sort_values(by=u'主力净占比',ascending=False))#按降序排序，升序去掉ascending=False即可 True为升序
    
    
    
    
    
def toPandas():
    cwd = os.getcwd()
    excel_path = os.path.join(cwd,"主力资金当日.xlsm")#cwd + "\主力资金当日.xlsm")#'F:\baiduyundownload\主力资金当日.xlsm'
    #excel_path = r'F:\baiduyundownload\主力资金当日.xlsm'
    df = pd.read_excel(excel_path)
    #df.head()
    #print(type(df))
    df['今日涨跌'].astype(float)
    df['代码'].astype(str)
    #"{:0>6d}".format(df['代码'])
    df2=(df.loc[0:30,["名称","最新价","今日涨跌","主力净额","主力净占比","刷新时间"]])
    
    #print("主力净占比小于40%，大于25%的股票")
    #df2=(df.loc[ (df['主力净占比']>25 ) & (df['主力净占比'] < 40 )&(df['今日涨跌'] < 8), ["代码","名称","最新价","今日涨跌","主力净额","主力净占比"]])
    #print(df["主力净占比"].min())
    print("实时主力净占比排名")
    #print(df2.sort_values(by=u'主力净占比'))#按降序排序，升序去掉ascending=False即可 True为升序
    print(df2.sort_values(by=u'主力净占比',ascending=False))#按降序排序，升序去掉ascending=False即可 True为升序


def datafx():
    pass

def yejidata(Q):#获取指定报告期业绩
    #构造数据地址
    ts=int(random.random()*100000000)
    url1="http://datainterface.eastmoney.com/EM_DataCenter/JS.aspx?type=SR&sty=YJBB&fd="
    url2="&st=13&sr=-1&p=1&ps=5000&js=var%20aaaaaaaa={pages:(pc),data:[(x)]}&stat=0&rt="
    url=url1+Q+url2+str(ts)
    print(url)
    
    
def kxdata():
    #'http://quotes.money.163.com/service/chddata.html?code=1002092&start=20061208&end=20180521&fields=TCLOSE;HIGH;LOW;TOPEN;LCLOSE;CHG;PCHG;TURNOVER;VOTURNOVER;VATURNOVER;TCAP;MCAP'    
    pass
    
def main():
    #yejidata("2017-12-31")
    
    
  
    updateBK()
    print ("概念板块资金更新完毕")
    
    BKZJ()
    print ("BKZJ")
    
    updatexls()
    print("个股主力资金更新完毕")
    
    toPandas()
    print("topandas ok")
    
    datafx()
    print("datafx ok")

main()