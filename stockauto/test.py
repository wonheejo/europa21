import win32com.client
import pandas as pd
import numpy as np
cpOhlc = win32com.client.Dispatch("CpSysDib.StockChart")
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

testCode = instCpStockCode.NameToCode('삼성전자')
nameCode = instCpStockCode.CodetoName(testCode)
print(testCode)
print(nameCode)

def bollinger(code):

    # CpStockCode = cpCodeMgr
    # StockChart = cpOhlc
    testcode = code

    cpOhlc.SetInputValue(0, testcode)
    cpOhlc.SetInputValue(1, ord('2'))
    #cpOhlc.SetInputValue(2, 20210309)
    #cpOhlc.SetInputValue(3, 20210101)
    cpOhlc.SetInputValue(4, 20)
    cpOhlc.SetInputValue(5, (0, 5, 8))
    cpOhlc.SetInputValue(6, ord('D'))
    cpOhlc.SetInputValue(7, 1)
    cpOhlc.SetInputValue(9, ord('1'))
    cpOhlc.SetInputvalue(10, ord('1'))

    # BlockRequest
    cpOhlc.BlockRequest()

    # GetHeaderValue
    numData = cpOhlc.GetHeaderValue(3)
    numField = cpOhlc.GetHeaderValue(1)

    # GetDataValue
    dates = []
    close = []
    vol = []

    for i in range(numData):
        dates.append(cpOhlc.GetDataValue(0, i))
        close.append(cpOhlc.GetDataValue(1, i))
        vol.append(cpOhlc.GetDataValue(2, i))

    sum = 0
    for i in range(len(dates)):
        sum += close[i]

    mid = sum/20
    stdv = round(np.std(close), 2)

    LB = mid-stdv
    UB = mid+stdv

    return LB, UB

lower, upper = bollinger(testCode)
print(lower, upper)