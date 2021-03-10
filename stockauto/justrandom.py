
import win32com.client
import pandas as pd
import numpy as np

# This page is for getting the bollinger band.

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

testCode = instCpStockCode.NameToCode('삼성전자')
nameCode = instCpStockCode.CodetoName(testCode)
print(testCode)
print(nameCode)

instStockChart.SetInputValue(0, testCode)
instStockChart.SetInputValue(1, ord('1'))
instStockChart.SetInputValue(2, 20210309)
instStockChart.SetInputValue(3, 20210101)
#instStockChart.SetInputValue(4, 78)
instStockChart.SetInputValue(5, (0, 5, 8))
instStockChart.SetInputValue(6, ord('D'))
instStockChart.SetInputValue(7, 1)
instStockChart.SetInputValue(9, ord('1'))
instStockChart.SetInputvalue(10, ord('1'))

# BlockRequest
instStockChart.BlockRequest()

# GetHeaderValue
numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)
# GetDataValue

dates = []
close = []
vol = []


for i in range(numData):

    #print(instStockChart.GetDataValue(0, i), end=' ')
    #print(instStockChart.GetDataValue(1, i), end=' ')
    #print(instStockChart.GetDataValue(2, i), end=' ')
    dates.append(instStockChart.GetDataValue(0, i))
    close.append(instStockChart.GetDataValue(1, i))
    vol.append(instStockChart.GetDataValue(2, i))
    #print("")

# Total number of data requested
print('numData:', numData)
# Number of fields(number of variables requested)
print('numField: ', numField)

# Fixed bollinger band
LB = []
midB = []
UB = []
bbdate = []
for i in range(len(dates)):
    sum = 0
    total = []

    if len(dates)-i > 20:
        #print('Current date and close:', dates[i], close[i])
        for j in range(20):
            sum += (close[i+j])
            total.append(close[i+j])

        mid = sum/20
        stdv = round(np.std(total), 2)
        #print('Average: {}, Stdv: {}'.format(mid, stdv))
        #print('LB: {}, Mid: {}, HB: {}'.format(mid-stdv, mid, mid+stdv))
        bbdate.append(dates[i])
        LB.append(mid-stdv)
        midB.append(mid)
        UB.append(mid+stdv)

for i in range(len(UB)):
    print('Date:', bbdate[i])
    print('LB:', LB[i])
    print('mid:', midB[i])
    print('UB:', UB[i])
    print('')





"""
Bollinger Band using 5min data in my own ways.....(Not correct method)


# Number of 5min in 1 day is equal to 77
onetime = 20
total = 0
count = 0
bollinger = []
for i in range(1, len(end)):
    count += 1
    profit = end[i] - end[i - 1]
    #print(dates[i], times[i], end[i], profit, profit**2)
    total += profit ** 2
    if count == oneday-1:
        #print('date: {}, bol: {} '.format(dates[i], total/oneday))
        bollinger.append(total/oneday)
        count = 0
        total = 0

temp = bollinger[0]
for i in range(1, len(bollinger)):
    temp += bollinger[i]

final = temp/20
upper = final*2
lower = final*(-2)
print(upper, final, lower)
"""
