
import win32com.client
import pandas as pd

# This page is for getting the bollinger band.

instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

testCode = instCpStockCode.NameToCode('삼성전자')
nameCode = instCpStockCode.CodetoName(testCode)
print(testCode)
print(nameCode)

instStockChart.SetInputValue(0, testCode)
instStockChart.SetInputValue(1, ord('1'))
instStockChart.SetInputValue(2, 20210224)
instStockChart.SetInputValue(3, 20210126)
#instStockChart.SetInputValue(4, 78)
instStockChart.SetInputValue(5, (0, 1, 5))
instStockChart.SetInputValue(6, ord('m'))
instStockChart.SetInputValue(7, 5)
instStockChart.SetInputValue(9, ord('1'))
instStockChart.SetInputvalue(10, ord('1'))

# BlockRequest
instStockChart.BlockRequest()

# GetHeaderValue
numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)
# GetDataValue

dates = []
times = []
end = []


for i in range(numData):

    #print(instStockChart.GetDataValue(0, i), end=' ')
    #print(instStockChart.GetDataValue(1, i), end=' ')
    #print(instStockChart.GetDataValue(2, i), end=' ')
    dates.append(instStockChart.GetDataValue(0, i))
    times.append(instStockChart.GetDataValue(1, i))
    end.append(instStockChart.GetDataValue(2, i))
    #print("")

# Total number of data requested
print('numData:', numData)
# Number of fields(number of variables requested)
print('numField: ', numField)


# Number of 5min in 1 day is equal to 77
oneday = 77
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
