"""

Used to get the 5min data of Samsung Electronics with a date of 2016 Feb 17th

"""
import win32com.client
import pandas as pd
import sys
instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

testCode = instCpStockCode.NameToCode('삼성전자')
nameCode = instCpStockCode.CodetoName(testCode)
print(testCode)
print(nameCode)

instStockChart.SetInputValue(0, testCode)
instStockChart.SetInputValue(1, ord('2'))
instStockChart.SetInputValue(2, 20160609)
instStockChart.SetInputValue(3, 20160215)
instStockChart.SetInputValue(4, 30000)
instStockChart.SetInputValue(5, (0, 1, 5, 8))
instStockChart.SetInputValue(6, ord('m'))
instStockChart.SetInputValue(7, 5)
instStockChart.SetInputValue(9, ord('1'))
instStockChart.SetInputvalue(10, ord('1'))

# BlockRequest
instStockChart.BlockRequest()

# GetHeaderValue
numData = instStockChart.GetHeaderValue(3)
numField = instStockChart.GetHeaderValue(1)
print('Date     Time End   Volume')
# GetDataValue

# Initialize lists to store data
dates = []
times = []
end = []
vols = []
#amounts = []
# Print and append into various lists
for i in range(numData):
    #if i%5 == 1:
    for j in range(numField):
        print(instStockChart.GetDataValue(j, i), end=" ")
        if j == 0:
            dates.append(instStockChart.GetDataValue(j, i))
        if j == 1:
            times.append(instStockChart.GetDataValue(j, i))
        if j == 2:
            end.append(instStockChart.GetDataValue(j, i))
        if j == 3:
            vols.append(instStockChart.GetDataValue(j, i))
        #if j == 4:
        #    amounts.append(instStockChart.GetDataValue(j, i))
    print("")


data = {'date': dates,
        'times': times,
        'end': end,
        'vols': vols}
        #'amounts': amounts}

df = pd.DataFrame(data)
print(df)
df = df.sort_values(by=['date', 'times'], ascending=True)
df = df.reset_index()
df = df.drop('index', axis=1)
print(df)
print(len(dates))
print('date size:', sys.getsizeof(dates))
print('times size:', sys.getsizeof(dates))
print('end size:', sys.getsizeof(dates))
print('vols size:', sys.getsizeof(dates))
#print('amounts size:', sys.getsizeof(dates))
#df.to_csv('삼성전자_20160215_20160429.csv')
