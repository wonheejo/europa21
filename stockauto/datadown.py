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
instStockChart.SetInputValue(1, ord('1'))
instStockChart.SetInputValue(2, 20190417)
instStockChart.SetInputValue(3, 20190226)
#instStockChart.SetInputValue(4, 30000)
instStockChart.SetInputValue(5, (0, 1, 2, 3, 4, 5, 8))
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

# Initialize lists to store data
dates = []
times = []
start = []
end = []
high = []
low = []
vols = []
#amounts = []
# Print and append into various lists
for i in range(numData):
    #if i%5 == 1:
    for j in range(numField):
        #print(instStockChart.GetDataValue(j, i), end=" ")
        if j == 0:
            dates.append(instStockChart.GetDataValue(j, i))
        if j == 1:
            times.append(instStockChart.GetDataValue(j, i))
        if j == 2:
            start.append(instStockChart.GetDataValue(j, i))
        if j == 3:
            high.append(instStockChart.GetDataValue(j, i))
        if j == 4:
            low.append(instStockChart.GetDataValue(j, i))
        if j == 5:
            end.append(instStockChart.GetDataValue(j, i))
        if j == 6:
            vols.append(instStockChart.GetDataValue(j, i))

    #print("")


data = {'date': dates,
        'times': times,
        'start': start,
        'high:': high,
        'low': low,
        'end': end,
        'vols': vols}
        #'amounts': amounts}

df = pd.DataFrame(data)
df = df.sort_values(by=['date', 'times'], ascending=True)
df = df.reset_index()
df = df.drop('index', axis=1)
print(df)
print('number of data: ', numData)
print(len(dates))
print('date size:', sys.getsizeof(dates))
print('times size:', sys.getsizeof(times))
print('start size:', sys.getsizeof(start))
print('high size:', sys.getsizeof(high))
print('low size:', sys.getsizeof(low))
print('end size:', sys.getsizeof(end))
print('vols size:', sys.getsizeof(vols))
#print('amounts size:', sys.getsizeof(dates))
df.to_csv('삼성전자_20190226_20190417.csv')
