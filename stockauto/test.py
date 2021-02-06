import win32com.client
import pandas as pd
instStockChart = win32com.client.Dispatch("CpSysDib.StockChart")
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")

testCode = instCpStockCode.NameToCode('삼성전자')
nameCode = instCpStockCode.CodetoName(testCode)
print(testCode)
print(nameCode)

instStockChart.SetInputValue(0, testCode)
instStockChart.SetInputValue(1, ord('1'))
instStockChart.SetInputValue(2, 20160217)
instStockChart.SetInputValue(3, 20160217)
#instStockChart.SetInputValue(4, 78)
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

dates = []
times = []
end = []
vols = []


for i in range(numData):
    #if i%5 == 1:
    for j in range(numField):
        print(instStockChart.GetDataValue(0, i))
        print(instStockChart.GetDataValue(1, i))
        print(instStockChart.GetDataValue(2, i))
        print(instStockChart.GetDataValue(3, i))
        dates.append(instStockChart.GetDataValue(0, i))
        times.append(instStockChart.GetDataValue(1, i))
        end.append(instStockChart.GetDataValue(2, i))
        vols.append(instStockChart.GetDataValue(3, i))
    print("")

data = {'date': dates,
        'times': times,
        'end': end,
        'vols': vols}

df = pd.DataFrame(data, columns=['Dates', 'Time', 'End', 'Vol'])
df = df.set_index('Dates')

df.to_csv('5MinTestData.csv')

