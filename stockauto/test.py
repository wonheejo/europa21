import win32com.client
import pandas as pd
import numpy as np
cpOhlc = win32com.client.Dispatch("CpSysDib.StockChart")
instCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')

testCode = instCpStockCode.NameToCode('삼성전자')
nameCode = instCpStockCode.CodetoName(testCode)
print(testCode)
print(nameCode)


# A005930 = Samsung Electronics
# A068270 = Celltrion
# A035720 = Kakao
# A011200 = HMM
# A009830 = Hanhwa Solution
# A042040 = KPM tech
symbol_list = ['A005930', 'A068270', 'A035720', 'A011200', 'A009830', 'A042040']

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

def get_current_price(code):
    """인자로 받은 종목의 현재가, 매수호가, 매도호가를 반환한다."""
    cpStock.SetInputValue(0, code)  # 종목코드에 대한 가격 정보
    cpStock.BlockRequest()
    item = {}
    item['cur_price'] = cpStock.GetHeaderValue(11)  # 현재가
    item['ask'] = cpStock.GetHeaderValue(16)  # 매수호가
    item['bid'] = cpStock.GetHeaderValue(17)  # 매도호가
    return item['cur_price'], item['ask'], item['bid']


lower = []
upper = []
cur_price = []
ask_price = []
bid_price = []

for i in range(len(symbol_list)):
    low, up = bollinger(symbol_list[i])
    cur, ask, bid = get_current_price(symbol_list[i])
    lower.append(low)
    upper.append(up)
    cur_price.append(cur)
    ask_price.append(ask)
    bid_price.append(bid)

for i in range(len(symbol_list)):
    nameCode = instCpStockCode.CodetoName(symbol_list[i])
    print(symbol_list[i], nameCode)
    print(lower[i], upper[i])
    print(cur_price[i])

