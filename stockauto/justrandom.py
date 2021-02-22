# This is for sending an alarm/message to slacker bot

from slacker import Slacker
import os, sys, ctypes
import win32com.client
import pandas as pd
from datetime import datetime
import time, calendar

slack = Slacker('xoxb-1709162090453-1712267545922-zBUfyirsvPaOjIXJVqdWle4R')

# Send a message to #general channel

def dbgout(message):
    """인자로 받은 문자열을 파이썬 셸과 슬랙으로 동시에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message)
    strbuf = datetime.now().strftime('[%m/%d %H:%M:%S] ') + message
    slack.chat.post_message('#stocks', strbuf)

def printlog(message, *args):
    """인자로 받은 문자열을 파이썬 셸에 출력한다."""
    print(datetime.now().strftime('[%m/%d %H:%M:%S]'), message, *args)

cpCodeMgr = win32com.client.Dispatch('CpUtil.CpStockCode')
cpStatus = win32com.client.Dispatch('CpUtil.CpCybos')
cpTradeUtil = win32com.client.Dispatch('CpTrade.CpTdUtil')
cpStock = win32com.client.Dispatch('DsCbo1.StockMst')
cpOhlc = win32com.client.Dispatch('CpSysDib.StockChart')
cpBalance = win32com.client.Dispatch('CpTrade.CpTd6033')
cpCash = win32com.client.Dispatch('CpTrade.CpTdNew5331A')
cpOrder = win32com.client.Dispatch('CpTrade.CpTd0311')

cpTradeUtil.TradeInit()
acc = cpTradeUtil.AccountNumber[0]  # 계좌번호
accFlag = cpTradeUtil.GoodsList(acc, 1)  # -1:전체, 1:주식, 2:선물/옵션
cpBalance.SetInputValue(0, acc)  # 계좌번호
cpBalance.SetInputValue(1, accFlag[0])  # 상품구분 - 주식 상품 중 첫번째
cpBalance.SetInputValue(2, 50)  # 요청 건수(최대 50)
cpBalance.BlockRequest()

#dbgout('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
print('계좌명: ' + str(cpBalance.GetHeaderValue(0)))
