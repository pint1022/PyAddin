from .utils import context
import os
os.add_dll_directory(r"C:\windows\system32")

from datetime import datetime, timedelta
from pytz import timezone
# import yahoo_fin.stock_info as si
import pandas_ta as ta
# import numpy as np
# import yfinance as yf

# return lastPrice
# other data:lastTradeDate, strike, lastPrice, bid, ask  change  percentChange  volume  openInterest  impliedVolatility, inTheMoney contractSize currency  
def get_option_data_by_symbol(ticker, symbol, expiration, label='lastPrice', flag=0):
    tick= yf.Ticker(ticker)
    if (flag==0): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    return options[options['contractSymbol']==symbol][label]

def get_option_data_test():
    return 100
    ticker='tsla'
    strike = 195
    expiration = '2023-02-24'
    label='lastPrice'
    flag=0    
    tick= yf.Ticker(ticker)
    if (flag==0): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    return options[options['strike']==strike][label]

def get_option_data_by_strike(ticker, strike, expiration, label='lastPrice', flag=0):
    tick= yf.Ticker(ticker)
    if (flag==0): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    return options[options['strike']==strike][label]
    # , options[options['strike']==strike]['strike']        

# def update(x:str, y:str):
#     wb = context.get_caller()
#     sheet = wb['Tickers']    
#     for row in sheet.iter_rows(min_row=1, max_col=7, values_only=True):
#         if not row[0]:
#             break
        
#         # Read the value of the corresponding cell in Column B
#         ticker = sheet.cell(row=0, column=1).value
#         strick = sheet.cell(row=0, column=2).value 
#         expiration = sheet.cell(row=0, column=3).value
#         flag = sheet.cell(row=0, column=4).value
#         res =  get_option_data_by_strike(ticker=ticker, strike=strike, expiration=expiration, flag=flag)
#         sheet.cell(row=0, column=5).value = res

# def update_test():
#     wb = context.get_caller()
#     sheet = wb['Tickers']    
#     ticker = sheet.cell(row=0, column=0).Value
#     strick = sheet.cell(row=0, column=1).Value 
#     expiration = sheet.cell(row=0, column=2).Value
#     flag = sheet.cell(row=0, column=3).Value
#     res = get_option_data_by_strike(ticker=ticker, strike=strike, expiration=expiration, flag=flag)
#     sheet.cell(row=0, colum=4).Value = res
#     return res



def monitor():
    # get the workbook calling this method, then do anything with win32com
    wb = context.get_caller()
    sheet = wb.ActiveSheet

    # get cells value
    x = sheet.Range('B1').Value
    y = sheet.Range('B2').Value

    # set cell value
    sheet.Range('B10').Value = x + y

def run_example_1(x:str, y:str):
    return int(x) + int(y)

def run_example_2():
    # get the workbook calling this method, then do anything with win32com
    wb = context.get_caller()
    sheet = wb.ActiveSheet

    # get cells value
    x = sheet.Range('A1').Value
    y = sheet.Range('A2').Value

    # set cell value
    sheet.Range('A3').Value = x + y   

