from .utils import context
import os
# os.add_dll_directory(r"C:\windows\system32")

from datetime import datetime, timedelta
from pytz import timezone
# import yahoo_fin.stock_info as si
# import pandas_ta as ta
import numpy as np
import yfinance as yf
# from openpyxl import Workbook

# return lastPrice
# other data:lastTradeDate, strike, lastPrice, bid, ask  change  percentChange  volume  openInterest  impliedVolatility, inTheMoney contractSize currency  
def get_option_data_by_symbol(ticker, symbol, expiration, label='lastPrice', flag=0):
    tick= yf.Ticker(ticker)
    if (flag==0): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    return options[options['contractSymbol']==symbol][label].tolist()[0]

def get_option_data(ticker, strike, expiration, label='lastPrice', flag='call'):
    tick= yf.Ticker(ticker)
    if (flag=='call'): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    current_price = tick.history(period='1d')['Close'][0]    
    return options[options['strike']==float(strike)][label].tolist()[0]
 
def get_option_data_by_strike(ticker, strike, expiration, label='lastPrice', flag='call'):
    tick= yf.Ticker(ticker)
    if (flag=='call'): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    current_price = tick.history(period='1d')['Close'][0]    
    return options[options['strike']==float(strike)][label].tolist()[0], current_price
    # , options[options['strike']==strike]['strike']        

def update():
    wb = context.get_caller()
    sheet = wb.Worksheets('Tickers')    
    currentRow = 2 # Start at row 2
    # sheet.Range("G" + str(currentRow)).Value= sheet.Range("A" + str(currentRow)).Value
    # sheet.Range("G" + str(currentRow)).Value= sheet.Range("C" + str(currentRow)).Value.format('YYYY-MM-DD')
    PRICE_CELL = 'F1'
    while (sheet.Range("A" + str(currentRow)).Value != ""):
        ticker = sheet.Range("A" + str(currentRow)).Value
        strike = sheet.Range("B" + str(currentRow)).Value
        expiration = sheet.Range("C" + str(currentRow)).Value
        flag = sheet.Range("D" + str(currentRow)).Value
        Label = "lastPrice"
        res, current_price =  get_option_data_by_strike(ticker=ticker, strike=strike, expiration=expiration, flag=flag)
        sheet.Range("E" + str(currentRow)).Value = res
        sheet.Range(PRICE_CELL).Value= current_price
        currentRow = currentRow + 1

def update_1():
    wb = context.get_caller()
    sheet = wb['Tickers']    
    current_price = tick.history(period='1d')['Close'][0]    
    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True):
        if not row[0]:
            break
        
        # Read the value of the corresponding cell in Column B
        ticker = sheet.cell(row=row, column=1).value
        strick = sheet.cell(row=row, column=2).value 
        expiration = sheet.cell(row=row, column=3).value
        flag = sheet.cell(row=row, column=4).value
        res =  get_option_data_by_strike(ticker=ticker, strike=strike, expiration=expiration, flag=flag)
        sheet.cell(row=row, column=5).value = ticker


