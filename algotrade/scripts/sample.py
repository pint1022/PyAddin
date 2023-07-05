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
import sys
import pandas as pd

from datetime import datetime, timedelta
import matplotlib.pyplot as plt
plt.style.use('seaborn')
import tensorflow as tf
# from tensorflow.keras.models import *
# from tensorflow.keras.layers import *

from .myUtils import *
from .myInds import *
from .Transformer_model import *

core = ["TSLA", "GOOG", "ENPH", "NVDA","MRNA","AAPL"]
energy = ["GUSH", "CF", "EQT","OXY","SEDG"]
semi = ["QCOM","LRCX","ASML","AMD","SOXL","TSM"]
index = ["SPY","QQQ","SQQQ","TQQQ"]
fin = ["BLK","BX","TLT"]
hot = ["SHOP","FSLY","U","ZM","nflx","pypl"]
btc = ["COIN","BTC-USD","MSTR","ETH-USD"]
software = ["ADBE","MSFT"]
base = ["COST","HD","KO","MCD","PFE","UNH"]
bio = ["XBI","amgn"]
cloud = ["OKTA","DDOG","DASH","NOW"]
random = ["NKX","DOG","SE","IBM"]
tickers = core + index 

# return lastPrice
# other data:lastTradeDate, strike, lastPrice, bid, ask  change  percentChange  volume  openInterest  impliedVolatility, inTheMoney contractSize currency  
def get_option_data_by_symbol(ticker, symbol, expiration, label='lastPrice', flag='call'):
    tick= yf.Ticker(ticker)
    if (flag==0): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    return options[options['contractSymbol']==symbol][label].tolist()[0]

def get_option(ticker, strike, expiration,  flag='call'):
    tick= yf.Ticker(ticker)    
    if (flag=='call'): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
    else:
        options = tick.option_chain(expiration).puts
    return options[options['strike']==strike]['lastPrice'].tolist()[0],\
           options[options['strike']==strike]['change'].tolist()[0],\
           options[options['strike']==strike]['impliedVolatility'].tolist()[0],\
           options[options['strike']==strike]['volume'].tolist()[0]

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

def get_data(ticker, strike, expiration, label='lastPrice', flag='call'):
    tick= yf.Ticker(ticker)
    if (flag=='call'): # calls, date 'yyyy-m-d'
        options = tick.option_chain(expiration).calls
        current_price = options[options['strike']==float(strike)][label].tolist()[0]
    elif (flag=='put'):
        options = tick.option_chain(expiration).puts
        current_price = options[options['strike']==float(strike)][label].tolist()[0]
    else:
        current_price = tick.history(period='1d')['Close'][0]    
    return  current_price

def get_price(ticker):
    tick= yf.Ticker(ticker)
    current_price = tick.history(period='1d')['Close'][0]    
    return  current_price

def update_holding(sheet): 
    currentRow = 2 # Start at row 2
    # sheet.Range("G" + str(currentRow)).Value= sheet.Range("A" + str(currentRow)).Value
    # sheet.Range("G" + str(currentRow)).Value= sheet.Range("C" + str(currentRow)).Value.format('YYYY-MM-DD')
    while (sheet.Range("A" + str(currentRow)).Value != ""):
        ticker = sheet.Range("A" + str(currentRow)).Value
        strike = sheet.Range("B" + str(currentRow)).Value
        expiration = sheet.Range("C" + str(currentRow)).Value
        flag = sheet.Range("D" + str(currentRow)).Value
        Label = "lastPrice"
        res =  get_data(ticker=ticker, strike=strike, expiration=expiration, flag=flag)
        sheet.Range("E" + str(currentRow)).Value = res
        currentRow = currentRow + 1

# model_dir='../'
model_dir=os.path.abspath(os.getcwd())

def predict_data(ticker, df, mean_k, seq_len):
    df[['open', 'high', 'low', 'close', 'volume']] = df[['open', 'high', 'low', 'close', 'volume']].rolling(int(mean_k)).mean() 
    df_mean_price = df
    df.name = ticker
    # times = sorted(df.index.values)

    # last_10pct = sorted(df.index.values)[-int(0.1*len(times))] # Last 10% of series
    # last_20pct = sorted(df.index.values)[-int(0.2*len(times))] # Last 20% of series
    '''Calculate percentage change'''

    df['open'] = df_mean_price['open'].pct_change() # Create arithmetic returns column
    df['high'] = df_mean_price['high'].pct_change() # Create arithmetic returns column
    df['low'] = df_mean_price['low'].pct_change() # Create arithmetic returns column
    df['close'] = df_mean_price['close'].pct_change() # Create arithmetic returns column
    df['volume'] = df_mean_price['volume'].pct_change()

    # min_return = min(df[(df.index < last_20pct)][['open', 'high', 'low', 'close']].min(axis=0))
    # max_return = max(df[(df.index < last_20pct)][['open', 'high', 'low', 'close']].max(axis=0))
    min_return = min(df[['open', 'high', 'low', 'close']].min(axis=0))
    max_return = max(df[['open', 'high', 'low', 'close']].max(axis=0))
    df['open'] = (df['open'] - min_return) / (max_return - min_return)
    df['high'] = (df['high'] - min_return) / (max_return - min_return)
    df['low'] = (df['low'] - min_return) / (max_return - min_return)
    df['close'] = (df['close'] - min_return) / (max_return - min_return)

    ###############################################################################
    '''Normalize volume column'''

    min_volume = df['volume'].min(axis=0)
    max_volume = df['volume'].max(axis=0)
    # min_volume = df[(df.index < last_20pct)]['volume'].min(axis=0)
    # max_volume = df[(df.index < last_20pct)]['volume'].max(axis=0)
    # Min-max normalize volume columns (0-1 range)
    df['volume'] = (df['volume'] - min_volume) / (max_volume - min_volume)
    df  = df[['open', 'high', 'low', 'close', 'volume']]

    seq_len = int(seq_len)
    df_target = df[-seq_len*2:]
    X_target = df_target.values
    X_pred = []
    for i in range(seq_len, len(X_target)):
        X_pred.append(X_target[i-seq_len:i])
    X_pred = np.array(X_pred)

    return X_pred, max_return, min_return


def get_predictions(ticker, mean_k=5, seq_len=32, current_date=datetime.now()):
    end_date = current_date.strftime('%Y-%m-%d')
    start_date = '1970-01-01'
    df = yf.download(ticker, 
                        start = start_date, 
                        end = end_date, 
                        interval='1d').fillna(0)
    # print(df.columns)
    # current_price = df[-1:]['Close']    
    df.columns = df.columns.str.lower()

    X_pred, max_return, min_return = predict_data(ticker=ticker, df=df, mean_k=mean_k, seq_len=seq_len)

    # chkpnt_file = f'{model_dir}/{ticker}_seq{int(seq_len)}_m{int(mean_k)}_Transformer+TimeEmbedding.hdf5'        
    chkpnt_file = os.path.join(model_dir, f'models\{ticker}_seq{int(seq_len)}_m{int(mean_k)}_Transformer+TimeEmbedding.hdf5')        
    model = tf.keras.models.load_model(chkpnt_file,
                                       custom_objects={'Time2Vector': Time2Vector, 
                                                       'SingleAttention': SingleAttention,
                                                       'MultiAttention': MultiAttention,
                                                       'TransformerEncoder': TransformerEncoder})
    _pred = model.predict(X_pred)  
    return  _pred[-1:]*(max_return -min_return) + min_return

def update_predictions(sheet): 
    currentRow = 2 # Start at row 2
    label = 'close'
    flagCol='D'
    runtimeCol = 'E'
    dateCol = 'F'
    preCLoseCol = 'G'
    predictCLoseCol = 'H'
    actCLose = 'I'
    delta = 'J'
    currPred = 'K'
    prevPred = 'L'
    oneDayClose = 'M'
    indCol = 'O'

    check_date = datetime.now()
    days = 1
    # check_date = sheet.Range(dateCol + str(currentRow)).Value
    # if check_date is None:
    #     check_date = datetime.now().strftime('%Y-%m-%d')
    # else:
    #     check_date = check_date + timedelta(days=days)
    #     if (check_date.weekday() >= 5):
    #         return

    predDate = check_date + timedelta(days=days)
    if (predDate.weekday() == 5):
        predDate = check_date + timedelta(days=2)
    elif (predDate.weekday() == 6):
        predDate = check_date + timedelta(days=3)



    currentDate = check_date.strftime('%Y-%m-%d')
    sheet.Range(actCLose + '1').Value = f'{currentDate} close'
    sheet.Range(oneDayClose + '1').Value = sheet.Range(preCLoseCol + '1').Value
    sheet.Range(prevPred + '1').Value = sheet.Range(currPred + '1').Value 
    sheet.Range(preCLoseCol + '1').Value = f'{currentDate} close'
    sheet.Range(currPred + '1').Value = f'{currentDate} predict'
    sheet.Range(delta + '1').Value = f'{currentDate} predict-close'
    sheet.Range(indCol + '1').Value = 'Indicator'

    # need check if it is not weekend or holiday
    # while preDate.weekday() >= 5:
    #     days += 1
    #     predDate = datetime.today() + timedelta(days=days)

    predDate = predDate.strftime('%Y-%m-%d')
    sheet.Range(predictCLoseCol + '1').Value = f'{predDate} predict'

    while (sheet.Range("A" + str(currentRow)).Value != ""):
        flag = sheet.Range(flagCol + str(currentRow)).Value

        # if (flag=='act'): #read the real close price, if the date is not before today, return 
        #     if check_date.strftime('%Y-%m-%d') > datetime.now().strftime('%Y-%m-%d'):
        #         return
        currentDate = check_date
        ticker = sheet.Range("A" + str(currentRow)).Value
        ticker.upper()
        act_price= get_price(ticker)
        if (flag=='act'):
            sheet.Range(actCLose + str(currentRow)).Value = act_price
        else:
            seq_len = sheet.Range("B" + str(currentRow)).Value
            mean_k = sheet.Range("C" + str(currentRow)).Value
            ts = time.time()
            pred_price =  get_predictions(ticker=ticker, mean_k=mean_k, seq_len=seq_len, current_date=currentDate)
            sheet.Range(runtimeCol + str(currentRow)).Value = (time.time() - ts) * 1000
            sheet.Range(prevPred + str(currentRow)).Value = sheet.Range(currPred + str(currentRow)).Value
            sheet.Range(currPred + str(currentRow)).Value = sheet.Range(predictCLoseCol + str(currentRow)).Value
            sheet.Range(oneDayClose + str(currentRow)).Value = sheet.Range(preCLoseCol + str(currentRow)).Value
            sheet.Range(preCLoseCol + str(currentRow)).Value = act_price
            sheet.Range(predictCLoseCol + str(currentRow)).Value = act_price * (pred_price + 1)
            if pred_price > 0.0:
                sheet.Range(indCol + str(currentRow)).Value = sheet.Range(indCol + str(currentRow)).Value + 1
            else:
                sheet.Range(indCol + str(currentRow)).Value = sheet.Range(indCol + str(currentRow)).Value + 1
        sheet.Range(dateCol + str(currentRow)).Value = currentDate.strftime('%Y-%m-%d')

        currentRow = currentRow + 1

# def get_predictions_2(ticker, mean_k=5, seq_len=32, current_date=datetime.now(), delta=0):
#     end_date = current_date.strftime('%Y-%m-%d')
#     start_date = '1970-01-01'
#     df = yf.download(ticker, 
#                         start = start_date, 
#                         end = end_date, 
#                         interval='1d').fillna(0)
#     # print(df.columns)
#     # current_price = df[-1:]['Close']    
#     df.columns = df.columns.str.lower()

#     X_pred, max_return, min_return = predict_data(ticker=ticker, df=df, mean_k=mean_k, seq_len=seq_len)

#     # chkpnt_file = f'{model_dir}/{ticker}_seq{int(seq_len)}_m{int(mean_k)}_Transformer+TimeEmbedding.hdf5'        
#     # chkpnt_file = os.path.join(model_dir, f'models\{ticker}_seq{int(seq_len)}_m{int(mean_k)}_Transformer+TimeEmbedding.hdf5')        
#     # model = tf.keras.models.load_model(chkpnt_file,
#     #                                    custom_objects={'Time2Vector': Time2Vector, 
#     #                                                    'SingleAttention': SingleAttention,
#     #                                                    'MultiAttention': MultiAttention,
#     #                                                    'TransformerEncoder': TransformerEncoder})
#     # _pred = model.predict(X_pred)  
#     return  delta * (max_return -min_return) + min_return

def update_predictions_2(sheet): 
    colnames=["index", "ticker", "seq_len", "mean_k","flag","runtime","last date","close","delta_c", "delta_o","delta_h","delta_l","delta_v","opt"]
    pred_file = os.path.join(os.getcwd(), 'models\predictions\predictions.csv') 

    data = np.array(pd.read_csv(pred_file, header=None, names=colnames))
    predcol = 7
    count = len(data)
    print(count)

    currentRow = 2 # Start at row 2
    label = 'close'
    flagCol='D'
    runtimeCol = 'E'
    dateCol = 'F'
    preCLoseCol = 'G'
    predictCLoseCol = 'H'
    actCLose = 'I'
    delta = 'J'
    currPred = 'K'
    prevPred = 'L'
    oneDayClose = 'M'
    indCol = 'O'

    check_date = datetime.now()
    days = 1
    # check_date = sheet.Range(dateCol + str(currentRow)).Value
    # if check_date is None:
    #     check_date = datetime.now().strftime('%Y-%m-%d')
    # else:
    #     check_date = check_date + timedelta(days=days)
    #     if (check_date.weekday() >= 5):
    #         return

    predDate = check_date + timedelta(days=days)
    if (predDate.weekday() == 5):
        predDate = check_date + timedelta(days=2)
    elif (predDate.weekday() == 6):
        predDate = check_date + timedelta(days=3)



    currentDate = check_date.strftime('%Y-%m-%d')
    sheet.Range(actCLose + '1').Value = f'{currentDate} close'
    sheet.Range(oneDayClose + '1').Value = sheet.Range(preCLoseCol + '1').Value
    sheet.Range(prevPred + '1').Value = sheet.Range(currPred + '1').Value 
    sheet.Range(preCLoseCol + '1').Value = f'{currentDate} close'
    sheet.Range(currPred + '1').Value = f'{currentDate} predict'
    sheet.Range(delta + '1').Value = f'{currentDate} predict-close'

    # need check if it is not weekend or holiday
    # while preDate.weekday() >= 5:
    #     days += 1
    #     predDate = datetime.today() + timedelta(days=days)

    predDate = predDate.strftime('%Y-%m-%d')
    sheet.Range(predictCLoseCol + '1').Value = f'{predDate} predict'
    for row in range(len(data)):
        if data[row][5] == 0:
            continue
        sheet.Range('A' + str(currentRow)).Value  = data[row][0]
        sheet.Range('B' + str(currentRow)).Value  = data[row][1]
        sheet.Range('C' + str(currentRow)).Value  = data[row][2]
        sheet.Range('D' + str(currentRow)).Value  = data[row][3]

        # if (flag=='act'): #read the real close price, if the date is not before today, return 
        #     if check_date.strftime('%Y-%m-%d') > datetime.now().strftime('%Y-%m-%d'):
        #         return
        currentDate = check_date
        ticker = data[row][0]
        ticker.upper()
        act_price= get_price(ticker)

        seq_len = data[row][1]
        mean_k = data[row][2]
        # pred_price =  get_predictions_2(ticker=ticker, mean_k=mean_k, seq_len=seq_len, current_date=currentDate, delta=data[row][predcol])
        pred_price =  data[row][predcol]

        sheet.Range(runtimeCol + str(currentRow)).Value = data[row][4]
        sheet.Range(prevPred + str(currentRow)).Value = sheet.Range(currPred + str(currentRow)).Value
        sheet.Range(currPred + str(currentRow)).Value = sheet.Range(predictCLoseCol + str(currentRow)).Value
        sheet.Range(oneDayClose + str(currentRow)).Value = sheet.Range(preCLoseCol + str(currentRow)).Value
        sheet.Range(preCLoseCol + str(currentRow)).Value = act_price
        if pred_price > 0.0:
            sheet.Range(indCol + str(currentRow)).Value = sheet.Range(indCol + str(currentRow)).Value + 1
        else:
            sheet.Range(indCol + str(currentRow)).Value = sheet.Range(indCol + str(currentRow)).Value - 1        
        sheet.Range(predictCLoseCol + str(currentRow)).Value = act_price * (pred_price + 1)
        sheet.Range(dateCol + str(currentRow)).Value = currentDate.strftime('%Y-%m-%d')

        currentRow = currentRow + 1

def predict():
    wb = context.get_caller()
    sheet = wb.Worksheets('predictions') 
    update_predictions(sheet)  

def update_actual(sheet): 
    currentRow = 2 # Start at row 2
    label = 'close'
    flagCol='D'
    runtimeCol = 'E'
    dateCol = 'F'
    preCLoseCol = 'G'
    predictCLoseCol = 'H'
    actCLose = 'I'
    delta = 'J'
    currPred = 'K'
    prevPred = 'L'
    days = 1
    currentDate = datetime.now().strftime('%Y-%m-%d')
    sheet.Range(actCLose + '1').Value = f'{currentDate} current'
    sheet.Range(delta + '1').Value = f'{currentDate} RT delta'
    current = datetime.today() - timedelta(days=days)

    while (sheet.Range("A" + str(currentRow)).Value != ""):
        flag = sheet.Range(flagCol + str(currentRow)).Value
        check_date = sheet.Range(dateCol + str(currentRow)).Value
        # if check_date is None or check_date.strftime('%Y-%m-%d') < current.strftime('%Y-%m-%d'):
        #     break
        # if check_date.strftime('%Y-%m-%d') < current.strftime('%Y-%m-%d'):
        #     break
        ticker = sheet.Range("A" + str(currentRow)).Value
        ticker.upper()
        act_price= get_price(ticker)
        sheet.Range(actCLose + str(currentRow)).Value = act_price
        sheet.Range(delta + str(currentRow)).Value = act_price - sheet.Range(predictCLoseCol + str(currentRow)).Value
        currentRow = currentRow + 1

def actual():
    wb = context.get_caller()
    # sheet = wb.Worksheets('predictions') 
    # update_actual(sheet)  
    sheet = wb.Worksheets('predictions2') 
    update_predictions_2(sheet)  

def research(): 
    row = 2 # Start at row 2
    wb = context.get_caller()
    sheet = wb.Worksheets['candidates']    
    NOT_AVAILABLE = 'F10'
    # sheet.Cells(10,1).Value  = "NOT AVAILABLE"
    tickerrow = 1
    tickercol = 1
    count = 4
    row = 5
    today = datetime.now()
    expiration = (today + timedelta( (4-today.weekday()) % 7 )).strftime('%Y-%m-%d')
    Label = "lastPrice"
    # if (sheet.Cells(tickerrow,tickerrowcol).Value == "") or (sheet.Cells(tickerrow+1,tickerrowcol).Value == ""):
    #     sheet.Cells(10,1).Value  = "NO data"
    #     return
    while ((sheet.Cells(tickerrow,tickercol).Value != "") and (sheet.Cells(tickerrow+1,tickercol).Value != "")):
        price = int(sheet.Cells(tickerrow+1,tickercol).Value)
        ticker = sheet.Cells(tickerrow,tickercol).Value
        if price < 10.0:
            sheet.Cells(10,1).Value  = "NO supported lower than $10 price"
            return
        elif price < 100.0:
            delta = 1
        # elif price < 300.0:
        #     delta = 5
        else:
            delta = 5
        start = round(price/10)*10  - count/2 * delta
        stop = start + (count + 1) * delta
        col = 1

        for i in  range(int(start), int(stop), delta):
            sheet.Cells(row,1).Value = ticker
            sheet.Cells(row,2).Value = i
            sheet.Cells(row,3).Value = expiration
            # sheet.Cells(row,4).Value = flag
            flag = "call"
            try:
                price, change, iv, vol =  get_option(ticker=ticker, strike=i, expiration=expiration, flag=flag)
            except:
                continue
            sheet.Cells(row,5).Value = price
            sheet.Cells(row,6).Value = change
            sheet.Cells(row,7).Value = iv
            sheet.Cells(row,8).Value = vol
            flag = "put"
            try:
                price, change, iv, vol =  get_option(ticker=ticker, strike=i, expiration=expiration, flag=flag)
            except:
                continue
            sheet.Cells(row,10).Value = price
            sheet.Cells(row,11).Value = change
            sheet.Cells(row,12).Value = iv
            sheet.Cells(row,13).Value = vol
            row = row + 1
        row = row + 1
        tickercol = tickercol + 1
        

def update_candidates():
    wb = context.get_caller()
    sheet = wb.Worksheets['candidates']    
    row = 1
    col = 1
    # sheet.Cells(row+1,col).Value = sheet.Cells(row,col).Value
    while (sheet.Cells(row,col).Value != ""):
        ticker = sheet.Cells(row,col).Value
        res =  get_price(ticker=ticker)
        sheet.Cells(row+1,col).Value = res
        col = col + 1


def update():
    wb = context.get_caller()
    sheet = wb.Worksheets('Tickers')  
    update_holding(sheet)     

   
