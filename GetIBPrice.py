import pandas as pd
import numpy as np
import xlrd, xlutils.copy
import datetime

import time
from ib.ext.Contract import Contract
from ib.opt import ibConnection, message

#TWS query timeout limit
time_out = 10
#TWS query interval
tws_interval = 10
#Yahoo query limit
query_limit = 50

def hist_data_handler(msg):
    global all_received
    global price_matrix
    global reqID
    global length1
    if msg.open != -1:
        print stk.m_primaryExch + ':' + stk.m_symbol + ',' + msg.date + ',' + str(msg.open)    
        if length1<3 and reqID == msg.reqId:
            length1 = length1 + 1
            price_matrix.append(msg.open)
    else:
        all_received = True

#Program Start
#Connect to IB TWS
con = ibConnection()
con.register(hist_data_handler, message.historicalData) 
print '[Request TWS Permission: Press Enter to continue]'
if con.connect():
    print('[Connected to TWS]')

print('[Open "master headline 2013.xlsm"]')
book = xlrd.open_workbook('master headline 2013.xlsm')
sheet = book.sheet_by_name('Ratings Changes')

row_index = []
dates = []
tickers = []

for i in range(1,sheet.nrows):
    if sheet.cell_value(i,0) != '':
        date_value = sheet.cell_value(i,0)
        ticker = sheet.cell_value(i,2)
        dt = datetime.datetime(*xlrd.xldate_as_tuple(date_value,book.datemode))
        row_index.append(i)
        dates.append(dt.strftime("%Y%m%d"))
        tickers.append(ticker)

df = pd.DataFrame({'Date':pd.Series(dates, index=row_index),'Ticker':pd.Series(tickers, index=row_index)})
df['Ticker'] = df['Ticker'].apply(lambda x: x.replace('-',' '))
                                  
#Find Unique Tickers and Get Exchange Info from Yahoo!
unique_tickers = pd.unique(df['Ticker']).tolist()
ref_position = pd.match(df['Ticker'].tolist(),unique_tickers).tolist()
unique_tickers = [i.replace(' ','-') for i in unique_tickers]

print('[Reading Exchange Information from Yahoo! Finance]')
exchange_info = []

for i in range(len(unique_tickers)/query_limit):
    print('Reading ' + str((i+1)*query_limit))
    query_url = 'http://download.finance.yahoo.com/d/quotes.csv?s=' + '+'.join(unique_tickers[i*query_limit:(i+1)*query_limit]) + '&f=x'
    if len(exchange_info) == 0:
        exchange_info = pd.read_csv(query_url,header=None).iloc[:,0].tolist()
    else:
        exchange_info.extend(pd.read_csv(query_url,header=None).iloc[:,0].tolist())

if len(unique_tickers)%query_limit > 0:
    print('Reading ' + str(len(unique_tickers)))
    query_url = 'http://download.finance.yahoo.com/d/quotes.csv?s=' + '+'.join(unique_tickers[len(unique_tickers)/query_limit*query_limit:len(unique_tickers)]) + '&f=x'
    exchange_info.extend(pd.read_csv(query_url,header=None).iloc[:,0].tolist())
        
exchange_reference = pd.DataFrame({'Ticker':pd.Series(unique_tickers),'Exchange':pd.Series(exchange_info)})

#Rename NasdaqNM to NASDAQ
exchange_reference['Exchange'] = exchange_reference['Exchange'].replace({"NasdaqNM":"NASDAQ"})
df['Exchange'] = exchange_reference['Exchange'][ref_position].tolist()


price_matrix = []
all_received = True
stk = Contract()
    
for i in range(len(df)):
    stk.m_secType = 'STK'
    stk.m_exchange = 'SMART'
    stk.m_symbol = df['Ticker'].iloc[i].encode('ascii')
    stk.m_primaryExch = df['Exchange'].iloc[i].encode('ascii')
    print('[Request ' + stk.m_symbol + ' on ' + df['Date'].iloc[i].encode('ascii') + ']')
    con.reqHistoricalData(int(df.index[i]), stk, df['Date'].iloc[i].encode('ascii')+' 10:30:00','2700 S','15 mins','TRADES',1,1)

    all_received = False

    length1 = 0
    reqID = int(df.index[i])
    
    start_time = time.time()
    while (not all_received) and (time.time() - start_time < time_out):
        pass

    if length1 < 3:
        for i in xrange(3 - length1):
            price_matrix.extend([0])

    time.sleep(tws_interval)

    if (i+1)%query_limit == 0:
        print('[Auto-save ' + str(i+1) + ' queries to Result.csv]')
        price_array = np.array(price_matrix)
        price_array = np.resize(price_array,(len(price_matrix)/3,3))
        df1 = pd.DataFrame(price_array,index=row_index[0:(i+1)],columns=["Price1","Price2","Price3"])
        df2 = df.iloc[0:i+1].merge(df1, left_index=True, right_index=True)
        df2.to_csv('Result.csv')

        
if con.disconnect():
    print '[Disconnected from TWS]'

price_array = np.array(price_matrix)
price_array = np.resize(price_array,(len(price_matrix)/3,3))
df1 = pd.DataFrame(price_array,index=row_index,columns=["Price1","Price2","Price3"])
df = df.merge(df1, left_index=True, right_index=True)

print('[Save to "Result.csv"]')
df.to_csv('Result.csv')

print('[Done!]')
