
import pandas as pd
import yfinance as yf
from yahoofinancials import YahooFinancials
import datetime
import time

import csv
import requests



import xlrd
 
# Give the location of the file
loc = ("Stocks.xlsx")
 
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
 
# Extracting number of rows
print(sheet.nrows)
Companies=[];
Symbols=[];
stock_final = pd.DataFrame()

for i in range(sheet.nrows):
    Companies.append(sheet.cell_value(i,0));
    ss=sheet.cell_value(i,1).split(' ');
    Symbols.append(ss[0]);
    print(sheet.cell_value(i,0));
    print(sheet.cell_value(i,1));

print(Symbols);
deselected = ['Low', 'Open','High','Volume','Dividends','Stock Splits']
for symbol in Symbols:  
    try:
        msft = yf.Ticker(symbol+".NS")
        print(msft)
        #mydata=msft.history(start="2021-6-02", end="2021-10-11", interval="1d"); #it is 90 working days of stock market 
        mydata=msft.history(start="2021-7-11", end="2021-10-11", interval="1d"); #it is for last 90 days of calander
        
        selectlist =[x for x in mydata.columns if x not in deselected]
        datatowrite = mydata[selectlist]
        print(datatowrite)
        #datatowrite.to_csv("AllStocksFiles/"+symbol+'.csv')  #it containing ending price for 90 working days of stock market 
        datatowrite.to_csv("allStocksEndPriceof90days/"+symbol+'.csv') #it containing ending price of all stocks for last 90 days of calander
    except Exception:
        print("it goes here");
        None