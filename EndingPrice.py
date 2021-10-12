import pandas as pd
import yfinance as yf
import datetime
import time
import csv
import xlrd

# Give the location of the file
loc = ("Stocks.xlsx")
 
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)
 
# Extracting number of rows
Companies=[]
Symbols=[]

for i in range(sheet.nrows):
    Companies.append(sheet.cell_value(i, 0))
    Symbols.append(sheet.cell_value(i, 1).strip())

deselected = ['Low', 'Open','High','Volume','Dividends','Stock Splits']

for symbol in Symbols:  
    try:
        msft = yf.Ticker(symbol+".NS")
        #mydata=msft.history(start="2021-6-02", end="2021-10-11", interval="1d") #it is 90 working days of stock market 
        mydata=msft.history(start="2021-7-11", end="2021-10-11", interval="1d") #it is for last 90 days of calander
        
        selectlist =[x for x in mydata.columns if x not in deselected]
        datatowrite = mydata[selectlist]
       
        #datatowrite.to_csv("AllStocksFiles/"+symbol+'.csv')  #it containing ending price for 90 working days of stock market 
        datatowrite.to_csv("allStocksEndPriceof90days/"+symbol+'.csv') #it containing ending price of all stocks for last 90 days of calendar

    except Exception:
        None
