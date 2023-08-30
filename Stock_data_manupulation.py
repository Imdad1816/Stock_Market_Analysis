import pymysql
import sqlalchemy
import pandas as pd
from datetime import datetime
import threading

#Creating connection string
engine = sqlalchemy.create_engine('mysql+pymysql://Imdad:root@127.0.0.1:3306/Market_data')

def getOptionsData():
    #opening workbook and importing dataframe and creating dictionary
    df = pd.read_excel('D:\Data_Analytics_Project\Stock_market_Analysis\Stock_Market_Live_Data.xlsm')
    dict_list = df.to_dict('records')

    #formating datatypes
    df['CURRENT PRICE'] = pd.to_numeric(df['CURRENT PRICE'])
    df['52 WEEK HIGH'] = pd.to_numeric(df['52 WEEK HIGH'])
    df['52 WEEK LOW'] = pd.to_numeric(df['52 WEEK LOW'])
    df['MARKET CAP (RS CR.)'] = pd.to_numeric(df['MARKET CAP (RS CR.)'])
    df['TIME'] = pd.to_datetime(df['TIME'])

    #exporting to mysqldatabase
    df.to_sql(name='stock_data_table',con=engine,index=False,if_exists='replace')
    print('Entry done for',datetime.now())

#updating_database_after_time_interval
def startTimer():
    interval = 120 #seconds
    threading.Timer(interval, startTimer).start()
    getOptionsData()

if __name__ == "__main__":
    startTimer()
