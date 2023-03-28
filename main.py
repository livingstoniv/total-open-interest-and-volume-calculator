from socket import create_connection
from yahoo_fin import options as op
import pandas as pd
from datetime import datetime, date
import numpy as np
from yahoo_fin import stock_info as si
from xlsxwriter.utility import xl_rowcol_to_cell
import openpyxl as px
import sqlalchemy
import itertools
import sqlite3 as sql



#Empty Lists For Data



#Date Lists
expirationDate = []
daysToExpiration = []

#Volume Lists
totalLongVolumeByExpiration = []
totalShortVolumeByExpiration = []

#Open Int Lists
totalLongOpenInterestByExpiration = []
totalShortOpenInterestByExpiration = []

#Open Interest and Volume Percentage Lists
longOpenIntAsPerc = []
shortOpenIntAsPerc = []
longVolAsPerc = []
shortVolAsPerc = []
#OTM and ITM Both Long and Short Lists
OTMLongList = []
ITMLongList = []
OTMShortList = []
ITMShortList = []

#EDIT 2023/03/26 3pm
OTMLongVolList = []
ITMLongVolList = []
OTMShortVolList = []
ITMShortVolList = []

#OTM and ITM Percentage Both Long and Short Lists
OTMLongPercList = []
ITMLongPercList = []
OTMShortPercList = []
ITMShortPercList = []

#EDIT 2023/03/26 3pm
OTMLongVolPercList = []
ITMLongVolPercList = []
OTMShortVolPercList = []
ITMShortVolPercList = []

#implied volatility lists
AvgIVLongList = []
AvgIVOTMLongList = []
AvgIVITMLongList = []
AvgIVShortList = []
AvgIVOTMShortList = []
AvgIVITMShortList = []



#Input For Finding Stock Info, Run The Program Then Enter The Ticker Into The Terminal
ticker = 'SPY'


#Expiration Date Variables
today = date.today()
nowFile = datetime.now()
fileDate = nowFile.strftime("%Y-%m-%d-%H-%M")
todaysDate = datetime(today.year, today.month, today.day)
expDates = op.get_expiration_dates(ticker)

running = True
while running == True:

    #Loop For Fiding Expiration Dates, Iterating Data Through Each One and Combining Data Into a Central Structure
    for i in range(0, len(expDates)):
        expDate = datetime.strptime(expDates[i], '%B %d, %Y')
        DTE = (expDate - todaysDate).days
        nowie = datetime.now()
        noway = nowie.strftime("%Y-%m-%d %H:%M:%S")

        chain = op.get_options_chain(ticker, date= expDates[i])
        livePrice = si.get_live_price(ticker)
        optionTicker = ticker




        #Creates Numeric Data Structure For String Values 
        chain['calls']['Volume'] = pd.to_numeric(chain['calls']['Volume'], errors = 'coerce')
        chain['calls']['Open Interest'] = pd.to_numeric(chain['calls']['Open Interest'], errors = 'coerce')
        chain['puts']['Volume'] = pd.to_numeric(chain['puts']['Volume'], errors = 'coerce')
        chain['puts']['Open Interest'] = pd.to_numeric(chain['puts']['Open Interest'], errors = 'coerce')

        #Creates/Splices List Values For Initial DataFrame Used For Calculations
        LongOpenInt = chain['calls']['Open Interest']
        ShortOpenInt = chain['puts']['Open Interest']
        LongStrike = chain['calls']['Strike']
        ShortStrike = chain['puts']['Strike']
        LongIV = chain['calls']['Implied Volatility']
        ShortIV = chain['puts']['Implied Volatility']

        LongLastPrice = chain['calls']['Last Price']
        LongBid = chain['calls']['Bid']
        LongAsk = chain['calls']['Ask']
        LongChange = chain['calls']['Change']
        LongPercChange = chain['calls']['% Change']
    
        ShortLastPrice = chain['puts']['Last Price']
        ShortBid = chain['puts']['Bid']
        ShortAsk = chain['puts']['Ask']
        ShortChange = chain['puts']['Change']
        ShortPercChange = chain['puts']['% Change']

        #EDIT 2023/03/26 3pm !!!!!
        LongVolume = chain['calls']['Volume']
        ShortVolume = chain['puts']['Volume']


        #Iterates Through Implied Volatility Lists Long and Short and Strips The Perentage Sign From Implied Volatility List Values So It Can Be Converted Into A Float
        LongIVStrip = [i.strip('%') for i in LongIV]
        ShortIVStrip = [c.strip('%') for c in ShortIV]
        
        #Iterates Through Newly Stripped Implied Volatility Lists To Convert Strings Into Float Numbers For Average Equations
        floatLongIV = [float(g) for g in LongIVStrip]
        floatShortIV = [float(w) for w in ShortIVStrip]

        
        #Finds The Sum Of Long Volume and Open Interest
        totalLongVol = chain['calls']['Volume'].sum()
        totalLongInt = chain['calls']['Open Interest'].sum()
        
        #Finds The Sum Of Short Volume and Open Interest
        totalShortVol = chain['puts']['Volume'].sum()
        totalShortInt = chain['puts']['Open Interest'].sum()


        #Equates the Total Long Vs Short Volume and Open Interest To Find Percentage Of Short vs Long Allocation
        totalVol = totalLongVol + totalShortVol 
        totalOpenInt = totalLongInt + totalShortInt

        #Finds Percentage Of Long Vs Short Volume and Open Interest
        longVolPercent = 100 * totalLongVol / totalVol
        shortVolPercent = 100 * totalShortVol / totalVol
        longOpenIntPecent = 100 * totalLongInt / totalOpenInt
        shortOpenIntPercent = 100 * totalShortInt / totalOpenInt
        
        #Rounds The Messy Percentage Value To A Two Decimal Number For Easier Data Translation
        roundLongvolPerc = round(longVolPercent, 2)
        roundShortvolPerc = round(shortVolPercent, 2)
        roundLongOpenInt = round(longOpenIntPecent, 2)
        roundShortOpenInt = round(shortOpenIntPercent, 2)
        
        #Finds Max Value Of Long Vs Short Open Interest (Currently In Development)


        #Creates A Seperate DataFrame For Further Equations
        df = pd.DataFrame(data = list(zip(LongOpenInt, LongVolume, LongStrike, floatLongIV, ShortOpenInt, ShortVolume, ShortStrike, floatShortIV)), columns= ['Long Open Int', 'Long Volume', 'Long Strike', 'Long IV', 'Short Open Int', 'Short Volume', 'Short Strike', 'Short IV'])
        dfMainOptionChain = pd.DataFrame(data = list(zip(LongStrike, LongLastPrice, LongBid, LongAsk, LongChange, LongPercChange, LongVolume, LongOpenInt, LongIV, ShortStrike, ShortLastPrice, ShortBid, ShortAsk, ShortChange, ShortPercChange, ShortVolume, ShortOpenInt, ShortIV)), columns= ['Call Strike', "Call Last Price", 'Call Bid', 'Call Ask', 'Call Change', 'Call % Change', 'Call Volume', 'Call Open Interest', 'Call Implied Volatility', 'Put Strike', 'Put Last Price', 'Put Bid', 'Put Ask', 'Put Change', 'Put % Change', 'Put Volume', 'Put Open Interest', 'Put Implied Volatility'])
        
    
        #Finds The Average Long Vs Short IV
        AvgLongIV = sum(floatLongIV) / len(floatLongIV)
        AvgShortIV = sum(floatShortIV) / len(floatShortIV)

        
        #Rounds The Values To A Two Decimal Number For Easier Data Translation
        roundAvgLongIV = round(AvgLongIV, 2)
        roundAvgShortIV = round(AvgShortIV, 2)
    
        # maxLongStrike = df[df['Long Open Int'] == max(LongOpenInt)]['Long Strike']
        # maxLongStrikeInt = float(maxLongStrike)
        # print(maxLongStrikeInt)

        # maxShortStrike = df[df['Short Open Int'] == max(ShortOpenInt)]['Short Strike']
        # maxShortStrikeInt = float(maxShortStrike)
        # print(maxShortStrikeInt)

        

        

        #Finds all of the OTM Long Open Interest 
        OTMLong = df[df['Long Strike'] < livePrice]['Long Open Int'].sum()

        
        #Finds all of the ITM Long Open Interest
        ITMLong = df[df['Long Strike'] > livePrice]['Long Open Int'].sum()
        

        #EDIT 2023/03/26 3pm !!!!!!
        OTMLongVol = df[df['Long Strike'] < livePrice]['Long Volume'].sum()
        ITMLongVol = df[df['Long Strike'] > livePrice]['Long Volume'].sum()

        
        #Finds all of the OTM Long Implied Volatility
        OTMLongIV = df[df['Long Strike'] < livePrice]['Long IV']
        
        #Finds all of the ITM Long Implied Volatility
        ITMLongIV = df[df['Long Strike'] > livePrice]['Long IV']


        
        #Finds all of the OTM Short
        OTMShort = df[df['Short Strike'] > livePrice]['Short Open Int'].sum()
        
        
        #Finds all of the ITM Short
        ITMShort = df[df['Short Strike'] < livePrice]['Short Open Int'].sum()
        
        #EDIT 2023/03/26 3pm !!!!!!
        OTMShortVol = df[df['Short Strike'] > livePrice]['Short Volume'].sum()
        ITMShortVol = df[df['Short Strike'] < livePrice]['Short Volume'].sum()


        #Finds all of the OTM Short Implied Volatility
        OTMShortIV = df[df['Short Strike'] > livePrice]['Short IV']
        
        #Finds all of the ITM Short Implied Volatility
        ITMShortIV = df[df['Short Strike'] < livePrice]['Short IV']
        


        #Finds all of the OTM Long Implied Volatility
        AvgOTMLongIV = sum(OTMLongIV) / len(OTMLongIV)
        #Rounds the Value to a Two Decimal Number For Easier Data Translation
        roundAvgOTMLongIV = round(AvgOTMLongIV, 2)


        
        #Finds the Average Long ITM Implied Volatility
        AvgITMLongIV = sum(ITMLongIV) / len(ITMLongIV)
        #Rounds the Value to a Two Decimal Number for Easier Data Translation
        roundAvgITMLongIV = round(AvgITMLongIV, 2)
        #Finds the Average Short OTM Implied Volatility
        AvgOTMShortIV = sum(OTMShortIV) / len(OTMShortIV)
        #Rounds the Value to a Two Decimal Number For Easier Data Translation
        roundAvgOTMShortIV = round(AvgOTMShortIV, 2)
        #Finds the Average Short ITM Implied Volatility
        AvgITMShortIV = sum(ITMShortIV) / len(ITMShortIV)
        #Rounds the Value to a Two Decimal Number for Easier Data Translation
        roundAvgITMShortIV = round(AvgITMShortIV, 2)
    
        #Finds The Percentage of OTM and ITM Both Long Vs Short
        OTMLongPerc = 100 * OTMLong / totalLongInt
        ITMLongPerc = 100 * ITMLong / totalLongInt
        OTMShortPerc = 100 * OTMShort / totalShortInt
        ITMShortPerc = 100 * ITMShort / totalShortInt
        #EDIT 2023/03/26 3pm !!!!!!!
        OTMLongVolPerc = 100 * OTMLongVol / totalLongVol
        ITMLongVolPerc = 100 * ITMLongVol / totalLongVol
        OTMShortVolPerc = 100 * OTMShortVol / totalShortVol
        ITMShortVolPerc = 100 * ITMShortVol / totalShortVol


        #Rounds the Values to a Two Decimal Number for Easier Data Translation
        roundOTMLongPerc = round(OTMLongPerc, 2)
        roundITMLongPerc = round(ITMLongPerc, 2)
        roundOTMShortPerc = round(OTMShortPerc, 2)
        roundITMShortPerc = round(ITMShortPerc, 2)

        #EDIT 2023/03/26 3pm
        roundOTMLongVolPerc = round(OTMLongVolPerc, 2)
        roundITMLongVolPerc = round(ITMLongVolPerc, 2)
        roundOTMShortVolPerc = round(OTMShortVolPerc, 2)
        roundITMShortVolPerc = round(ITMShortVolPerc, 2)


        #Appends All Empty Lists That Hold Iterable Data To Each Variable That Holds Information For The Main DataFrame
        expirationDate.append(expDate.strftime('%m/%d/%Y'))
        daysToExpiration.append(DTE)
        totalLongVolumeByExpiration.append(totalLongVol)
        totalShortVolumeByExpiration.append(totalShortVol)
        totalLongOpenInterestByExpiration.append(totalLongInt)
        totalShortOpenInterestByExpiration.append(totalShortInt)
        longVolAsPerc.append(roundLongvolPerc)
        shortVolAsPerc.append(roundShortvolPerc)
        longOpenIntAsPerc.append(roundLongOpenInt)
        shortOpenIntAsPerc.append(roundShortOpenInt)
        OTMLongList.append(OTMLong)
        ITMLongList.append(ITMLong)
        OTMShortList.append(OTMShort)
        ITMShortList.append(ITMShort)
        OTMLongPercList.append(roundOTMLongPerc)
        ITMLongPercList.append(roundITMLongPerc)
        OTMShortPercList.append(roundOTMShortPerc)
        ITMShortPercList.append(roundITMShortPerc)
        AvgIVLongList.append(roundAvgLongIV)
        AvgIVShortList.append(roundAvgShortIV)
        AvgIVOTMLongList.append(roundAvgOTMLongIV)
        AvgIVITMLongList.append(roundAvgITMLongIV)
        AvgIVOTMShortList.append(roundAvgOTMShortIV)
        AvgIVITMShortList.append(roundAvgITMShortIV) 

        #EDIT 2023/03/26 3pm
        OTMLongVolList.append(OTMLongVol)
        ITMLongVolList.append(ITMLongVol)
        OTMShortVolList.append(OTMShortVol)
        ITMShortVolList.append(ITMShortVol)

        OTMLongVolPercList.append(roundOTMLongPerc)
        ITMLongVolPercList.append(roundITMLongPerc)
        OTMShortVolPercList.append(roundOTMShortPerc)
        ITMShortVolPercList.append(roundITMShortPerc)

        
        
        #Interates The On-Going Expiration Dates Until Complete (To Show Progress Of The Programs Data Collection)
        print('Finished with ' + expDate.strftime('%m/%d/%Y') + ' expiration.')
        


    #Variable For Main DataFrame That Holds All Of The Iterated Information
    byExpirationData = pd.DataFrame(data = list(zip(itertools.repeat(noway), expirationDate, totalLongVolumeByExpiration, longVolAsPerc, OTMLongVolList, OTMLongVolPercList, ITMLongVolList, ITMLongVolPercList,\
        totalLongOpenInterestByExpiration, longOpenIntAsPerc, OTMLongList, OTMLongPercList, ITMLongList, ITMLongPercList, AvgIVLongList, AvgIVOTMLongList, \
        AvgIVITMLongList, totalShortVolumeByExpiration, shortVolAsPerc, OTMShortVolList, OTMShortVolPercList, ITMShortVolList, ITMShortVolPercList,\
        totalShortOpenInterestByExpiration, shortOpenIntAsPerc, OTMShortList, OTMShortPercList, ITMShortList, ITMShortPercList,\
        AvgIVShortList, AvgIVOTMShortList, AvgIVITMShortList)), columns= ['DOE', 'Expiration_Date', 'Total_Long_Volume',  'Long_Volume_Percent', 'OTM_Long_Volume', 'OTM_Long_Volume_Percent','ITM_Long_Volume', 'ITM_Long_Volume_Percent',  \
        'Total_Long_Open_Interest', 'Long_Open_Int_Percent', \
        'OTM_Long', 'OTM_Long_Percent', 'ITM_Long', 'ITM_Long_Percent', 'Avg_Long_IV', 'Avg_Long_OTM_IV', 'Avg_Long_ITM_IV','Total_Short_Volume', 'Short_Vol_Percent', 'OTM_Short_Volume', 'OTM_Short_Volume_Percent', 'ITM_Short_Volume', 'ITM_Short_Volume_Percent', 'Total_Short_Open_Interest', \
        'Short_Open_Int_Percent', 'OTM_Short', 'OTM_Short_Percent','ITM_Short', 'ITM_Short_Percent', 'Avg_Short_IV', 'Avg_Short_OTM_IV', 'Avg_Short_ITM_IV'])
    #Prints The Final DataFrame Into The Terminal
    print(byExpirationData)

    #Connects to existing database file in project folder, can be accessed through DB Browser for SQLITE
    conn = sql.connect('name_of_database.db')
    byExpirationData.to_sql('SPY_Extra_Option_Data', conn, if_exists='append', index=False)
    dfMainOptionChain.to_sql('SPY_Option_Chain', conn, if_exists='append', index=False)


    #Exports Dataframe To An Excel Sheet And Uses Variables To Name The File By Ticker Name, Date, and Time (Hour + Minute) In The Format (ticker-year-month-day-hour-minute)
    # file_name = str(ticker) + "-" + str(fileDate)
    # sheet_name_var = str(ticker)
    # file_extension = '.xlsx'



    # #Code to Make Titles Of Each Column in Excel Sheet Fit in The Excel Sheet Cells By Iterating Through Each Column And Resizing It To Fit Titles
    # writer = pd.ExcelWriter(file_name + file_extension, engine='xlsxwriter')
    # byExpirationData.to_excel(writer, sheet_name=sheet_name_var, index=False, na_rep='NaN')
    # for column in byExpirationData:
    #     column_length = max(byExpirationData[column].astype(str).map(len).max(), len(column))
    #     col_idx = byExpirationData.columns.get_loc(column)
    #     writer.sheets[sheet_name_var].set_column(col_idx, col_idx, column_length)
    # #Saves The Excel Sheet To The Folder That Holds Your Python Files
    # writer.save() 

    #ENJOY!!!! 
    #More features coming soon!!!!
