from yahoo_fin import options as op
import pandas as pd
from datetime import datetime, date
import numpy as np
from yahoo_fin import stock_info as si
from xlsxwriter.utility import xl_rowcol_to_cell
import openpyxl as px


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
#OTM and ITM Percentage Both Long and Short Lists
OTMLongPercList = []
ITMLongPercList = []
OTMShortPercList = []
ITMShortPercList = []
#implied volatility lists
AvgIVLongList = []
AvgIVOTMLongList = []
AvgIVITMLongList = []
AvgIVShortList = []
AvgIVOTMShortList = []
AvgIVITMShortList = []



#Input For Finding Stock Info, Run The Program Then Enter The Ticker Into The Terminal
ticker = input("Enter Stock Ticker Here (Make Sure It Has Option Data: ")


#Expiration Date Variables
today = date.today()
now = datetime.now()
fileDate = now.strftime("%Y-%m-%d-%H-%M")
todaysDate = datetime(today.year, today.month, today.day)
expDates = op.get_expiration_dates(ticker)

#Loop For Fiding Expiration Dates, Iterating Data Through Each One and Combining Data Into a Central Structure
for i in range(0, len(expDates)):
    expDate = datetime.strptime(expDates[i], '%B %d, %Y')
    DTE = (expDate - todaysDate).days

    chain = op.get_options_chain(ticker, date= expDates[i])
    livePrice = si.get_live_price(ticker)
    





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
    maxLongOpenInt = max(LongOpenInt)
    maxShortOpenInt = max(ShortOpenInt)

     #Creates A Seperate DataFrame For Further Equations
    df = pd.DataFrame(data = list(zip(LongOpenInt, LongStrike, floatLongIV, ShortOpenInt, ShortStrike, floatShortIV)), columns= ['Long Open Int', 'Long Strike', 'Long IV', 'Short Open Int', 'Short Strike', 'Short IV'])

    
 
    #Finds The Average Long Vs Short IV
    AvgLongIV = sum(floatLongIV) / len(floatLongIV)
    AvgShortIV = sum(floatShortIV) / len(floatShortIV)
    
    #Rounds The Values To A Two Decimal Number For Easier Data Translation
    roundAvgLongIV = round(AvgLongIV, 2)
    roundAvgShortIV = round(AvgShortIV, 2)
  
    
    
    #Finds all of the OTM Long Open Interest 
    OTMLong = df[df['Long Strike'] < livePrice]['Long Open Int'].sum()

    
    #Finds all of the ITM Long Open Interest
    ITMLong = df[df['Long Strike'] > livePrice]['Long Open Int'].sum()
    # print('Total ITM Long: ' + str(ITMLong))
    
    #Finds all of the OTM Long Implied Volatility
    OTMLongIV = df[df['Long Strike'] < livePrice]['Long IV']
    
    #Finds all of the ITM Long Implied Volatility
    ITMLongIV = df[df['Long Strike'] > livePrice]['Long IV']
    
    #Finds all of the OTM Short
    OTMShort = df[df['Short Strike'] > livePrice]['Short Open Int'].sum()
    # print("Total OTM Short: " + str(OTMShort))
    
    #Finds all of the ITM Short
    ITMShort = df[df['Short Strike'] < livePrice]['Short Open Int'].sum()
    
    #Finds all of the OTM Short Implied Volatility
    OTMShortIV = df[df['Short Strike'] > livePrice]['Short IV']
     
    #Finds all of the ITM Short Implied Volatility
    ITMShortIV = df[df['Short Strike'] < livePrice]['Short IV']
    
    #Finds all of the OTM Long IV
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
    
    #Rounds the Values to a Two Decimal Number for Easier Data Translation
    roundOTMLongPerc = round(OTMLongPerc, 2)
    roundITMLongPerc = round(ITMLongPerc, 2)
    roundOTMShortPerc = round(OTMShortPerc, 2)
    roundITMShortPerc = round(ITMShortPerc, 2)



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

    
    
    #Interates The On-Going Expiration Dates Until Complete (To Show Progress Of The Programs Data Collection)
    print('Finished with ' + expDate.strftime('%m/%d/%Y') + ' expiration.')

#Variable For Main DataFrame That Holds All Of The Iterated Information
byExpirationData = pd.DataFrame(data = list(zip(expirationDate, daysToExpiration, totalLongVolumeByExpiration, longVolAsPerc,\
     totalLongOpenInterestByExpiration, longOpenIntAsPerc, OTMLongList, OTMLongPercList, ITMLongList, ITMLongPercList, AvgIVLongList, AvgIVOTMLongList, \
     AvgIVITMLongList, totalShortVolumeByExpiration, shortVolAsPerc, \
     totalShortOpenInterestByExpiration, shortOpenIntAsPerc, OTMShortList, OTMShortPercList, ITMShortList, ITMShortPercList,\
     AvgIVShortList, AvgIVOTMShortList, AvgIVITMShortList)), columns= ['Expiration Date', 'DTE', 'Total Long Volume', 'Long Vol (Percent)', 'Total Long Open Interest', 'Long Open Int (Percent)', \
    'OTM Long', 'OTM Long Percent', 'ITM Long', 'ITM Long Percent', 'Avg Long IV', 'Avg Long OTM IV', 'Avg Long ITM IV','Total Short Volume', 'Short Vol (Percent)', 'Total Short Open Interest', \
    'Short Open Int (Percent)', 'OTM Short', 'OTM Short Percent','ITM Short', 'ITM Short Percent', 'Avg Short IV', 'Avg Short OTM IV', 'Avg Short ITM IV'])
#Prints The Final DataFrame Into The Terminal
print(byExpirationData)

#Exports Dataframe To An Excel Sheet And Uses Variables To Name The File By Ticker Name, Date, and Time (Hour + Minute) In The Format (ticker-year-month-day-hour-minute)
file_name = str(ticker) + "-" + str(fileDate)
sheet_name_var = str(ticker)
file_extension = '.xlsx'

#Code to Make Titles Of Each Column in Excel Sheet Fit in The Excel Sheet Cells By Iterating Through Each Column And Resizing It To Fit Titles
writer = pd.ExcelWriter(file_name + file_extension, engine='xlsxwriter')
byExpirationData.to_excel(writer, sheet_name=sheet_name_var, index=False, na_rep='NaN')
for column in byExpirationData:
    column_length = max(byExpirationData[column].astype(str).map(len).max(), len(column))
    col_idx = byExpirationData.columns.get_loc(column)
    writer.sheets[sheet_name_var].set_column(col_idx, col_idx, column_length)
#Saves The Excel Sheet To The Folder That Holds Your Python Files
writer.save() 

#ENJOY!!!! 

#P.S. More Data Coming Soon!!!!
