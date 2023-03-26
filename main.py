from yahoo_fin import options as op
import pandas as pd
from datetime import datetime, date
import numpy as np
from yahoo_fin import stock_info as si
from xlsxwriter.utility import xl_rowcol_to_cell
import openpyxl as px


#Empty Lists For Data
expirationDate = []
daysToExpiration = []
totalLongVolumeByExpiration = []
totalShortVolumeByExpiration = []
totalLongOpenInterestByExpiration = []
totalShortOpenInterestByExpiration = []
longOpenIntAsPerc = []
shortOpenIntAsPerc = []
longVolAsPerc = []
shortVolAsPerc = []
OTMLongList = []
ITMLongList = []
OTMShortList = []
ITMShortList = []
OTMLongPercList = []
ITMLongPercList = []
OTMShortPercList = []
ITMShortPercList = []



#Input For Finding Stock Info, Run The Program Then Enter The Ticker Into The Terminal
ticker = input("Enter Ticker Here (Make Sure It Has Option Data): ")


#Expiration Date Variables
today = date.today()
now = datetime.now()
fileDate = now.strftime("%Y-%m-%d-%H-%M")
todaysDate = datetime(today.year, today.month, today.day)

expDates = op.get_expiration_dates(ticker)

#Loop for fiding expiration dates and combining data into a central structure
for i in range(0, len(expDates)):
    expDate = datetime.strptime(expDates[i], '%B %d, %Y')
    DTE = (expDate - todaysDate).days

    chain = op.get_options_chain(ticker, date= expDates[i])
    livePrice = si.get_live_price(ticker)
    



# print(expDates)
# print(chain)

#Creates Numeric Data Structure
    chain['calls']['Volume'] = pd.to_numeric(chain['calls']['Volume'], errors = 'coerce')
    chain['calls']['Open Interest'] = pd.to_numeric(chain['calls']['Open Interest'], errors = 'coerce')


    chain['puts']['Volume'] = pd.to_numeric(chain['puts']['Volume'], errors = 'coerce')
    chain['puts']['Open Interest'] = pd.to_numeric(chain['puts']['Open Interest'], errors = 'coerce')

    LongOpenInt = chain['calls']['Open Interest']
    ShortOpenInt = chain['puts']['Open Interest']
    LongStrike = chain['calls']['Strike']
    ShortStrike = chain['puts']['Strike']


    totalLongVol = chain['calls']['Volume'].sum()
    totalLongInt = chain['calls']['Open Interest'].sum()

    totalShortVol = chain['puts']['Volume'].sum()
    totalShortInt = chain['puts']['Open Interest'].sum()


    totalVol = totalLongVol + totalShortVol 
    totalOpenInt = totalLongInt + totalShortInt

    longVolPercent = 100 * totalLongVol / totalVol
    shortVolPercent = 100 * totalShortVol / totalVol

    longOpenIntPecent = 100 * totalLongInt / totalOpenInt
    shortOpenIntPercent = 100 * totalShortInt / totalOpenInt

    roundLongvolPerc = round(longVolPercent, 2)
    roundShortvolPerc = round(shortVolPercent, 2)
    roundLongOpenInt = round(longOpenIntPecent, 2)
    roundShortOpenInt = round(shortOpenIntPercent, 2)
    
    maxLongOpenInt = max(LongOpenInt)
    maxShortOpenInt = max(ShortOpenInt)

    df = pd.DataFrame(data = list(zip(LongOpenInt, LongStrike, ShortOpenInt, ShortStrike)), columns= ['Long Open Int', 'Long Strike', 'Short Open Int', 'Short Strike'])
    # print(df)
    OTMLong = df[df['Long Strike'] < livePrice]['Long Open Int'].sum()
    print("Total OTM Long: " + str(OTMLong))
    ITMLong = df[df['Long Strike'] > livePrice]['Long Open Int'].sum()
    print('Total ITM Long: ' + str(ITMLong))
    OTMShort = df[df['Short Strike'] < livePrice]['Short Open Int'].sum()
    print("Total OTM Short: " + str(OTMShort))
    ITMShort = df[df['Short Strike'] > livePrice]['Short Open Int'].sum()
    print("Total ITM Short: "+ str(ITMShort))
    
    OTMLongPerc = 100 * OTMLong / totalLongInt
    ITMLongPerc = 100 * ITMLong / totalLongInt
    OTMShortPerc = 100 * OTMShort / totalShortInt
    ITMShortPerc = 100 * ITMShort / totalShortInt

    roundOTMLongPerc = round(OTMLongPerc, 2)
    roundITMLongPerc = round(ITMLongPerc, 2)
    roundOTMShortPerc = round(OTMShortPerc, 2)
    roundITMShortPerc = round(ITMShortPerc, 2)



    
#Appends Expiration Date Info, Open Interest Info, and Volume Info
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

    
    
#Interates The On-Going Expiration Dates Until Complete
    print('Finished with ' + expDate.strftime('%m/%d/%Y') + ' expiration.')

byExpirationData = pd.DataFrame(data = list(zip(expirationDate, daysToExpiration, totalLongVolumeByExpiration, longVolAsPerc,\
     totalLongOpenInterestByExpiration, longOpenIntAsPerc, OTMLongList, OTMLongPercList, ITMLongList, ITMLongPercList, totalShortVolumeByExpiration, shortVolAsPerc, \
     totalShortOpenInterestByExpiration, shortOpenIntAsPerc, OTMShortList, OTMShortPercList, ITMShortList, ITMShortPercList)), columns= ['Expiration Date', 'DTE', 'Total Long Volume', 'Long Vol (Percent)', 'Total Long Open Interest', 'Long Open Int (Percent)', \
    'OTM Long', 'OTM Long Percent', 'ITM Long', 'ITM Long Percent', 'Total Short Volume', 'Short Vol (Percent)', 'Total Short Open Interest', \
    'Short Open Int (Percent)', 'OTM Short', 'OTM Short Percent','ITM Short', 'ITM Short Percent'])
print(byExpirationData)
#Exports Dataframe To An Excel Sheet, Inputs Are Used To Make Naming The File In The Terminal Easier
file_name = str(ticker) + "-" + str(fileDate)
sheet_name_var = str(ticker)
file_extension = '.xlsx'

#Code to make column titles in excel sheet fit in Excel sheet cells
writer = pd.ExcelWriter(file_name + file_extension, engine='xlsxwriter')
byExpirationData.to_excel(writer, sheet_name=sheet_name_var, index=False, na_rep='NaN')
for column in byExpirationData:
    column_length = max(byExpirationData[column].astype(str).map(len).max(), len(column))
    col_idx = byExpirationData.columns.get_loc(column)
    writer.sheets[sheet_name_var].set_column(col_idx, col_idx, column_length)
writer.save() 
#ENJOY
