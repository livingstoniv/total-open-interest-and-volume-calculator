from yahoo_fin import options as op
import pandas as pd
from datetime import datetime, date
import numpy as np
from yahoo_fin import stock_info as si
from xlsxwriter.utility import xl_rowcol_to_cell
import openpyxl as px


#Empty Lists For Data
#date lists
expirationDate = []
daysToExpiration = []
#volume lists
totalLongVolumeByExpiration = []
totalShortVolumeByExpiration = []
#open int lists
totalLongOpenInterestByExpiration = []
totalShortOpenInterestByExpiration = []
#open interest and volume percentage lists
longOpenIntAsPerc = []
shortOpenIntAsPerc = []
longVolAsPerc = []
shortVolAsPerc = []
#OTM and ITM both Long and Short Lists
OTMLongList = []
ITMLongList = []
OTMShortList = []
ITMShortList = []
#otm and itm percentage lists
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
ticker = 'SPY'


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
    





#Creates Numeric Data Structure
    chain['calls']['Volume'] = pd.to_numeric(chain['calls']['Volume'], errors = 'coerce')
    chain['calls']['Open Interest'] = pd.to_numeric(chain['calls']['Open Interest'], errors = 'coerce')


    chain['puts']['Volume'] = pd.to_numeric(chain['puts']['Volume'], errors = 'coerce')
    chain['puts']['Open Interest'] = pd.to_numeric(chain['puts']['Open Interest'], errors = 'coerce')

    LongOpenInt = chain['calls']['Open Interest']
    ShortOpenInt = chain['puts']['Open Interest']
    LongStrike = chain['calls']['Strike']
    ShortStrike = chain['puts']['Strike']
    LongIV = chain['calls']['Implied Volatility']
    ShortIV = chain['puts']['Implied Volatility']

    LongIVStrip = [i.strip('%') for i in LongIV]
    ShortIVStrip = [c.strip('%') for c in ShortIV]

    floatLongIV = [float(g) for g in LongIVStrip]
    floatShortIV = [float(w) for w in ShortIVStrip]



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

    df = pd.DataFrame(data = list(zip(LongOpenInt, LongStrike, floatLongIV, ShortOpenInt, ShortStrike, floatShortIV)), columns= ['Long Open Int', 'Long Strike', 'Long IV', 'Short Open Int', 'Short Strike', 'Short IV'])
    # print(df)
    
 

    AvgLongIV = sum(floatLongIV) / len(floatLongIV)
    AvgShortIV = sum(floatShortIV) / len(floatShortIV)
    roundAvgLongIV = round(AvgLongIV, 2)
    roundAvgShortIV = round(AvgShortIV, 2)
    # print(roundAvgLongIV)
    # print(roundAvgShortIV)


    OTMLong = df[df['Long Strike'] < livePrice]['Long Open Int'].sum()
    # print("Total OTM Long: " + str(OTMLong))
    ITMLong = df[df['Long Strike'] > livePrice]['Long Open Int'].sum()
    # print('Total ITM Long: ' + str(ITMLong))
    OTMLongIV = df[df['Long Strike'] < livePrice]['Long IV']
    ITMLongIV = df[df['Long Strike'] > livePrice]['Long IV']
    OTMShort = df[df['Short Strike'] < livePrice]['Short Open Int'].sum()
    # print("Total OTM Short: " + str(OTMShort))
    ITMShort = df[df['Short Strike'] > livePrice]['Short Open Int'].sum()
    OTMShortIV = df[df['Short Strike'] < livePrice]['Short IV']
    ITMShortIV = df[df['Short Strike'] > livePrice]['Short IV']
    # print("Total ITM Short: "+ str(ITMShort))
    AvgOTMLongIV = sum(OTMLongIV) / len(OTMLongIV)
    roundAvgOTMLongIV = round(AvgOTMLongIV, 2)
    # print(roundAvgOTMLongIV)
    AvgITMLongIV = sum(ITMLongIV) / len(ITMLongIV)
    roundAvgITMLongIV = round(AvgITMLongIV, 2)
    # print(roundAvgITMLongIV)
    AvgOTMShortIV = sum(OTMShortIV) / len(OTMShortIV)
    roundAvgOTMShortIV = round(AvgOTMShortIV, 2)
    # print(roundAvgOTMShortIV)
    AvgITMShortIV = sum(ITMShortIV) / len(ITMShortIV)
    roundAvgITMShortIV = round(AvgITMShortIV, 2)
    # print(roundAvgITMShortIV)

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
    AvgIVLongList.append(roundAvgLongIV)
    AvgIVShortList.append(roundAvgShortIV)
    AvgIVOTMLongList.append(roundAvgOTMLongIV)
    AvgIVITMLongList.append(roundAvgITMLongIV)
    AvgIVOTMShortList.append(roundAvgOTMShortIV)
    AvgIVITMShortList.append(roundAvgITMShortIV) 

    
    
#Interates The On-Going Expiration Dates Until Complete
    print('Finished with ' + expDate.strftime('%m/%d/%Y') + ' expiration.')

byExpirationData = pd.DataFrame(data = list(zip(expirationDate, daysToExpiration, totalLongVolumeByExpiration, longVolAsPerc,\
     totalLongOpenInterestByExpiration, longOpenIntAsPerc, OTMLongList, OTMLongPercList, ITMLongList, ITMLongPercList, AvgIVLongList, AvgIVOTMLongList, \
     AvgIVITMLongList, totalShortVolumeByExpiration, shortVolAsPerc, \
     totalShortOpenInterestByExpiration, shortOpenIntAsPerc, OTMShortList, OTMShortPercList, ITMShortList, ITMShortPercList,\
     AvgIVShortList, AvgIVOTMShortList, AvgIVITMShortList)), columns= ['Expiration Date', 'DTE', 'Total Long Volume', 'Long Vol (Percent)', 'Total Long Open Interest', 'Long Open Int (Percent)', \
    'OTM Long', 'OTM Long Percent', 'ITM Long', 'ITM Long Percent', 'Avg Long IV', 'Avg Long OTM IV', 'Avg Long ITM IV','Total Short Volume', 'Short Vol (Percent)', 'Total Short Open Interest', \
    'Short Open Int (Percent)', 'OTM Short', 'OTM Short Percent','ITM Short', 'ITM Short Percent', 'Avg Short IV', 'Avg Short OTM IV', 'Avg Short ITM IV'])
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
