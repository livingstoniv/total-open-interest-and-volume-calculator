
from yahoo_fin import options as op
import pandas as pd
from datetime import datetime, date
import numpy as np


expirationDate = []
daysToExpiration = []
totalLongVolumeByExpiration = []
totalShortVolumeByExpiration = []
totalLongOpenInterestByExpiration = []
totalShortOpenInterestByExpiration = []

ticker = input('Enter your ticker here (make sure it has options data): ')

today = date.today()
todaysDate = datetime(today.year, today.month, today.day)

expDates = op.get_expiration_dates(ticker)

for i in range(0, len(expDates)):
    expDate = datetime.strptime(expDates[i], '%B %d, %Y')
    DTE = (expDate - todaysDate).days

    chain = op.get_options_chain(ticker, date= expDates[i])


# print(expDates)
# print(chain)

    chain['calls']['Volume'] = pd.to_numeric(chain['calls']['Volume'], errors = 'coerce')
    chain['calls']['Open Interest'] = pd.to_numeric(chain['calls']['Open Interest'], errors = 'coerce')


    chain['puts']['Volume'] = pd.to_numeric(chain['puts']['Volume'], errors = 'coerce')
    chain['puts']['Open Interest'] = pd.to_numeric(chain['puts']['Open Interest'], errors = 'coerce')

    totalLongVol = chain['calls']['Volume'].sum()
    totalLongInt = chain['calls']['Open Interest'].sum()

    totalShortVol = chain['puts']['Volume'].sum()
    totalShortInt = chain['puts']['Open Interest'].sum()

    expirationDate.append(expDate.strftime('%m/%d/%Y'))
    daysToExpiration.append(DTE)
    totalLongVolumeByExpiration.append(totalLongVol)
    totalShortVolumeByExpiration.append(totalShortVol)
    totalLongOpenInterestByExpiration.append(totalLongInt)
    totalShortOpenInterestByExpiration.append(totalShortInt)

    print('Finished with ' + expDate.strftime('%m/%d/%Y') + ' expiration.')

byExpirationData = pd.DataFrame(data = list(zip(expirationDate, daysToExpiration, totalLongVolumeByExpiration, totalLongOpenInterestByExpiration, totalShortVolumeByExpiration, totalShortOpenInterestByExpiration)), columns= ['Expiration Date', 'DTE', 'Total Long Volume', 'Total Long Open Interest', 'Total Short Volume', 'Total Short Open Interest'])
print(byExpirationData)

byExpirationData.to_excel('SpyData.xlsx', sheet_name='Spy Data')
