# total-open-interest-and-volume-calculator
This program finds the total open long interest both OTM and ITM (with percentages), short interest both OTM and ITM (with percentages), long volume and short volume both overall, OTM and ITM (with percentages), the average implied volatility long and short both OTM and ITM, and it translates that into a percentage of total allocation into a pandas data frame which can be exported into an excel sheet, simply run the program and enter a ticker (make sure the ticker has options contracts, or the program may not work). 

The idea behind this program is to see where the most money in the options market is allocated to make better investment decisions.


Dependencies:

yahoo-fin,
pandas,
datetime,
numpy,
xlsxwriter,
openpyxl
