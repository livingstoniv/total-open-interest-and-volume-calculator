# total-open-interest-and-volume-calculator
This program finds the total open long interest both OTM and ITM (with percentages), short interest both OTM and ITM (with percentages), long volume and short volume both overall, OTM and ITM (with percentages), the average implied volatility long and short both OTM and ITM, and it translates that into a percentage of total allocation into a pandas data frame which can be exported into a SQLITE database or excel sheet, simply create a database in DB Browser for SQLITE or just comment out the database part of my code and uncomment the excel sheet export part and run the program and enter a ticker (make sure the ticker has options contracts, or the program may not work). THIS CODE RUNS CONTINOUSLY EVERY 5x MINUTES, TO HAVE IT RUN ONCE REMOVE THE WHILE LOOP, UNINDENT THE CODE BELOW IT, AND REMOVE THE TIME.SLEEP() PART.

The idea behind this program is to see where the most money in the options market is allocated to make better investment decisions.


Dependencies:

yahoo-fin,
pandas,
datetime,
numpy,
xlsxwriter,
openpyxl
