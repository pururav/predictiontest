import xlrd
import sys
import datetime
from datetime import date
from datetime import datetime


# getting the xls file name from the command line
filename = 'Data-2011Mar6.xls'
i=0

for arg in sys.argv:

    if (i==1):
        filename = arg
    i += 1

# Open excel workbook
print 'reading data from file %s...' % (filename)
xl = xlrd.open_workbook(filename)
#print 'date mode: %s' % xl.datemode


dates = []

for sheet_name in xl.sheet_names():
	
	sheet = xl.sheet_by_name(sheet_name)
	
	for colnum in range(sheet.ncols):
		
		symbol = ''
		indices = []
		column = sheet.col_values(colnum)
		
		if ((sheet_name == 'Dates') & (colnum == 2)):
			
			tempDates = column[3:]
			
			for	dateObject in tempDates:
				
				# print date
				date_string = dateObject.split('-')
				
				year = int(date_string[2])
				month = int(date_string[0])
				day = int(date_string[1])
				
				indexDate = date(year, month, day)
				dates.append(indexDate)
				
			#print dates
		
		elif ((sheet_name != 'Dates') & (colnum > 0)):
			
			symbol = column[0]
			# print symbol
			
			indices = column[3:]
			# print len(indices)
			break
		

# for RSI Calculation
Numdays = 28
alpha = 2/(Numdays + 1)
1malpha = 1. - alpha
currIndex = len(indices)

# Upper and Lower time series
for i=1:Numdays+3:
    increment = indices[currIndex - i] - indices[currIndex - i +1]
    Upper[i] = 0.0
    Lower[i] = 0.0
    if(increment > 0):
        Upper[i] = increment
    if(increment < 0):
        Lower[i] = -increment

# exponential moving average
# EMAU[1] = Upper[1]
#EMAD[1] = Lower[1]
#for i=2:Numdays:
#    EMAU[i] = alpha * Upper[i-1] + 1malpha * EMAU[i-1]
#    EMAD[i] = alpha * Lower[i-1] + 1malpha * EMAD[i-1]

#RSI calculation
EMAU[3] = 0.0
EMAL[3] = 0.0
for i = 0:Numdays-1
    EMAU[3] += alpha * pow(1malpha,i) * Upper[i+3]
    EMAL[3] += alpha * pow(1malpha,i) * Lower[i+3]


EMAU[2] = alpha * (Upper[2] - EMAU[3] ) + EMAU[3]
EMAL[2] = alpha * (Lower[2] - EMAL[3] ) + EMAL[3]

EMAU[1] = alpha * (Upper[1] - EMAU[2] ) + EMAU[2]
EMAL[1] = alpha * (Lower[1] - EMAL[2] ) + EMAL[2]

for i=1:3
    rs = EMAU[i]/EMAL[i]
    rsi[i] = 100 - 100/(1+rs)
# fit curve of a*x*x + b*x + c


