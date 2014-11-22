import os
import xlwt
from xlrd import *
import re

fpath = '/Users/azade/Desktop/TweetsInDays/Iran-Election-3M-Merged-Data-DEDUP.tsv'
book = open_workbook(fpath)

#initiate a workbook to add rows to
rtbook = xlwt.Workbook()
rtsheet = rtbook.add_sheet('Retweets')

sheet = book.sheet_by_index(0)

# cell = sheet.cell(0,0)
# print cell.value
# print sheet.nrows
line = 1
recipient = re.compile('@[\w]+:', re.IGNORECASE)
for i in range(sheet.nrows): #sheet.nrows
	# for j in range(sheet.ncols):
 	if (sheet.cell(i,7).ctype==XL_CELL_TEXT and sheet.cell_value(i,7).startswith('RT')):
 		print i
 		try:
	 		for j in range(sheet.ncols):
	 			# print j, sheet.cell_value(i,j)
	 			rtsheet.write(line,j,sheet.cell_value(i,j))
	 		rtsheet.write(line,16,'RT')
 		
	 		recipient = re.findall('@[\w]+',sheet.cell_value(i,7))
	 		# print recipient[0]
	 		rtsheet.write(line,17,recipient[0])
	 		tweet_text = sheet.cell_value(i,7).split(recipient[0])[-1]
	 		# print tweet_text
	 		rtsheet.write(line,18,tweet_text.strip(': '))
	 		print "inserted row number", line
	 		line += 1	
	 	except:
	 		print "Error on row number", i
	 		line += 1
	 		continue

rtbook.save('Retweets_Election.xls')
print "Number of rows inserted: ", line
 	# for i in range(5): #sheet.ncols
  #   print i,sheet.cell_value(1,i)

