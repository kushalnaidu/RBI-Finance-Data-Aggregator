import xlrd 
import calendar
import datetime
import glob
import os
list_of_months = list(calendar.month_abbr)

#First documentation begins from september 2012
initial_date = datetime.datetime(2012, 9, 1)
thisdir = os.getcwd()

excel_files = []

def diff_month(d1, d2):
	    return (d1.year - d2.year) * 12 + d1.month - d2.month
'''
for file in glob.glob("*.xls"):
    excel_files.append(file)

'''
for r, d, f in os.walk(thisdir):
    for file in f:
        if ".xls" in file.lower():
            excel_files.append(os.path.join(r, file))

#print(excel_files)
row = 44
col = 7
data = {}
for file in excel_files:
	wb = xlrd.open_workbook(file)
	sheet = wb.sheet_by_index(0)
	row = 44
	col = 7
	heading_row = 3
	

	for i in range(3):
		temp = i

		year = sheet.cell_value(heading_row,(col+i))
		month = sheet.cell_value(heading_row+1,col+i)
		print(file)
		print(year,month)
		if '-' in str(year):
			col+=1
			year = sheet.cell_value(heading_row,(col+i))
			month = sheet.cell_value(heading_row+1,col+i)
		if type(month) != str:
			heading_row+=1
			year = sheet.cell_value(heading_row,(col+i))
			month = sheet.cell_value(heading_row+1,col+i)
		if month[-1] == '.':
			month = month[:-1] #remove the '.' Takes only the name of the month
		while year == '':
			temp-=1
			year = sheet.cell_value(3,col+temp)
		year = int(year)
		diff = diff_month(datetime.datetime(year,list_of_months.index(month),1) , initial_date)
		#print(diff,sheet.cell_value(44,7+i))
		val = sheet.cell_value(row,col+i)
		while val == '–':
			row+=1
			val = sheet.cell_value(row,col+i)
		data[diff] = sheet.cell_value(row,col+i)
print(data)
