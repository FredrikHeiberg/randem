import os
import glob
from xlutils.copy import copy
from xlrd import *
from openpyxl import Workbook
import xlwt
from xlutils.filter import process,XLRDReader,XLWTWriter

wb = Workbook()

os.chdir('/Users/fredrikheiberg/Documents/randem/static/sheets')

searchCondition = "mal"
fileList = glob.glob('%s.xlsx' %searchCondition)








listOfthings = ['2','3','5']

print listOfthings[len(listOfthings)-1]


w = copy(open_workbook('/Users/fredrikheiberg/Documents/randem/static/sheets/alternativeMal.xls',formatting_info=True))
w.get_sheet(0).write(1,1,"Dette er en test")
w.get_sheet(0).write_value(2,1,"Dette er en test")
#w.get_sheet(0).write(4,1,"Dette er en test")
#w.get_sheet(0).write(5,1,"Dette er en test")
#w.get_sheet(0).write(6,1,"Dette er en test")
#w.get_sheet(0).write(7,1,"Dette er en test")
#w.get_sheet(0).write(8,1,"Dette er en test")
#w.get_sheet(0).write(9,1,"Dette er en test")
#w.get_sheet(0).write(10,1,"Dette er en test")
#w.get_sheet(0).write(11,1,"Dette er en test")
#w.get_sheet(0).write(12,1,"Dette er en test")
w.save('/Users/fredrikheiberg/Documents/randem/static/sheets/notWebPage.xls')





# for sh in fileList:
# 	file_location = '/Users/fredrikheiberg/Documents/randem/static/sheets/%s' %sh
# 	orderDetails = []
# 	#workbook = xlrd.open_workbook(file_location)
# 	sheet = workbook.sheet_by_index(0)

# 	# data = sheet.cell_value(1,0)

# 	# rb = open_workbook(file_location)
# 	# r_sheet = rb.sheet_by_index(0)
# 	# wb = copy(rb) 
# 	# w_sheet = wb.get_sheet(0)

# 	# w_sheet.write(0, 0, 'TEST')
# 	# wb.save(file_location + 'test' + os.path.splitext(file_location)[-1])

# 	w = copy(open_workbook(file_location))
# 	print w
	#w.get_sheet(0).write(1,1,"foo")
	#w.save('%sbook2.xls'%file_location)

	#print "R_SHEET: %s" %r_sheet

	# Date
# 	cellValue = sheet.cell_value(4,1)
# 	cellDateValue = xlrd.xldate_as_tuple(cellValue, workbook.datemode)
# 	dateString = str(cellDateValue[2])+"."+str(cellDateValue[1])+"."+str(cellDateValue[0])
# 	orderDetails.append(dateString)

# 	# Arrival time
# 	timeOfOrder = sheet.cell_value(5,1)
# 	orderDetails.append(timeOfOrder)

# 	# Place
# 	place = sheet.cell_value(6,1)
# 	orderDetails.append(place)

# 	# Order description
# 	orderDescription = sheet.cell_value(8,1)
# 	orderDetails.append(orderDescription)

# 	# Customer
# 	customer = sheet.cell_value(2,1)
# 	orderDetails.append(customer)

# 	# Buss
# 	buss = sheet.cell_value(10,1)
# 	orderDetails.append(buss)

# 	# Telephone number
# 	cellNumber = sheet.cell_value(11,1)
# 	orderDetails.append(cellNumber)

# 	#for detail in orderDetails:
# 	#	print detail

# 	listOfOrders.append(orderDetails)
# return listOfOrders




