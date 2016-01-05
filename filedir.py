import os
import glob
import xlrd


listOfOrders = []

# Set directory where the sheets are 
os.chdir('/Users/fredrikheiberg/Documents/randem/sheets')

# Gather search field conditions and create a list of corresponding files
searchCondition = "05.01.16"
fileList = glob.glob('%s-*.xlsx' %searchCondition)

# Iterate through all corresponding files and create a list with relevant 
# information from each sheet (List in list - a list of all sheets that 
# have a list of relevant information)
for sh in fileList:
	file_location = '/Users/fredrikheiberg/Documents/randem/sheets/%s' %sh
	orderDetails = []
#	print "File Location: " +file_location
	workbook = xlrd.open_workbook(file_location)
	sheet = workbook.sheet_by_index(0)
	#sheet2 = workbook.sheet_by_index(1)

	# Dato, ank tid, sted, oppdrag, kunde, buss, mobil
	# TODO! add functionality if field is not filled in!!!!!

	# Date
	cellValue = sheet.cell_value(4,1)
	cellDateValue = xlrd.xldate_as_tuple(cellValue, workbook.datemode)
	dateString = str(cellDateValue[2])+"."+str(cellDateValue[1])+"."+str(cellDateValue[0])
	orderDetails.append(dateString)

	# Arrival time
	timeOfOrder = sheet.cell_value(5,1)
	orderDetails.append(timeOfOrder)

	# Place
	place = sheet.cell_value(6,1)
	orderDetails.append(place)

	# Order description
	orderDescription = sheet.cell_value(8,1)
	orderDetails.append(orderDescription)

	# Customer
	customer = sheet.cell_value(2,1)
	orderDetails.append(customer)

	# Buss
	buss = sheet.cell_value(10,1)
	orderDetails.append(buss)

	# Telephone number
	cellNumber = sheet.cell_value(11,1)
	orderDetails.append(cellNumber)

	#for detail in orderDetails:
	#	print detail

	listOfOrders.append(orderDetails)
	return listOfOrders

	# 25 mulige lengder, make input from decimal to dateTime
	# row = 12
	# for i in range(row, 36):
	# 	customer = sheet.cell_value(1,1)

	# 	orderDetails = []
	# 	if sheet.cell_value(i,0) != "":
	# 		# Add customer to the list
	# 		orderDetails.append(customer)
	# 		# Date value
	# 		cellValue = sheet.cell_value(i,0)
	# 		cellDateValue = xlrd.xldate_as_tuple(cellValue, workbook.datemode)
	# 		dateString = str(cellDateValue[2])+"."+str(cellDateValue[1])+"."+str(cellDateValue[0])
	# 		orderDetails.append(dateString)

	# 		# Time value
	# 		timeOfOrder = sheet.cell_value(i,1)
	# 		#print timeOfOrder
	# 		orderDetails.append(timeOfOrder)

	# 		# Order description
	# 		orderDescription = sheet.cell_value(i,2)
	# 		#print orderDescription
	# 		orderDetails.append(orderDescription)

	# 		if sheet2.cell_value(i,0) != "":
	# 			buss = sheet2.cell_value(i,0)
	# 			# Add Buss
	# 			orderDetails.append(buss)
	# 		if sheet2.cell_value(i,1) != "":
	# 			cellNumber = sheet2.cell_value(i,1)
	# 			# Add cell number
	# 			orderDetails.append(cellNumber)

	# 		print orderDetails
	# 		listOfOrders.append(orderDetails)
	# 		print "\n"

	#orderNumber = sheet.cell_value(2,5)
	#telephoneNumber = sheet.cell_value(7,5)
	#address = sheet.cell_value(12,1)

	#print sh
	#print orderNumber
	#print telephoneNumber
	#print address
	print "\n"

		# searchCondition = ""
	# listOfOrders = []
	# orderListLength = 0

	# # Set directory where the sheets are 
	# os.chdir('/Users/fredrikheiberg/Documents/randem/sheets')

	# # Gather search field conditions and create a list of corresponding files
	# #searchCondition = "03.01.16"
	# searchCondition = str(request.args.get('date'))
	# print("SEARCH CONDITION: %s" % searchCondition)
	# fileList = glob.glob('%s-*.xlsx' %searchCondition)

# 	# Iterate through all corresponding files and create a list with relevant 
# 	# information from each sheet (List in list - a list of all sheets that 
# 	# have a list of relevant information)
# 	for sh in fileList:
# 		file_location = '/Users/fredrikheiberg/Documents/randem/sheets/%s' %sh
# #		print "File Location: " +file_location
# 		workbook = xlrd.open_workbook(file_location)
# 		sheet = workbook.sheet_by_index(0)
# 		sheet2 = workbook.sheet_by_index(1)

# 		# 25 mulige lengder, make input from decimal to dateTime
# 		row = 12
# 		for i in range(row, 36):
# 			customer = sheet.cell_value(1,1)

# 			orderDetails = []
# 			if sheet.cell_value(i,0) != "":
# 				# Date value
# 				cellValue = sheet.cell_value(i,0)
# 				print("CELL VALUE: %s" %cellValue)
# 				cellDateValue = xlrd.xldate_as_tuple(cellValue, workbook.datemode)
# 				dateString = str(cellDateValue[2])+"."+str(cellDateValue[1])+"."+str(cellDateValue[0])
# 				orderDetails.append(dateString)

# 				# Time value
# 				timeOfOrder = sheet.cell_value(i,1)
# 				#print timeOfOrder
# 				orderDetails.append(timeOfOrder)

# 				# Order description
# 				orderDescription = sheet.cell_value(i,2)
# 				#print orderDescription
# 				orderDetails.append(orderDescription)

# 				if sheet2.cell_value(i,0) != "":
# 					buss = sheet2.cell_value(i,0)
# 					# Add Buss
# 					orderDetails.append(buss)
# 				if sheet2.cell_value(i,1) != "":
# 					cellNumber = sheet2.cell_value(i,1)
# 					# Add cell number
# 					orderDetails.append(cellNumber)

# 				# Add customer to the list
# 				orderDetails.append(customer)

# 				listOfOrders.append(orderDetails)




