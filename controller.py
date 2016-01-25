from routes import *

##### excel related functions #####

def getInfoFromExcel():
	global listOfOrders
	listOfOrders = []
	searchConditionList = []
	datetimeList = []
	# Set directory where the sheets are 
	os.chdir(str(UPLOAD_FOLDER))

	# Gather search field conditions and create a list of corresponding files
	searchConditionOne = str(request.args.get('date'))
	searchConditionList.append(searchConditionOne)
	searchConditionTwo = str(request.args.get('date2'))
	searchConditionList.append(searchConditionTwo)

	if (len(searchConditionList) == 2):
		# if statement does not work (both of them!)
		if re.match('\d{2}.\d{2}.\d{4}',searchConditionList[0]) is not None:
			firstDate = datetime.datetime.strptime(searchConditionList[0].translate(None, '.'), '%d%m%Y').date()
		else:
			firstDate = ""
			secondDate = ""

		if re.match('\d{2}.\d{2}.\d{4}',searchConditionList[1]) is not None:
			secondDate = datetime.datetime.strptime(searchConditionList[1].translate(None, '.'), '%d%m%Y').date()
			#print "S: %s" %secondDate
		else:
			secondDate = 0
			loopDatesExcel(searchConditionList[0])

		if secondDate != 0:
			delta = secondDate - firstDate
			day = datetime.timedelta(days=1)
			for i in range(delta.days + 1):
				searchDate = firstDate + td(days=i)
				dateToString = searchDate.strftime('%d.%m.%Y')
				loopDatesExcel(dateToString)

	else:
		loopDatesExcel(searchConditionList[0])
	return listOfOrders


### TODO: Change to loop over date value in sheet - not file name (DONE?)
def loopDatesExcel(dateValue):

	global listOfOrders
	searchCondition = dateValue
	#fileList = glob.glob('%s-*.xls*' %searchCondition)

	fileList = glob.glob('*.xls*')

	# Iterate through all corresponding files and create a list with relevant 
	# information from each sheet (List in list - a list of all sheets that 
	# have a list of relevant information)
	for sh in fileList:
		tempFileLocation = str(UPLOAD_FOLDER)
		file_location = tempFileLocation + "/%s" %sh
		orderDetails = []
	
		workbook = xlrd.open_workbook(file_location)
		sheet = workbook.sheet_by_index(0)

		#if re.match('\d{2}.\d{2}.\d{4}', getCellInfo(3,1,sheet)) is not None:
			#print "YEYEYEYEYEYEYE"

		print "TEST CELL VALUE %s" %getCellInfo(3,1,sheet)
		if getCellInfo(3,1,sheet) == searchCondition and re.match('\d{2}.\d{2}.\d{4}', getCellInfo(3,1,sheet)) is not None:
			#print "CONDITION MEET"
			# Dato, ank tid, sted, oppdrag, kunde, buss, mobil
			# TODO! add functionality if field is not filled in!!!!!

			# Date
			cellValue = getCellInfo(3,1,sheet)
			#if cellValue != "Ikke spesifisert":
			orderDetails.append(cellValue)

			#print "TEEEEESSSSSSTTTT"
			#cellDateValue = xlrd.xldate_as_tuple(cellValue, workbook.datemode)
			#dateString = str(cellDateValue[2])+"."+str(cellDateValue[1])+"."+str(cellDateValue[0])
			#orderDetails.append(dateString)
			#else:
			#orderDetails.append(cellValue)

			# Arrival time
			timeOfOrder = getCellInfo(4,1,sheet)
			orderDetails.append(timeOfOrder)

			# Place
			place = getCellInfo(5,1,sheet)
			orderDetails.append(place)

			# Order description
			orderDescription = getCellInfo(7,1,sheet)
			orderDetails.append(orderDescription)

			# Customer
			customer = getCellInfo(2,1,sheet)
			orderDetails.append(customer)

			# Number of People
			tempNumberOfPeople = getCellInfo(6,1,sheet)
			numberOfPeople = str(tempNumberOfPeople).split(".")
			orderDetails.append(numberOfPeople[0])

			# Buss
			buss = sheet.cell_value(9,1)
			orderDetails.append(buss)

			# Telephone number
			cellNumber = sheet.cell_value(10,1)
			orderDetails.append(cellNumber)

			listOfOrders.append(orderDetails)

			sortedList = listOfOrders.sort();
#	return listOfOrders

def getCellInfo(row,col,sheet):
	#print "ROW %s COL %s" %(row, col)
	if sheet.cell_value(row,col) != "":
		return sheet.cell_value(row,col)
	else:
		return "Ikke spesifisert"

def allowed_file(filename):
	return '.' in filename and \
			filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

def getListOfSheets():
	global formatedList;
	newFormatedList = []
	formatedList = os.listdir(UPLOAD_FOLDER)
	for string in formatedList:
		newFormatedList.append(string.decode('utf-8'))
	formatedList = newFormatedList
	listOfUploadedFiles = []
	for element in formatedList:
		#tempString = element.split("/")
		if element.endswith('.xlsx') or element.endswith('.xls'):
			listOfUploadedFiles.append(element)
		#listOfUploadedFiles.append(tempString[1])
	return listOfUploadedFiles


##### delete/upload sheet functions #####

def delete_item(item_id):
    #new_id = item_id
    #item = self.session.query(Item).get(item_id)
    #os.remove(os.path.join(app.config['UPLOAD_FOLDER'], item.filename))
    #self.session.delete(item)
    #db.session.commit()
    #print "DETTE SKAL SLETTES %s" %item_id
    os.remove(str(UPLOAD_FOLDER)+"/%s" %item_id)

def download_item(item_id):
	item_id = item_id
	getItemName = item_id.split("/")
	fileId = getItemName[-1]
	return send_from_directory(UPLOAD_FOLDER, fileId)


##### create/edit sheet functions #####

def createDocument():
	global infoList
	searchCondition = "template.xls"

	#workbook = copy(open_workbook(UPLOAD_FOLDER+"/%s"%searchCondition, formatting_info=True, on_demand=True))
	workbook = xlrd.open_workbook(TEMPLATE_FOLDER+"/%s"%searchCondition, formatting_info=True, on_demand=True)
	inSheet = workbook.sheet_by_index(0)
	outBook, outStyle = copy2(workbook)
	#workbook.save(UPLOAD_FOLDER+"/testMal"+".xls")

	# Looper through all elements except the last one (name of file)
	for i in range(len(infoList) - 1):
		xf_index = inSheet.cell_xf_index(i+1, 1)
		saved_style = outStyle[xf_index]
		outBook.get_sheet(0).write(i+1,1,infoList[i], saved_style)
		#workbook.get_sheet(0).write(i+1,1,infoList[i])

	# Set the name of the file
	outBook.save(UPLOAD_FOLDER+"/%s"%infoList[len(infoList)-1]+".xls")
	#outBook.ExportAsFixedFormat(0, UPLOAD_FOLDER+"/%s"%infoList[len(infoList)-1]+".pdf")

def editDocument(path):
	path = path
	fileInfo = []
	## OPEN workbook and from path, and store information in a list that is returned
	workbook = xlrd.open_workbook(path, formatting_info=True, on_demand=True)
	inSheet = workbook.sheet_by_index(0)

	orderNumber = getCellInfo(1,1,inSheet)
	customerGrp = getCellInfo(2,1,inSheet)
	dateDate = getCellInfo(3,1,inSheet)
	departmentTime = getCellInfo(4,1,inSheet)
	meetingPlace = getCellInfo(5,1,inSheet)
	numPers = getCellInfo(6,1,inSheet)
	assignment = getCellInfo(7,1,inSheet)
	otherInfo = getCellInfo(8,1,inSheet)
	executeBy = getCellInfo(9,1,inSheet)
	mobileNumber = getCellInfo(10,1,inSheet)
	price = getCellInfo(11,1,inSheet)

	fileInfo.append(orderNumber)
	fileInfo.append(customerGrp)
	fileInfo.append(dateDate)
	fileInfo.append(departmentTime)
	fileInfo.append(meetingPlace)
	fileInfo.append(numPers)
	fileInfo.append(assignment)
	fileInfo.append(otherInfo)
	fileInfo.append(executeBy)
	fileInfo.append(mobileNumber)
	fileInfo.append(price)

	return fileInfo

# Copy formating of excel sheet
def copy2(wb):
	w = XLWTWriter()
	process(
		XLRDReader(wb,'unknown.xls'), w)
	return w.output[0][1], w.style_list