# coding=utf-8
from flask import Flask, render_template, request, url_for, redirect, session, flash, g, send_from_directory #, abort
from functools import wraps
from flask.ext.uploads import UploadSet, configure_uploads, DOCUMENTS
#from xlrd import *
from xlutils.copy import copy
#from flask.ext.wtf import form
#from form import LoginForm
from werkzeug import secure_filename
from utils import UPLOAD_FOLDER, TEMPLATE_FOLDER, BACKUP_FOLDER, app, headers, params, dburl #, bcrypt
from flask_wtf.csrf import CsrfProtect
#from flask.ext.bcrypt import Bcrypt
from datetime import date, timedelta as td
import os, glob, xlrd, datetime, re
from xlutils.filter import process,XLRDReader,XLWTWriter
#from models import User

#from models import *
#from flask.ext.sqlalchemy import SQLAlchemy
import json, requests, hashlib
import logging
import sys
#import sqlite3

#
#	NOTE TO SELF: convert excel to PDF?
#			HUSK: add log and order id to .git ignore and add debug=False
#
#

#---------------------------------------------------------------#
# App config.
#---------------------------------------------------------------#

# create the application object
#app = Flask(__name__)
#bcrypt = Bcrypt(app)

# config
#app.config.from_object(os.environ['APP_SETTINGS'])

# Create the sqlalchemy object
#db = SQLAlchemy(app)




# import Db from models
#from models import BlogPost

ALLOWED_EXTENSIONS = set(['xlsx','xls'])
CsrfProtect(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['TEMPLATE_FOLDER'] = TEMPLATE_FOLDER
#excelFiles = UploadSet('excelf', DOCUMENTS)

#app.config['UPLOADED_FILES_DEST'] = '/sheets'
#configure_uploads(app, excelFiles)

# Defaults to stdout
logging.basicConfig(level=logging.INFO)

# get the logger for the current Python module
log = logging.getLogger(__name__)

try: 

    log.info('Start reading database')
    # do risky stuff

except:

    # http://docs.python.org/2/library/sys.html
    _, ex, _ = sys.exc_info()
    log.error(ex.message)

def login_required(f):
	@wraps(f)
	def wrap(*args, **kwargs):
		if 'logged_in' in session:
			return f(*args, **kwargs)
		else:
			return redirect(url_for('login'))
	return wrap



formatedList = []
listOfOrders = []
infoList = []

@app.route('/', methods=['GET', 'POST'])
@login_required
def index():
	error = None
	#g.db = connect_db()
	#cur = g.db.execute('select * from posts')
	#posts = [dict(title=row[0], description=row[1]) for row in cur.fetchall()]
	#g.db.close()

	return render_template('index.html', error=error)


	#if request.method == 'POST':
	#	if request.form['date1'] != 'admin':
	#		error = 'Feil format, bruk DD.MM.YY'
	#	else:
	#		return redirect(url_for('results'))
	#return render_template('index.html', error=error)


#@app.route('/<user>')
#def index(user=None):
#	return render_template('index.html', user=user)

@app.route('/results', methods=['GET'])
@login_required
def results():

	listOfOrders = getInfoFromExcel()
	orderListLength = len(listOfOrders)

	return render_template('results.html', listOfOrders=listOfOrders, orderListLength=orderListLength)

@app.route('/create')
@login_required
def create():
	return render_template('create.html')

@app.route('/order')
@login_required
def order():
	return render_template('order.html')

#@app.route('/upload')
#@login_required
#def upload():
#	return render_template('upload_file.html')

@app.route('/uploadedfiles', methods=['GET', 'POST'])
@login_required
def uploadedfiles():
	listOfFiles = []
	listOfFiles = getListOfSheets()

	if request.method == 'POST':
		for fileElement in formatedList:
			#print "FILE ELEMENT %s" %fileElement
			state = "%s" %fileElement in request.form

			#print state
			if state == True:
				#print "INDEX OF DELETE SHEET %s" %fileElement
				delete_item(fileElement)
		return redirect(url_for('uploadedfiles'))

		#return redirect(url_for('uploadedfiles'))
	return render_template('uploadedfiles.html', listOfFiles=listOfFiles)

#@app.route('/uploadFlask')
#def uploadFlask():
#	if request.method == 'POST' and 'photo' in request.files:
#		filename = excelFiles.save(request.files['excel'])
#		return "Fil er lastet opp"
#	return render_template('upload.html')

@app.route('/upload', methods=['GET', 'POST'])
@login_required
def upload_file():
	if request.method == 'POST':
		file = request.files['file']
		if file and allowed_file(file.filename):
			filename = secure_filename(file.filename)
			file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
			return redirect(url_for('index'))
	return render_template('upload_file.html')

@app.route('/download/<file>', methods=['GET', 'POST'])
@login_required
def download(file):
	listOfFiles = []
	listOfFiles = getListOfSheets()

	if request.method == 'GET':
		fileNamePath = UPLOAD_FOLDER+"/"+file
		return download_item(fileNamePath)

	return redirect(url_for('uploadedfiles'), file=file)

@app.route('/edit/<file>', methods=['GET', 'POST'])
@login_required
def edit(file):
	#print "FILE!!! %s" %file
	listOfFiles = []
	editList = []
	listOfFiles = getListOfSheets()

	fileNamePath = UPLOAD_FOLDER+"/"+file
	infoList = editDocument(fileNamePath)
	infoList.append(file)
	searchCondition = "template.xls"
	#print infoList
	#return editDocument(fileNamePath)

	if request.method == 'POST':
		#print "FILE NAME: %s"%file
		#print "PATH: %s/%s"%(UPLOAD_FOLDER, file)
		workbook = xlrd.open_workbook(TEMPLATE_FOLDER+"/%s"%searchCondition, formatting_info=True, on_demand=True)
		inSheet = workbook.sheet_by_index(0)
		outBook, outStyle = copy2(workbook)

		#print "ADD NEW INFO TO LIST"
		orderNumber = infoList[0]
		customerGrp = request.form['group']
		dateDate = str(request.form['dOfOrder'])
		departmentTime = str(request.form['time'])
		meetingPlace = request.form['depPlace']
		numPers = str(request.form['nPeople'])
		assignment = request.form['assignment']
		otherInfo = request.form['eInfo']
		executeBy = request.form['pBy']
		mobileNumber = str(request.form['mobileNr'])
		price = str(request.form['price'])
		sheetName = "%s-%s"%(customerGrp, orderNumber)

		editList.append(orderNumber)
		editList.append(customerGrp)
		editList.append(dateDate)
		editList.append(departmentTime)
		editList.append(meetingPlace)
		editList.append(numPers)
		editList.append(assignment)
		editList.append(otherInfo)
		editList.append(executeBy)
		editList.append(mobileNumber)
		editList.append(price)
		editList.append(sheetName)

		# Looper through all elements except the last one (name of file)
		for i in range(len(editList) - 1):
			xf_index = inSheet.cell_xf_index(i+1, 1)
			saved_style = outStyle[xf_index]
			outBook.get_sheet(0).write(i+1,1,editList[i], saved_style)

		# Set the name of the file
		outBook.save(UPLOAD_FOLDER+"/%s.xls"%sheetName)

		# Save to backup
		outBook.save(BACKUP_FOLDER+"/%s"%infoList[len(infoList)-1])

		#print "fileName: %s, %s"%(infoList[-1], editList[-1])
		fullEditFileName = "%s.xls"%editList[-1]
		# Delete the old file
		if infoList[-1] != fullEditFileName:
			os.remove(str(UPLOAD_FOLDER)+"/%s" %infoList[-1])

		return redirect(url_for('uploadedfiles'))
	
	return render_template('edit.html', file=file, infoList=infoList)



#----------------------------#
# 			restdb.io    	 #
#----------------------------#
@app.route('/login', methods=['GET', 'POST'])
def login():
	error = None

	if request.method == 'POST':

		inputUsername = request.form['username']
		inputPasswordPreHash = request.form['password']
		hash_object = hashlib.sha256(inputPasswordPreHash) # put password here
		inputPassword = hash_object.hexdigest()

		validation = checkCredentials(inputUsername, inputPassword)

		if validation != True:
			error = 'Feil brukernavn/passord'
		else:
			session['logged_in'] = True
			flash('Du er naa logget inn!')
			return redirect(url_for('index'))
	return render_template('login.html', error=error)

### FIND A WAY TO REFACTOR SO THAT THIS WILL WORK -- REFERENCE BETWEEN EACHOTHER
# @app.route('/login', methods=['GET', 'POST'])
# def login():
# 	error = None

# 	form = LoginForm(request.form)

# 	if request.method == 'POST':
# 		if form.validate_on_submit():
# 			user = User.query.filter_by(name=request.form['username']).first()
# 			if user is not None and bcrypt.check_password_hash(
# 				user.password, request.form['password']):
# 				session['logged_in'] = True
# 				flash('Du er naa logget inn!')
# 				return redirect(url_for('index'))
# 			else:
# 				#print "Feil brukernavn/passord"
# 				error = 'Feil brukernavn/passord'
# 	return render_template('login.html', form=form, error=error)

@app.route('/logout')
@login_required
def logout():
	session.pop('logged_in', None)
	flash('Du er naa logget ut!')
	return redirect(url_for('index'))

@app.route('/createFile', methods=['GET', 'POST'])
@login_required
def createFile():
	error = None
	global infoList
	infoList = []
	if request.method == 'POST':
		# Gather search field conditions and create a list of corresponding files

		templateUrd = str(TEMPLATE_FOLDER)
		staticUrl = templateUrd+"/"+"orderId.txt"
		ordernNumberLocation = open(staticUrl, 'r')

		tempOrderNumber = ordernNumberLocation.read()
		orderNumber = tempOrderNumber
		ordernNumberLocation.close()
		#print (str(request.form['group']))
		customerGrp = request.form['group']
		dateDate = str(request.form['dOfOrder'])
		departmentTime = str(request.form['time'])
		meetingPlace = request.form['depPlace']
		numPers = str(request.form['nPeople'])
		assignment = request.form['assignment']
		otherInfo = request.form['eInfo']
		executeBy = request.form['pBy']
		mobileNumber = str(request.form['mobileNr'])
		price = str(request.form['price'])
		sheetName = customerGrp+"-"+orderNumber

		infoList.append(orderNumber)
		infoList.append(customerGrp)
		infoList.append(dateDate)
		infoList.append(departmentTime)
		infoList.append(meetingPlace)
		infoList.append(numPers)
		infoList.append(assignment)
		infoList.append(otherInfo)
		infoList.append(executeBy)
		infoList.append(mobileNumber)
		infoList.append(price)
		infoList.append(sheetName)

		createDocument()

		# Add document to log
		logFileDir = templateUrd+"/"+"orderLog.txt"
		logFile = open(logFileDir, 'a')
		#logFile.write("Fil navn: \t\t%s.xls \n \tOrder nummer: \t%s\n \tKunde: \t\t%s \n\tDato: \t\t%s\n" %(sheetName, orderNumber, customerGrp, dateDate))
		logFile.close()

		# Update order id number
		numberIdValue = int(tempOrderNumber)
		newIdNumber = numberIdValue + 1

		updateId = open(staticUrl, 'w')
		updateId.write(str(newIdNumber))
		updateId.close()
		return render_template('index.html')

	return render_template('createFile.html', error=error)

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

		#print "TEST CELL VALUE %s" %getCellInfo(3,1,sheet)
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

	# Save to backup
	outBook.save(BACKUP_FOLDER+"/%s"%infoList[len(infoList)-1]+".xls")
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


def checkCredentials(inputUsername, inputPassword):
	req = requests.get(dburl, params=params, headers=headers)
	listOfCredentials = []

	for cred in req.json():
		stringValues = '{},{}'.format(cred['username'], cred['password'])
		tempList = stringValues.split(",")
		listOfCredentials.append(tempList)

	for pair in listOfCredentials:
		if (inputUsername == pair[0] and inputPassword == pair[1]):
			return True
		else:
			return False

#
# Excel to PDF convert - use command line unoconv -f pdf your_excel.xls - read on this (will only work on the server)
#
#
#	

#w = copy(open_workbook('/Users/fredrikheiberg/Documents/randem/static/sheets/mal.xls',formatting_info=True))


if __name__ == '__main__':
	app.run() 










