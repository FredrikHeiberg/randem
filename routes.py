# coding=utf-8
from flask import Flask, render_template, request, url_for, redirect, session, flash, g, abort, send_from_directory
from functools import wraps
from flask.ext.uploads import UploadSet, configure_uploads, DOCUMENTS
from xlrd import *
from xlutils.copy import copy
from form import LoginForm
from werkzeug import secure_filename
from utils import UPLOAD_FOLDER, TEMPLATE_FOLDER #, app, bcrypt
from flask_wtf.csrf import CsrfProtect
from flask.ext.bcrypt import Bcrypt
from datetime import date, timedelta as td
import os, glob, xlrd, datetime, re
from xlutils.filter import process,XLRDReader,XLWTWriter
#from models import User
from models import *
from flask.ext.sqlalchemy import SQLAlchemy
import logging
import sys
#import sqlite3

#
#	NOTE TO SELF: convert excel to PDF?
#			HUSK: add log and order id to .git ignore
#
#

#---------------------------------------------------------------#
# App config.
#---------------------------------------------------------------#

# create the application object
app = Flask(__name__)
bcrypt = Bcrypt(app)

# config
app.config.from_object(os.environ['APP_SETTINGS'])

# Create the sqlalchemy object
db = SQLAlchemy(app)




# import Db from models
#from models import BlogPost

CsrfProtect(app)

ALLOWED_EXTENSIONS = set(['xlsx','xls'])
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

		#print "fileName: %s, %s"%(infoList[-1], editList[-1])
		fullEditFileName = "%s.xls"%editList[-1]
		# Delete the old file
		if infoList[-1] != fullEditFileName:
			os.remove(str(UPLOAD_FOLDER)+"/%s" %infoList[-1])

		return redirect(url_for('uploadedfiles'))
	
	return render_template('edit.html', file=file, infoList=infoList)


### FIND A WAY TO REFACTOR SO THAT THIS WILL WORK -- REFERENCE BETWEEN EACHOTHER
@app.route('/login', methods=['GET', 'POST'])
def login():
	error = None

	form = LoginForm(request.form)

	if request.method == 'POST':
		if form.validate_on_submit():
			user = User.query.filter_by(name=request.form['username']).first()
			if user is not None and bcrypt.check_password_hash(
				user.password, request.form['password']):
				session['logged_in'] = True
				flash('Du er naa logget inn!')
				return redirect(url_for('index'))
			else:
				#print "Feil brukernavn/passord"
				error = 'Feil brukernavn/passord'
	return render_template('login.html', form=form, error=error)

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



#
# Excel to PDF convert - use command line unoconv -f pdf your_excel.xls - read on this (will only work on the server)
#
#
#	

#w = copy(open_workbook('/Users/fredrikheiberg/Documents/randem/static/sheets/mal.xls',formatting_info=True))


if __name__ == '__main__':
	app.run() 










