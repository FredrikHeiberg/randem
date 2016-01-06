# coding=utf-8
from flask import Flask, render_template, request, url_for, redirect, session, flash, g, abort 
from functools import wraps
from flask.ext.uploads import UploadSet, configure_uploads, DOCUMENTS
from werkzeug import secure_filename
from utils import UPLOAD_FOLDER
from flask_wtf.csrf import CsrfProtect
import os, glob, xlrd
import sqlite3


app = Flask(__name__)

app.secret_key = "p1sJ24AcT9"
app.database = "sample.db"

formatedList = []

ALLOWED_EXTENSIONS = set(['xlsx'])
CsrfProtect(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
#excelFiles = UploadSet('excelf', DOCUMENTS)

#app.config['UPLOADED_FILES_DEST'] = '/sheets'
#configure_uploads(app, excelFiles)

def login_required(f):
	@wraps(f)
	def wrap(*args, **kwargs):
		if 'logged_in' in session:
			return f(*args, **kwargs)
		else:
			return redirect(url_for('login'))
	return wrap

@app.route('/', methods=['GET', 'POST'])
@login_required
def index():
	error = None
	print "TEST!"
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

@app.route('/login', methods=['GET', 'POST'])
def login():
	error = None

	if request.method == 'POST':
		if request.form['username'] != 'admin' or request.form['password'] != 'admin':
			error = 'Feil brukernavn/passord'
		else:
			session['logged_in'] = True
			flash('Du er naa logget inn!')
			return redirect(url_for('index'))
	return render_template('login.html', error=error)

@app.route('/logout')
@login_required
def logout():
	session.pop('logged_in', None)
	flash('Du er naa logget ut!')
	return redirect(url_for('index'))
#	if validate == False:
#		return redirect(url_for('login'))
#	error=None
#	global validate
#	validate = False
#	return redirect(url_for('index'))

def getInfoFromExcel():
	listOfOrders = []

	# Set directory where the sheets are 
	os.chdir(str(UPLOAD_FOLDER))

	# Gather search field conditions and create a list of corresponding files
	searchCondition = str(request.args.get('date'))
	fileList = glob.glob('%s-*.xlsx' %searchCondition)

	# Iterate through all corresponding files and create a list with relevant 
	# information from each sheet (List in list - a list of all sheets that 
	# have a list of relevant information)
	for sh in fileList:
		tempFileLocation = str(UPLOAD_FOLDER)
		file_location = tempFileLocation + "/%s" %sh
		orderDetails = []
	
		workbook = xlrd.open_workbook(file_location)
		sheet = workbook.sheet_by_index(0)

		# Dato, ank tid, sted, oppdrag, kunde, buss, mobil
		# TODO! add functionality if field is not filled in!!!!!

		# Date
		cellValue = getCellInfo(4,1,sheet)
		cellDateValue = xlrd.xldate_as_tuple(cellValue, workbook.datemode)
		dateString = str(cellDateValue[2])+"."+str(cellDateValue[1])+"."+str(cellDateValue[0])
		orderDetails.append(dateString)

		# Arrival time
		timeOfOrder = getCellInfo(5,1,sheet)
		orderDetails.append(timeOfOrder)

		# Place
		place = getCellInfo(6,1,sheet)
		orderDetails.append(place)

		# Order description
		orderDescription = getCellInfo(8,1,sheet)
		orderDetails.append(orderDescription)

		# Customer
		customer = getCellInfo(2,1,sheet)
		orderDetails.append(customer)

		# Number of People
		tempNumberOfPeople = getCellInfo(7,1,sheet)
		numberOfPeople = str(tempNumberOfPeople).split(".")
		orderDetails.append(numberOfPeople[0])

		# Buss
		buss = sheet.cell_value(10,1)
		orderDetails.append(buss)

		# Telephone number
		cellNumber = sheet.cell_value(11,1)
		orderDetails.append(cellNumber)

		listOfOrders.append(orderDetails)

		sortedList = listOfOrders.sort();
	return listOfOrders

def getCellInfo(row,col,sheet):
	if sheet.cell_value(row,col) != "":
		return sheet.cell_value(row,col)
	else:
		return "Ikke spesifisert"

def allowed_file(filename):
	return '.' in filename and \
			filename.rsplit('.', 1)[1] in ALLOWED_EXTENSIONS

def getListOfSheets():
	global formatedList;
	#formatedList = glob.glob('sheets/*.xlsx')
	formatedList = os.listdir(str(UPLOAD_FOLDER))
	listOfUploadedFiles = []
	for element in formatedList:
		#tempString = element.split("/")
		if element.endswith('.xlsx'):
			listOfUploadedFiles.append(element)
		#listOfUploadedFiles.append(tempString[1])
	return listOfUploadedFiles

def delete_item(item_id):
    #new_id = item_id
    #item = self.session.query(Item).get(item_id)
    #os.remove(os.path.join(app.config['UPLOAD_FOLDER'], item.filename))
    #self.session.delete(item)
    #db.session.commit()
    #print "DETTE SKAL SLETTES %s" %item_id
    os.remove(str(UPLOAD_FOLDER)+"/%s" %item_id)

def connect_db():
	return sqlite3.connect(app.database)

if __name__ == '__main__':
	app.run(debug=True) 










