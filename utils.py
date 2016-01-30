import os, requests, json
from flask.ext.sqlalchemy import SQLAlchemy
from flask import Flask
#from flask.ext.bcrypt import Bcrypt

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static/sheets')
BACKUP_FOLDER = os.path.join(BASE_DIR, 'static/backup')
TEMPLATE_FOLDER = os.path.join(BASE_DIR, 'static/fileTemplates')

# create the application object
app = Flask(__name__)
#bcrypt = Bcrypt(app)

# config
#app.config.from_object(os.environ['APP_SETTINGS'])

# Create the sqlalchemy object
db = SQLAlchemy(app)

dburl = 'https://obp-randem.restdb.io/rest/brukere'
headers = {'x-apikey': 'd94ff3edb06800d3d74035321bdd03fa2d423', 'Content-Type': 'application/json'}
params = {'sort': 'title'}
