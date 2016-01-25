import os
#from flask.ext.sqlalchemy import SQLAlchemy
#from flask import Flask
#from flask.ext.bcrypt import Bcrypt

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'static/sheets')
TEMPLATE_FOLDER = os.path.join(BASE_DIR, 'static/fileTemplates')

# create the application object
#app = Flask(__name__)
#bcrypt = Bcrypt(app)

# config
#app.config.from_object(os.environ['APP_SETTINGS'])

# Create the sqlalchemy object
#db = SQLAlchemy(app)