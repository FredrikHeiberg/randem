#-----------------------#
# imports				#
#-----------------------#

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

#-----------------------#
# config				#
#-----------------------#

app = Flask(__name__)
bcrypt = Bcrypt(app)
app.config.from_object(os.environ['APP_SETTINGS'])
db = SQLAlchemy(app)