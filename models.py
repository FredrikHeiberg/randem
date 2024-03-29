from sqlalchemy import ForeignKey
from sqlalchemy.orm import relationship
#from routes import db, bcrypt
from utils import db #, bcrypt

class BlogPost(db.Model):

	__tablename__ = "posts"

	id = db.Column(db.Integer, primary_key=True)
	title = db.Column(db.String, nullable=False)
	description = db.Column(db.String, nullable=False)
	author_id = db.Column(db.Integer, ForeignKey("users.id"))

	def __init__(self, title, description):
		self.title = title
		self.description = description

	def __repr__(self):
		return '<title {}'.format(self.title)

class User(db.Model):

	__tablename__ = "users"

	id = db.Column(db.Integer, primary_key=True)
	name = db.Column(db.String, nullable=False)
	email = db.Column(db.String, nullable=False)
	password = db.Column(db.String, nullable=False)
	posts = relationship("BlogPost", backref="author")

	def __init__(self, name, email, password):
		self.name = name
		self.email = email
		#self.password = password
		#self.password = bcrypt.generate_password_hash(password)

	def __repr__(self):
		return '<name {}'.format(self.name)
		