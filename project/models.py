from project import db
from project import bcrypt
from sqlalchemy import ForeignKey
from sqlalchemy.org import relationship

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
		self.password = bcrypt.generate_password_hash(password)

	def __repr__(self):
		return '<name {}'.format(self.name)