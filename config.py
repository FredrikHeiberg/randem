import os

# add utils config to base config e.g.

# Default config
class BaseConfig(object):
	DEBUG = False
	SECRET_KEY = os.urandom(24)
	BASE_DIR = os.path.abspath(os.path.dirname(__file__))
	SQLALCHEMY_DATABASE_URI = os.environ['DATABASE_URL']
	print SQLALCHEMY_DATABASE_URI


class DevelopmentConfig(BaseConfig):
	DEBUG = True

class ProductionConfig(BaseConfig):
	# Set to False!
	DEBUG = True

# Update to search for database!
# export DATABASE_URL="sqlite:///posts.db"

# sett config settings
# export APP_SETTIGS="config.DevelopmentConfig"


