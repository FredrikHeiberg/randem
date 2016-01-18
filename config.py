import os

# add utils config to base config e.g.

# Default config
class BaseConfig(object):
	DEBUG = False
	SECRET_KEY = os.urandom(24)
	# Set correct URI here - see part 9  + part 10 around 11 min (export DATABASE_URL)
	#SQLALCHEMY_DATABASE_URI = os.environ['DATABASE_URI']


class DevelopmentConfig(BaseConfig):
	DEBUG = True

class ProductionConfig(BaseConfig):
	DEBUG = False
