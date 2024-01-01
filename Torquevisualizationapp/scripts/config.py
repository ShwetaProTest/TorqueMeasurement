import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # This sets BASE_DIR to the directory where config.py is located
REQUIREMENTS_PATH = os.path.join(BASE_DIR, 'requirements.txt')
SECRET_KEY = os.environ.get('SECRET_KEY', default="DEFAULT_SECRET_KEY")
STATIC_FOLDER = os.path.join(BASE_DIR, 'app', 'static')
TEMPLATES_FOLDER = os.path.join(BASE_DIR, 'app', 'templates')
INPUT_DIR = os.path.join(BASE_DIR, 'Ready_Files')