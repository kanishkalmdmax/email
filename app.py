import openpyxl
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from openpyxl.styles import Alignment, Border, Side, PatternFill

# Flask imports
from flask import Flask, request, redirect, send_file, render_template
from werkzeug.utils import secure_filename
import os

# Initialize Flask app
app = Flask(__name__)

# Configure upload folder for Flask app
UPLOAD_FOLDER = os.path.join(app.root_path, 'uploads')
if not os.path.isdir(UPLOAD_FOLDER):
    os.mkdir(UPLOAD_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Configure download folder for Flask app
DOWNLOAD_FOLDER = os.path.join(app.root_path, 'downloads')
if not os.path.isdir(DOWNLOAD_FOLDER):
    os.mkdir(DOWNLOAD_FOLDER)
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

# Define allowed file types for upload
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

# Define border style
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Define header fill
header_fill = PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')
