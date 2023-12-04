from flask import Flask, request, render_template, send_file
from werkzeug.utils import secure_filename
import xml.etree.ElementTree as ET
import os
import pandas as pd
import re

# Create Flask app
app = Flask(__name__)

# Define the upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def clean_tag_name(tag_name):
    # Replace invalid XML tag characters with underscores
    tag_name = re.sub(r'\s+', '_', tag_name)
    tag_name = re.sub(r'[^a-zA-Z0-9_]', '', tag_name)
    return tag_name

# Function to check if the file has an allowed extension
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Route to render the main HTML page
@app.route('/')
def index():
    return render_template('index.html')

# Route to handle file upload
@app.route('/upload', methods=['POST'])
def upload_file():
    # Check if the post request has the file part
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    # If the user does not select a file, browser also
    # submit an empty part without filename
    if file.filename == '':
        return "No selected file"

    # If the file is allowed, save it to the upload folder
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Process the uploaded file (you can add your conversion logic here)
        # For demonstration, we'll just read the Excel file using pandas
        df = pd.read_excel(file_path)
        # Add your conversion logic here
        # Create root element
        root = ET.Element("data")

        # Iterate through rows
        for _, row in df.iterrows():
            # Create element for each row
            row_element = ET.SubElement(root, "row")

            # Iterate through columns
            for col_name, col_value in row.items():
                # Clean column name to make a valid XML tag
                col_name = clean_tag_name(col_name)

                # Create element for each column
                col_element = ET.SubElement(row_element, col_name)
                col_element.text = str(col_value)

        # Save the XML file
        xml_filename = filename.rsplit('.', 1)[0] + '.xml'
        xml_filepath = os.path.join(app.config['UPLOAD_FOLDER'], xml_filename)
        tree = ET.ElementTree(root)
        tree.write(xml_filepath)


        # Provide a link to download the converted file (dummy link for demonstration)
        # return f"File uploaded and converted successfully! <a href='/download/{secure_filename(xml_filename)}'>Download converted file</a>"
        return  render_template('download.html', filename=xml_filename)
    else:
        return "Invalid file format"

# Route to download the converted file
@app.route('/download/<filename>')
def download_file(filename):
    # Ensure the requested file has the correct extension
    if not filename.endswith('.xml'):
        return "Invalid file format"

    # Provide the correct filepath for the XML file
    xml_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)

    # Check if the file exists
    if not os.path.exists(xml_filepath):
        return "File not found"

    return send_file(xml_filepath, as_attachment=True, download_name=f"{filename}")
