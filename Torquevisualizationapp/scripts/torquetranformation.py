import os
from flask import Flask, render_template, request, jsonify, send_from_directory
import pandas as pd
from config import REQUIREMENTS_PATH

# Function to install requirements
def install_requirements():
    if os.path.exists(REQUIREMENTS_PATH): 
        os.system(f"pip install -r {REQUIREMENTS_PATH}")
    else:
        print(f"{REQUIREMENTS_PATH} not found!")

# Calling the function to install requirements
install_requirements()

app = Flask(__name__)
app.config .from_pyfile('config.py')  

app.template_folder = app.config['TEMPLATES_FOLDER']
app.static_folder = app.config['STATIC_FOLDER']
INPUT_DIR = app.config['INPUT_DIR']

@app.route('/')
def index():
    if not os.path.exists(INPUT_DIR) or INPUT_DIR == "DEFAULT_INPUT_DIR":
        error_message = f"Directory not found: {INPUT_DIR}. Check your INPUT_DIR environment variable."
        app.logger.error(error_message)
        return render_template('error.html', error=error_message), 500

    product_folders = [d for d in os.listdir(INPUT_DIR) if os.path.isdir(os.path.join(INPUT_DIR, d))]
    return render_template('index.html', folders=product_folders)

@app.route('/get_file_data', methods=['GET'])
def get_file_data():
    file_name = request.args.get('file', '')
    folder_name = request.args.get('folder', '')
    full_path = os.path.join(INPUT_DIR, folder_name, file_name)

    try:
        df = pd.read_excel(full_path)
        data = {
            'x': df[df.columns[0]].tolist(),
            'y': df[df.columns[1]].tolist()
        }
        if len(df.columns) > 2:  # Check for the presence of a third column (Temperature for Behavior)
            data['temperature'] = df[df.columns[2]].tolist()

        return jsonify(data)

    except Exception as e:
        app.logger.error(f"Error reading file {full_path}. Error: {e}")
        return jsonify({"error": f"Failed to read the file {full_path}. Error: {str(e)}"}), 400

@app.route('/get_files_ajax', methods=['GET'])
def get_files_ajax():
    folder_name = request.args.get('folder_name', '')
    folder_path = os.path.join(INPUT_DIR, folder_name)
    files = [f for f in os.listdir(folder_path) if f.startswith('Ready_') and f.lower().endswith('.xlsx')]
    return jsonify(files)

@app.route('/get_sheet2_data', methods=['GET'])
def get_sheet2_data():
    file_name = request.args.get('file', '')
    folder_name = request.args.get('folder', '')
    full_path = os.path.join(INPUT_DIR, folder_name, file_name)

    try:
        df = pd.read_excel(full_path, sheet_name="Sheet2")
        data = {
            'parameters': df[df.columns[0]].tolist(),
            'values': df[df.columns[1]].tolist()
        }
        return jsonify(data)

    except Exception as e:
        app.logger.error(f"Error reading Sheet2 of file {full_path}. Error: {e}")
        return jsonify({"error": f"Failed to read the Sheet2 of file {full_path}. Error: {str(e)}"}), 400

@app.route('/get_logo')
def get_logo():
    return send_from_directory(app.static_folder, 'images/Fanatec_logo_black.jpg')

@app.route('/get_endor_logo')
def get_endor_logo():
    return send_from_directory(app.static_folder, 'images/Endor_and_Claim_color.jpg')

@app.errorhandler(500)
def internal_server_error(error):
    app.logger.error(f"Internal server error: {error}")
    return render_template('error.html', error="Internal Server Error"), 500

if __name__ == '__main__':
    app.run()