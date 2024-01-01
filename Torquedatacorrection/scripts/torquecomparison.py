from flask import Flask, render_template, request, jsonify
import os
import pandas as pd
import plotly.graph_objs as go
from plotly.offline import plot
from flask import send_from_directory

app = Flask(__name__, template_folder=os.path.join(os.path.dirname(os.path.realpath(__file__)), '..', 'templates'))
app.secret_key = 'Fanatec@123'

BASE_DIR = "C:\\Users\\Shweta\\PycharmProjects\\torquecomparisonproducts"

@app.route('/')
def index():
    product_folders = [d for d in os.listdir(BASE_DIR) if os.path.isdir(os.path.join(BASE_DIR, d))]
    return render_template('index.html', folders=product_folders)

@app.route('/get_file_data', methods=['GET'])
def get_file_data():
    file_name = request.args.get('file', '')
    folder_name = request.args.get('folder', '')
    full_path = os.path.join(BASE_DIR, folder_name, file_name)
    
    try:
        df = pd.read_excel(full_path)
        data = {
            'x': df[df.columns[0]].tolist(),
            'y': df[df.columns[1]].tolist(),
            'name': file_name
        }
        return jsonify(data)
    except Exception as e:
        print(f"Error reading file {full_path}. Error: {e}")
        return jsonify({"error": f"Failed to read the file {full_path}. Error: {str(e)}"}), 400

@app.route('/get_files_ajax', methods=['GET'])
def get_files_ajax():
    folder_name = request.args.get('folder_name', '')
    folder_path = os.path.join(BASE_DIR, folder_name)
    files = [f for f in os.listdir(folder_path) if f.startswith('Combined_') and f.lower().endswith('.xlsx')]
    return jsonify(files)

@app.route('/get_sheet2_data', methods=['GET'])
def get_sheet2_data():
    file_name = request.args.get('file', '')
    folder_name = request.args.get('folder', '')
    full_path = os.path.join(BASE_DIR, folder_name, file_name)
    
    try:
        df = pd.read_excel(full_path, sheet_name="Sheet2")  # Assuming the sheet name is "Sheet2"
        data = {
            'parameters': df[df.columns[0]].tolist(),
            'values': df[df.columns[1]].tolist()
        }
        return jsonify(data)
    except Exception as e:
        print(f"Error reading Sheet2 of file {full_path}. Error: {e}")
        return jsonify({"error": f"Failed to read the Sheet2 of file {full_path}. Error: {str(e)}"}), 400

@app.route('/get_logo')
def get_logo():
    return send_from_directory('C:\\Users\\Shweta\\PycharmProjects\\torquecomparison\\scripts', 'Fanatec_logo_black.jpg')

@app.route('/get_endor_logo')
def get_endor_logo():
    return send_from_directory('C:\\Users\\Shweta\\PycharmProjects\\torquecomparison\\scripts', 'Endor_and_Claim_color.jpg')

if __name__ == '__main__':
    app.run(debug=True)