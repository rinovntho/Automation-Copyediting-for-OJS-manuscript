
from flask import Flask, render_template, request, send_file
import os
import docx
from script import main

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process_files():
    # Get the uploaded files
    file1 = request.files['file1']
    file2 = request.files['file2']

    # Save the files to the upload folder
    file1_path = os.path.join(app.config['UPLOAD_FOLDER'], file1.filename)
    file2_path = os.path.join(app.config['UPLOAD_FOLDER'], file2.filename)
    file1.save(file1_path)
    file2.save(file2_path)

    # Process the files
    doawnload_file_name = main(file1_path, file2_path)
    


    for file_name in os.listdir(app.config['UPLOAD_FOLDER']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
        os.remove(file_path)

    return render_template('result.html', output_file=f'{doawnload_file_name}')

@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join(app.config['OUTPUT_FOLDER'], filename)
    response = send_file(file_path, as_attachment=True)

    return response

@app.route('/done', methods=['POST'])
def button_pressed():
    file_path = os.path.join(app.config['OUTPUT_FOLDER'])
    if request.method == 'POST':
        for files in os.listdir(file_path):
            file_full_path = os.path.join(file_path, files)
            os.remove(file_full_path)
        return render_template('index.html')

@app.route('/clear_folders', methods=['GET'])
def clear_folders():
    # Clear the uploads folder
    for file_name in os.listdir(app.config['UPLOAD_FOLDER']):
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
        os.remove(file_path)

    # Clear the output folder
    for file_name in os.listdir(app.config['OUTPUT_FOLDER']):
        file_path = os.path.join(app.config['OUTPUT_FOLDER'], file_name)
        os.remove(file_path)

    return "success"

if __name__ == '__main__':
    app.run(host='0.0.0.0', debug=True)