from flask import Flask, render_template, request, redirect, url_for
import pandas as pd
from pymongo import MongoClient
import json, io
from bson import ObjectId


app = Flask(__name__)

# MongoDB connection
client = MongoClient('mongodb://localhost:27017/')
db = client['excel_files']
collection = db['files']

class JSONEncoder(json.JSONEncoder):
    def default(self, o):
        if isinstance(o, ObjectId):
            return str(o)
        return super().default(o)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/form')
def form():
    return render_template('form.html')

@app.route('/visualization')
def visualization():
    form_data = list(collection.find())
    processed_data = []

    for data in form_data:
        column_names = data.get('column_names', [])
        row_data = data.get('row_data', [])

        processed_data.append({
            'column_names': column_names,
            'row_data': row_data
        })

    return render_template('visualization.html', form_data=processed_data)





@app.route('/process_data', methods=['POST'])
def process_data():
    # Get the form data
    form_data = request.form

    # Create a Pandas DataFrame from the form data
    df = pd.DataFrame(form_data, index=[0])

    # Convert the DataFrame to an Excel file in memory
    excel_file = io.BytesIO()
    with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

    # Save the Excel file in the database
    collection.insert_one({'filename': 'form_data.xlsx', 'data': excel_file.getvalue()})

    return redirect(url_for('upload_file'))


@app.route('/excel', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if a file was uploaded
        if 'file' not in request.files:
            return render_template('upload.html', error='No file selected')

        file = request.files['file']

        # Check if the file has a valid extension
        if file.filename == '':
            return render_template('upload.html', error='No file selected')
        if not file.filename.endswith('.xlsx'):
            return render_template('upload.html', error='Invalid file format. Only Excel files are allowed.')

        # Convert file data to bytes
        file_data = file.read()

        # Save the file in MongoDB
        collection.insert_one({'filename': file.filename, 'data': file_data})

        return redirect(url_for('upload_file'))

    # Get a list of stored files from MongoDB
    files = collection.find({}, {'filename': 1})
    return render_template('file_list.html', files=files)

@app.route('/visualize/<filename>')
def visualize_file(filename):
    # Retrieve the Excel file from MongoDB
    file = collection.find_one({'filename': filename})

    if file:
        # Load data from Excel using pandas
        data = pd.read_excel(file['data'])

        # Get column names and rows
        column_names = data.columns.tolist()
        rows = data.to_dict('records')

        return render_template('visualization.html', filename=filename, column_names=column_names, rows=rows)
    else:
        return 'File not found'

if __name__ == '__main__':
    app.run(debug=True)
