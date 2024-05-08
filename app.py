from flask import Flask, request, render_template, send_file
import os
from pdf2docx import Converter

app = Flask(__name__)

# Define the upload folder
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Define allowed file extensions
ALLOWED_EXTENSIONS = {'pdf'}

# Function to check if file extension is allowed
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Route for the home page
@app.route('/')
def home():
    return render_template('index.html')
folder_path = 'uploads'

def delete_files_in_folder(folder_path):
    # Allowed file extensions to delete
    allowed_extensions = {'.pdf', '.docx'}

    # Iterate over all files in the folder
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                # Check if the file extension is allowed
                _, file_extension = os.path.splitext(filename)
                if file_extension.lower() in allowed_extensions:
                    # If it's a file with allowed extension, delete it
                    os.unlink(file_path)
                    print(f"Deleted {file_path}")
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")

# Route for file upload
@app.route('/upload', methods=['POST'])
def upload_file():
    delete_files_in_folder(folder_path)
    # Check if a file was uploaded
    if 'file' not in request.files:
        return render_template('index.html', error='No file part')

    file = request.files['file']

    # Check if file is empty
    if file.filename == '':
        return render_template('index.html', error='No selected file')

    # Check if file extension is allowed
    if not allowed_file(file.filename):
        return render_template('index.html', error='File type not allowed')

    # Save the file
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(file_path)

    # Convert PDF to DOCX
    docx_filename = file.filename.rsplit('.', 1)[0] + '.docx'
    docx_path = os.path.join(app.config['UPLOAD_FOLDER'], docx_filename)

    # Convert PDF to DOCX using pdf2docx
    try:
        convert = Converter(file_path)
        convert.convert(docx_path)
        convert.close()
    except Exception as e:
        return render_template('index.html', error=f'Error converting PDF to DOCX: {str(e)}')

    # Delete the uploaded PDF file
    os.remove(file_path)
    # Return the converted file for download with the correct filename
    return send_file(docx_path, as_attachment=True, download_name=docx_filename)
    

if __name__ == '__main__':
    app.run(debug=True)
