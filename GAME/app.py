
from flask import Flask, render_template, request, send_file, redirect, url_for, flash
import os
from werkzeug.utils import secure_filename
from filter_logic import filter_excel_data
import zipfile

UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'

app = Flask(__name__)
app.secret_key = 'secretkey'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    file_sv = request.files.get('file_sv')
    file_cong = request.files.get('file_cong')

    if not file_sv or not file_cong:
        flash("Vui lòng chọn cả hai file Excel")
        return redirect(url_for('index'))

    filename_sv = secure_filename(file_sv.filename)
    filename_cong = secure_filename(file_cong.filename)

    path_sv = os.path.join(app.config['UPLOAD_FOLDER'], filename_sv)
    path_cong = os.path.join(app.config['UPLOAD_FOLDER'], filename_cong)

    file_sv.save(path_sv)
    file_cong.save(path_cong)

    output_files = filter_excel_data(path_sv, path_cong, app.config['RESULT_FOLDER'])

    zip_path = os.path.join(app.config['RESULT_FOLDER'], 'ket_qua_loc.zip')
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for file in output_files:
            zipf.write(file, os.path.basename(file))

    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
