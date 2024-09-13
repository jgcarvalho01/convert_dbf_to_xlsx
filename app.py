from flask import Flask, request, redirect, url_for, render_template, send_from_directory
import os
import pandas as pd
from dbfread import DBF

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['ALLOWED_EXTENSIONS'] = {'dbf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def dbf_to_xlsx(dbf_file_path, xlsx_file_path):
    # Ler o arquivo DBF com codificação definida
    table = DBF(dbf_file_path, encoding='latin1')
    df = pd.DataFrame(iter(table))
    
    # Salvar o DataFrame como um arquivo XLSX
    df.to_excel(xlsx_file_path, index=False)


@app.route('/', methods=['GET', 'POST'])
def index():
    filename = None
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = file.filename
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)
            
            # Criar o arquivo XLSX
            xlsx_filename = filename.rsplit('.', 1)[0] + '.xlsx'
            xlsx_file_path = os.path.join(app.config['UPLOAD_FOLDER'], xlsx_filename)
            dbf_to_xlsx(file_path, xlsx_file_path)
            
            return render_template('index.html', filename=xlsx_filename)
    
    return render_template('index.html')

@app.route('/uploads/<filename>')
def download_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    if not os.path.exists(app.config['UPLOAD_FOLDER']):
        os.makedirs(app.config['UPLOAD_FOLDER'])
    app.run(debug=True)
