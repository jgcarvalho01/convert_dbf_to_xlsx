from io import BytesIO
from flask import Flask, request, redirect, url_for, render_template, send_file
import pandas as pd
from dbfread import DBF

app = Flask(__name__)

def dbf_to_xlsx_buffer(dbf_file):
    # Criar um buffer para armazenar o arquivo XLSX
    xlsx_buffer = BytesIO()
    
    # Salvar o arquivo DBF temporariamente
    temp_dbf_path = "/tmp/temp.dbf"
    with open(temp_dbf_path, 'wb') as temp_dbf_file:
        temp_dbf_file.write(dbf_file.read())
    
    # Ler o arquivo DBF
    table = DBF(temp_dbf_path, encoding='latin1')
    df = pd.DataFrame(iter(table))
    
    # Salvar o DataFrame como um arquivo XLSX no buffer
    with pd.ExcelWriter(xlsx_buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    
    xlsx_buffer.seek(0)  # Voltar ao início do buffer
    return xlsx_buffer

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            xlsx_buffer = dbf_to_xlsx_buffer(file)
            return send_file(
                xlsx_buffer,
                as_attachment=True,
                download_name=file.filename.rsplit('.', 1)[0] + '.xlsx',
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    return render_template('index.html')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in {'dbf'}

if __name__ == '__main__':
    app.run(debug=True)
