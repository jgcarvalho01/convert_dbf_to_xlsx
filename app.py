from flask import Flask, request, redirect, url_for, render_template, Response
import io
import pandas as pd
from dbfread import DBF
from openpyxl import Workbook

app = Flask(__name__)

app.config['ALLOWED_EXTENSIONS'] = {'dbf'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def dbf_to_xlsx_buffer(dbf_file):
    # Ler o arquivo DBF com codificação definida
    table = DBF(dbf_file, encoding='latin1')
    df = pd.DataFrame(iter(table))
    
    # Criar um buffer em memória para o arquivo XLSX
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            # Converter DBF para XLSX e obter o buffer de memória
            xlsx_buffer = dbf_to_xlsx_buffer(file)
            
            # Preparar a resposta para o download
            return Response(
                xlsx_buffer,
                mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                headers={"Content-Disposition": "attachment;filename=converted.xlsx"}
            )
    
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
