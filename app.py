from flask import Flask, render_template, request, make_response, send_file
from werkzeug.wrappers import Response
import fitz #as f
import re
import pandas as pd
import math
import io
import xlrd
import os
import openpyxl


app = Flask(__name__)

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/catalogo", methods=["GET"])
def catalogo():
    return render_template("catalogo.html")

@app.route("/conversor", methods=["POST"])
def conversor():

    def convert_to_xlsx(file_path):
        file_name, file_ext = os.path.splitext(file_path)
        if file_ext.lower() == '.xls':
            new_file_path = file_name + '.xlsx'
            workbook = openpyxl.Workbook()
            worksheet = workbook.active
        
            with xlrd.open_workbook(file_path) as xls_workbook:
                sheet = xls_workbook.sheet_by_index(0)
                for row in range(sheet.nrows):
                    for col in range(sheet.ncols):
                        worksheet.cell(row=row+1, column=col+1).value = sheet.cell_value(row, col)

            workbook.save(new_file_path)
            return new_file_path

        return file_path

    print("Solicitud recibida en /convert")
    
    # Verificar que la solicitud tenga los campos requeridos
    if "excelFile" not in request.files or "pdfFile" not in request.files:
        return "Error: Campos de archivo faltantes", 400

    excel_file = request.files["excelFile"]
    pdf_file = request.files["pdfFile"]
    ganancia = float(request.form["ganancia"])
    new_pdf_name = request.form["newPdfName"]
    
    print(f"Archivo de Excel recibido: {excel_file.filename}")
    print(f"Archivo PDF recibido: {pdf_file.filename}")
    print(f"Ganancia recibida: {ganancia}")
    print(f"Nuevo nombre del archivo PDF: {new_pdf_name}")

    temp_excel_path = "temp_excel" + os.path.splitext(excel_file.filename)[1]
    excel_file.save(temp_excel_path)

    # Convertir a .xlsx si es necesario
    temp_excel_path = convert_to_xlsx(temp_excel_path)
    
    df_precios = pd.read_excel(temp_excel_path)
    print(df_precios.head())
    print(temp_excel_path)

    df_precios = df_precios.rename(columns={df_precios.columns[0]: 'Código', df_precios.columns[2]: 'Precios'})
    df_codigo = df_precios[['Código', 'Precios']]
    
    df_codigo = df_codigo.dropna(subset=['Código', 'Precios'])
    df_codigo = df_codigo[df_codigo['Código'].str.contains('\d')]
    df_codigo = df_codigo[~df_codigo['Código'].str.contains(r'^\w$')]
    codigos_tupla = tuple(df_codigo['Código'].astype(str).tolist())
    df_codigo['Precios con ganancia'] = df_codigo['Precios'].apply(lambda x: math.ceil(x * (1 + ganancia/100) / 50) * 50)
    
    pdf_data = pdf_file.read()
    pdf_buffer = io.BytesIO(pdf_data)
    documento = fitz.open("pdf", pdf_buffer) #ACA

    regex = r'\b(?:{})\b'.format('|'.join(map(re.escape, codigos_tupla)))
    for numeroDePagina in range(len(documento)):
        pagina = documento.load_page(numeroDePagina)
        text = pagina.get_text("text")
        codigos = re.findall(regex, text)
        for codigo in codigos:
            precio_serie = df_codigo.loc[df_codigo['Código'] == codigo, 'Precios con ganancia']
            if not precio_serie.empty:
                precio = precio_serie.iloc[0]
                text_instances = pagina.search_for(codigo)
                for inst in text_instances:
                    bbox = inst.irect
                    new_y = bbox.y0 - 70
                    new_x = bbox.x0
                    pagina.insert_text((new_x, new_y), str(precio), fontsize=20, fill=(1, 0.3, 0), render_mode=2)
    
    new_pdf_buffer = io.BytesIO()
    documento.save(new_pdf_buffer)
    new_pdf_buffer.seek(0)



    print("Generando archivo PDF nuevo...")
    if os.path.exists("temp_excel.xlsx"):
        os.remove("temp_excel.xlsx")
    if os.path.exists("temp_excel.xls"):
        os.remove("temp_excel.xls")

    response = make_response(new_pdf_buffer.getvalue())
    response.headers['Content-Type'] = 'application/pdf'
    response.headers['Content-Disposition'] = f'attachment; filename={new_pdf_name}.pdf'
    response.headers['Access-Control-Expose-Headers'] = 'Content-Disposition'

    return response

app.run(host= "0.0.0.0", port= 3000) 
