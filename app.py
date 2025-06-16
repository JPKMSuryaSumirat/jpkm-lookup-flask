from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)

excel_data = None

@app.route('/', methods=['GET', 'POST'])
def index():
    global excel_data
    result = None
    error = None

    if request.method == 'POST':
        if 'excel_file' in request.files:
            file = request.files['excel_file']
            if file.filename.endswith(('.xls', '.xlsx')):
                try:
                    excel_data = pd.read_excel(file)
                except Exception as e:
                    error = f"Terjadi kesalahan saat membaca file: {e}"
            else:
                error = "Format file harus .xls atau .xlsx"
        elif 'participant_number' in request.form:
            if excel_data is not None:
                participant_number = request.form['participant_number']
                match = excel_data[excel_data.astype(str).apply(lambda row: participant_number in row.to_string(), axis=1)]
                if not match.empty:
                    result = match.to_dict(orient='records')
                else:
                    error = "Nomor peserta tidak ditemukan"
            else:
                error = "Silakan upload file Excel terlebih dahulu"

    return render_template('index.html', result=result, error=error)

if __name__ == '__main__':
    app.run(debug=True)
