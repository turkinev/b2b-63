from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os

app = Flask(__name__)
CORS(app)

FILE = 'заявки.xlsx'

def get_wb():
    if os.path.exists(FILE):
        return load_workbook(FILE)
    wb = Workbook()
    wb.active.append(['Дата', 'Компания', 'Контакт', 'Товар'])
    return wb

@app.route('/submit', methods=['POST'])
def submit():
    d = request.get_json()
    wb = get_wb()
    wb.active.append([datetime.now().strftime('%d.%m.%Y %H:%M'), d['company'], d['contact'], d.get('product', '')])
    wb.save(FILE)
    return jsonify({'ok': True})

app.run(port=5000)
