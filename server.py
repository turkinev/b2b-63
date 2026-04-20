from flask import Flask, request, jsonify
from flask_cors import CORS
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os, requests as req

app = Flask(__name__)
CORS(app)

FILE        = 'заявки.xlsx'
VISITS_FILE = 'посетители.xlsx'

def get_wb(path, headers):
    if os.path.exists(path):
        return load_workbook(path)
    wb = Workbook()
    wb.active.append(headers)
    return wb

@app.route('/submit', methods=['POST'])
def submit():
    d = request.get_json()
    wb = get_wb(FILE, ['Дата', 'Компания', 'Контакт', 'ИНН', 'Комментарий'])
    wb.active.append([datetime.now().strftime('%d.%m.%Y %H:%M'), d['company'], d['contact'], d.get('inn', ''), d.get('product', '')])
    wb.save(FILE)
    return jsonify({'ok': True})

@app.route('/track', methods=['POST'])
def track():
    d = request.get_json()
    ip = request.headers.get('X-Forwarded-For', request.remote_addr).split(',')[0].strip()

    country, city = '', ''
    try:
        geo = req.get(f'http://ip-api.com/json/{ip}?lang=ru&fields=country,city', timeout=2).json()
        country = geo.get('country', '')
        city    = geo.get('city', '')
    except:
        pass

    ua = request.headers.get('User-Agent', '')
    row = [
        datetime.now().strftime('%d.%m.%Y %H:%M:%S'),
        ip, country, city, ua,
        d.get('referrer', ''),
        d.get('screen', ''),
        d.get('language', ''),
        d.get('user_id', ''),
    ]
    headers = ['Дата/Время', 'IP', 'Страна', 'Город', 'Браузер (User-Agent)', 'Источник (Referrer)', 'Экран', 'Язык', 'User ID']
    wb = get_wb(VISITS_FILE, headers)
    wb.active.append(row)
    wb.save(VISITS_FILE)
    return jsonify({'ok': True})

app.run(port=5000)
