from flask import Flask, request, jsonify
from flask_cors import CORS
import re, os, uuid, ipaddress, requests as req

app = Flask(__name__)
# TODO: ограничить origins доменом сайта вместо «*», когда он зафиксирован:
#   CORS(app, origins=['https://63pokupki.ru'])
CORS(app)
# Жёсткий лимит на размер тела запроса — защита от раздувания/DoS.
app.config['MAX_CONTENT_LENGTH'] = 64 * 1024

MM_HOOK          = 'https://mm.63pokupki.ru:8443/hooks/k7xh9osx5tr1pbyr1b6y3rqcba'
MM_HOOK_PVZ      = 'https://mm.63pokupki.ru:8443/hooks/ycpetuzfn78u881yx793cfghpy'
MM_HOOK_SUPPLIER = 'https://mm.63pokupki.ru:8443/hooks/sds81m1gyjnw7guk8xjy9agqcw'


def notify_mm_hook(hook, text):
    try:
        req.post(hook, json={'text': text}, timeout=5)
    except Exception:
        pass


# Приём заявок поставщиков в хаб. Секретный токен — только из окружения:
#   systemd: Environment=SUPPLIER_WEBFORM_TOKEN=...  (или EnvironmentFile)
HUB_SUPPLIER_WEBHOOK   = 'https://hub.63pokupki.ru/api/v1/suppliers/sources/manual/candidates/webhook'
SUPPLIER_WEBFORM_TOKEN = os.environ.get('SUPPLIER_WEBFORM_TOKEN', '')


def send_supplier_candidate(payload):
    if not SUPPLIER_WEBFORM_TOKEN:
        print('[supplier-webhook] SKIP: SUPPLIER_WEBFORM_TOKEN не задан', flush=True)
        return
    try:
        r = req.post(HUB_SUPPLIER_WEBHOOK, json=payload,
                     headers={'X-Webhook-Token': SUPPLIER_WEBFORM_TOKEN}, timeout=5)
        if r.status_code >= 300:
            print(f'[supplier-webhook] HTTP {r.status_code}: {r.text[:500]}', flush=True)
        else:
            print(f'[supplier-webhook] OK {r.status_code}', flush=True)
    except Exception as e:
        print(f'[supplier-webhook] ERROR: {e!r}', flush=True)


def persist(kind, record):
    """Единая точка сохранения заявки.

    Хранение в .xlsx убрано по требованию — сейчас данные доставляются только
    в Mattermost. Когда появится БД, вставку делать ЗДЕСЬ и ТОЛЬКО через
    параметризованные запросы (никогда не форматировать SQL строками — иначе
    SQL-инъекция):

        cur.execute(
            "INSERT INTO leads (kind, inn, company, contact, payload, created) "
            "VALUES (%s, %s, %s, %s, %s, now())",
            (kind, record.get('inn'), record.get('company'),
             record.get('contact'), json.dumps(record, ensure_ascii=False)),
        )

    В админке значения выводить ЭКРАНИРОВАННО (html.escape / автоэкранирование
    шаблонизатора) — данные хранятся «как есть», а защита от XSS делается на
    выводе.
    """
    # placeholder: БД будет подключена позже
    pass


# ---------------------------------------------------------------------------
#  Валидация
# ---------------------------------------------------------------------------

class Invalid(Exception):
    pass


# Управляющие символы, кроме \t и \n (их обрабатываем отдельно).
_CONTROL_RE = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')
_EMAIL_RE   = re.compile(r'^[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}$')
_SITE_RE    = re.compile(r'^(https?://)?([\w\-]+\.)+[A-Za-z]{2,}(/\S*)?$')


def clean(value, maxlen, multiline=False):
    """Привести к строке, вырезать управляющие символы, ограничить длину."""
    if value is None:
        value = ''
    if not isinstance(value, str):
        raise Invalid('Некорректный формат данных')
    value = value.replace('\r\n', '\n').replace('\r', '\n')
    value = _CONTROL_RE.sub('', value).replace('\t', ' ')
    if not multiline:
        value = value.replace('\n', ' ')
    value = value.strip()
    if len(value) > maxlen:
        raise Invalid('Слишком длинное значение')
    return value


def require(value, field):
    if not value:
        raise Invalid(f'Поле «{field}» обязательно')
    return value


def valid_inn(inn):
    """Проверка ИНН (10 или 12 цифр) с контрольной суммой."""
    if not re.fullmatch(r'\d{10}|\d{12}', inn):
        return False
    d = [int(c) for c in inn]

    def csum(weights):
        return sum(w * x for w, x in zip(weights, d)) % 11 % 10

    if len(inn) == 10:
        return csum([2, 4, 10, 3, 5, 9, 4, 6, 8]) == d[9]
    n11 = csum([7, 2, 4, 10, 3, 5, 9, 4, 6, 8])
    n12 = csum([3, 7, 2, 4, 10, 3, 5, 9, 4, 6, 8])
    return n11 == d[10] and n12 == d[11]


def valid_phone(p):
    """Только российские номера: +7XXXXXXXXXX / 8XXXXXXXXXX (11 цифр, код 7/8)
    либо мобильный без кода — 9XXXXXXXXX (10 цифр)."""
    d = re.sub(r'\D', '', p)
    if len(d) == 11 and d[0] in '78':
        return True
    if len(d) == 10 and d[0] == '9':
        return True
    return False


def valid_email(e):
    return len(e) <= 254 and bool(_EMAIL_RE.match(e))


def valid_contact(c):
    return valid_phone(c) or valid_email(c)


def valid_site(s):
    return len(s) <= 200 and bool(_SITE_RE.match(s))


def valid_count(c):
    return c.isdigit() and 1 <= int(c) <= 100000


def md(value):
    """Экранирование для текста Mattermost: гасим markdown-разметку и переводы
    строк, чтобы ввод не ломал/не подделывал сообщение и не тащил картинки."""
    if not value:
        return '—'
    value = re.sub(r'([`*_~|<>\[\]\\])', r'\\\1', value)
    return value.replace('\n', ' ')


def is_spam(d):
    # honeypot: скрытое поле, которое заполняют только боты
    return bool(d.get('website'))


# ---------------------------------------------------------------------------
#  Эндпоинты
# ---------------------------------------------------------------------------

@app.route('/submit', methods=['POST'])
def submit():
    d = request.get_json(silent=True) or {}
    try:
        if is_spam(d):
            raise Invalid('Заявка отклонена')
        inn     = require(clean(d.get('inn'), 12), 'ИНН')
        company = require(clean(d.get('company'), 200), 'Компания')
        phone   = require(clean(d.get('phone'), 32), 'Телефон')
        email   = clean(d.get('email'), 254)
        product = clean(d.get('product'), 2000, multiline=True)

        if not valid_inn(inn):
            raise Invalid('Некорректный ИНН')
        if not valid_phone(phone):
            raise Invalid('Некорректный телефон')
        if email and not valid_email(email):
            raise Invalid('Некорректный e-mail')
    except Invalid as e:
        return jsonify({'error': str(e)}), 400

    record = {'inn': inn, 'company': company, 'phone': phone,
              'email': email, 'product': product}
    persist('lead', record)
    notify_mm_hook(MM_HOOK,
        f"📥 **Новая B2B заявка**\n**ИНН:** {md(inn)}\n**Компания:** {md(company)}\n"
        f"**Телефон:** {md(phone)}\n**Email:** {md(email)}\n**Комментарий:** {md(product)}")
    return jsonify({'ok': True})


@app.route('/submit-pvz', methods=['POST'])
def submit_pvz():
    d = request.get_json(silent=True) or {}
    try:
        if is_spam(d):
            raise Invalid('Заявка отклонена')
        inn     = require(clean(d.get('inn'), 12), 'ИНН')
        company = require(clean(d.get('company'), 200), 'Компания')
        address = require(clean(d.get('address'), 2000, multiline=True), 'Адреса')
        count   = require(clean(d.get('count'), 10), 'Количество ПВЗ')
        contact = require(clean(d.get('contact'), 100), 'Контакт')
        comment = clean(d.get('comment'), 2000, multiline=True)

        if not valid_inn(inn):
            raise Invalid('Некорректный ИНН')
        if not valid_count(count):
            raise Invalid('Некорректное количество ПВЗ')
        if not valid_contact(contact):
            raise Invalid('Контакт: укажите телефон или e-mail')
    except Invalid as e:
        return jsonify({'error': str(e)}), 400

    record = {'inn': inn, 'company': company, 'address': address,
              'count': count, 'contact': contact, 'comment': comment}
    persist('pvz', record)
    notify_mm_hook(MM_HOOK_PVZ,
        f"🏪 **Новая заявка ПВЗ**\n**ИНН:** {md(inn)}\n**Компания:** {md(company)}\n"
        f"**Адреса:** {md(address)}\n**Кол-во ПВЗ:** {md(count)}\n"
        f"**Контакт:** {md(contact)}\n**Комментарий:** {md(comment)}")
    return jsonify({'ok': True})


@app.route('/submit-supplier', methods=['POST'])
def submit_supplier():
    d = request.get_json(silent=True) or {}
    try:
        if is_spam(d):
            raise Invalid('Заявка отклонена')
        inn      = require(clean(d.get('inn'), 12), 'ИНН')
        company  = require(clean(d.get('company'), 200), 'Компания')
        category = require(clean(d.get('category'), 100), 'Категория')
        site     = require(clean(d.get('site'), 200), 'Сайт')
        phone    = clean(d.get('phone'), 32)
        email    = require(clean(d.get('email'), 254), 'E-mail')
        comment  = clean(d.get('comment'), 2000, multiline=True)

        if not valid_inn(inn):
            raise Invalid('Некорректный ИНН')
        if not valid_site(site):
            raise Invalid('Некорректный сайт')
        if not valid_email(email):
            raise Invalid('Некорректный e-mail')
        if phone and not valid_phone(phone):
            raise Invalid('Некорректный телефон')
    except Invalid as e:
        return jsonify({'error': str(e)}), 400

    record = {'inn': inn, 'company': company, 'category': category,
              'site': site, 'phone': phone, 'email': email, 'comment': comment}
    persist('supplier', record)

    notes_parts = [f'Категория: {category}'] if category else []
    if comment:
        notes_parts.append(comment)
    send_supplier_candidate({
        'external_id': 'web-form-' + uuid.uuid4().hex,
        'name':    company,
        'inn':     inn,
        'email':   email,
        'phone':   phone,
        'website': site,
        'notes':   '. '.join(notes_parts),
    })

    notify_mm_hook(MM_HOOK_SUPPLIER,
        f"🚚 **Новая заявка поставщика**\n**ИНН:** {md(inn)}\n**Компания:** {md(company)}\n"
        f"**Категория:** {md(category)}\n**Сайт:** {md(site)}\n"
        f"**Телефон:** {md(phone)}\n**E-mail:** {md(email)}\n**Комментарий:** {md(comment)}")
    return jsonify({'ok': True})


@app.route('/track', methods=['POST'])
def track():
    d = request.get_json(silent=True) or {}
    ip = request.headers.get('X-Forwarded-For', request.remote_addr or '').split(',')[0].strip()
    try:
        ipaddress.ip_address(ip)
    except ValueError:
        ip = ''
    record = {
        'ip':       ip,
        'ua':       clean(request.headers.get('User-Agent', ''), 500),
        'referrer': clean(d.get('referrer', ''), 500),
        'screen':   clean(d.get('screen', ''), 20),
        'language': clean(d.get('language', ''), 40),
        'user_id':  clean(d.get('user_id', ''), 100),
    }
    persist('visit', record)
    return jsonify({'ok': True})


if __name__ == '__main__':
    # Прод: запускать под WSGI (gunicorn/uwsgi) за nginx с TLS, не dev-сервером.
    app.run(port=5000)
