from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, send_from_directory, jsonify
from werkzeug.utils import secure_filename
from functools import wraps
import sqlite3
import json
import sys
import subprocess
import os
import base64
from datetime import datetime, timedelta
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
import requests
from google import genai

# ============== اتصال قاعدة البيانات ==============
try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
    HAS_POSTGRES = True
except ImportError:
    HAS_POSTGRES = False
    print("⚠ psycopg2 غير مثبت. سيتم استخدام SQLite فقط.")

app = Flask(__name__)
app.secret_key = 'qamh_fleet_secret_2026'

DATA_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(DATA_DIR, 'qamh_fleet.db')
UPLOAD_FOLDER = os.path.join(DATA_DIR, 'uploads')
PASSWORDS_FILE = os.path.join(DATA_DIR, 'passwords.json')
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'gif', 'pdf'}
DATABASE_URL = os.environ.get('DATABASE_URL')

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

TELEGRAM_TOKEN = os.environ.get('TELEGRAM_TOKEN', 'ضع_توكن_البوت_هنا')
GEMINI_API_KEY = os.environ.get('GEMINI_API_KEY', 'ضع_مفتاح_جيميني_هنا')

SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_SECRET_KEY = os.environ.get('SUPABASE_SECRET_KEY', '')
SUPABASE_BUCKET = 'qamh-images'

def upload_image_to_supabase(filename, img_data):
    if not SUPABASE_URL or not SUPABASE_SECRET_KEY: return None
    try:
        url = f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_BUCKET}/{filename}"
        headers = {'Authorization': f'Bearer {SUPABASE_SECRET_KEY}', 'Content-Type': 'image/jpeg', 'x-upsert': 'true'}
        response = requests.post(url, headers=headers, data=img_data, timeout=30)
        return filename if response.status_code in (200, 201) else None
    except Exception as e: print(f"⚠ خطأ في رفع الصورة لـ Supabase: {e}"); return None

def delete_image_from_supabase(filename):
    if not filename or not SUPABASE_URL or not SUPABASE_SECRET_KEY: return False
    try:
        url = f"{SUPABASE_URL}/storage/v1/object/{SUPABASE_BUCKET}/{filename}"
        headers = {'Authorization': f'Bearer {SUPABASE_SECRET_KEY}'}
        response = requests.delete(url, headers=headers, timeout=30)
        return response.status_code in (200, 204)
    except Exception: return False

def load_passwords():
    if not os.path.exists(PASSWORDS_FILE):
        default = {'admin_password': 'admin2026', 'viewer_password': 'qamh2026'}
        with open(PASSWORDS_FILE, 'w', encoding='utf-8') as f: json.dump(default, f, ensure_ascii=False, indent=4)
        return default
    try:
        with open(PASSWORDS_FILE, 'r', encoding='utf-8') as f: return json.load(f)
    except Exception: return {'admin_password': 'admin2026', 'viewer_password': 'qamh2026'}

def save_passwords(data):
    with open(PASSWORDS_FILE, 'w', encoding='utf-8') as f: json.dump(data, f, ensure_ascii=False, indent=4)

def get_admin_password(): return load_passwords().get('admin_password', 'admin2026')
def get_viewer_password(): return load_passwords().get('viewer_password', 'qamh2026')
def allowed_file(filename): return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_connection():
    if DATABASE_URL: return psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
    else:
        conn = sqlite3.connect(DB_PATH); conn.row_factory = sqlite3.Row; return conn

def is_postgres(): return DATABASE_URL is not None

def execute_query(cursor, query, params=()):
    if is_postgres(): query = query.replace('?', '%s')
    cursor.execute(query, params); return cursor

def is_admin(): return session.get('role') == 'admin'
def is_viewer(): return session.get('role') in ('admin', 'viewer')

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not is_viewer(): return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function

def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not is_admin():
            flash('هذا الإجراء متاح للمسؤول فقط', 'error'); return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function

def get_car_current_km(car_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT MAX(end_km) as max_km FROM trips WHERE car_id=? AND end_km IS NOT NULL", (car_id,)); row = cursor.fetchone(); trip_km = row['max_km'] if row else None
    execute_query(cursor, "SELECT MAX(km_at_service) as max_km FROM maintenance WHERE car_id=? AND km_at_service IS NOT NULL", (car_id,)); row = cursor.fetchone(); maint_km = row['max_km'] if row else None
    execute_query(cursor, "SELECT MAX(km_at_fill) as max_km FROM fuel WHERE car_id=? AND km_at_fill IS NOT NULL", (car_id,)); row = cursor.fetchone(); fuel_km = row['max_km'] if row else None
    conn.close()
    candidates = [x for x in [trip_km, maint_km, fuel_km] if x is not None]
    return max(candidates) if candidates else 0

def get_maintenance_alerts():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT m.*, c.name as car_name FROM maintenance m JOIN cars c ON m.car_id = c.id WHERE m.next_service_km IS NOT NULL AND m.category = 'دورية' ORDER BY m.car_id, m.km_at_service DESC")
    all_maint = cursor.fetchall(); conn.close()
    latest_per_car = {}
    for m in all_maint:
        if m['car_id'] not in latest_per_car: latest_per_car[m['car_id']] = m
    alerts = []
    for car_id, m in latest_per_car.items():
        current_km = get_car_current_km(car_id); next_km = m['next_service_km']; remaining = next_km - current_km
        if remaining <= 0: status = 'red'; message = f"تجاوزت الصيانة بـ {abs(remaining):.0f} كم"
        elif remaining <= 500: status = 'yellow'; message = f"باقي {remaining:.0f} كم للصيانة"
        else: continue
        alerts.append({'car_name': m['car_name'], 'car_id': car_id, 'type': m['type'], 'current_km': current_km, 'next_km': next_km, 'remaining': remaining, 'status': status, 'message': message, 'last_date': m['date']})
    alerts.sort(key=lambda x: x['remaining']); return alerts

def should_count_cost(f):
    if f['source'] == 'خزان الشركة': return False
    ps = f['payment_status'] or ''
    return ps in ('مدفوعه من قسم الحركه', 'مدفوعة من قسم الحركة')

def get_effective_cost(f):
    if not should_count_cost(f): return 0
    return f['final_cost'] if f['final_cost'] else (f['total_cost'] or 0)

def get_estimated_cost(f):
    if f['final_cost']: return f['final_cost']
    if f['total_cost']: return f['total_cost']
    if f['price_per_liter'] and f['liters']:
        try: return float(f['price_per_liter']) * float(f['liters'])
        except: return 0
    return 0

def get_car_cost_per_km(car_id, conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor(); execute_query(cursor, "SELECT * FROM fuel WHERE car_id=? AND km_at_fill IS NOT NULL ORDER BY km_at_fill ASC, fuel_date ASC", (car_id,)); fills = cursor.fetchall()
    if close_conn: conn.close()
    total_km = 0; total_cost = 0
    for i in range(1, len(fills)):
        prev_km = fills[i-1]['km_at_fill']; curr_km = fills[i]['km_at_fill']
        if prev_km is None or curr_km is None: continue
        period_km = curr_km - prev_km
        if period_km <= 0: continue
        total_km += period_km; total_cost += get_estimated_cost(fills[i])
    return total_cost / total_km if total_km > 0 else 0

def get_car_liters_per_100km(car_id, conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor(); execute_query(cursor, "SELECT * FROM fuel WHERE car_id=? AND km_at_fill IS NOT NULL ORDER BY km_at_fill ASC, fuel_date ASC", (car_id,)); fills = cursor.fetchall()
    if close_conn: conn.close()
    total_km = 0; total_liters = 0
    for i in range(1, len(fills)):
        prev_km = fills[i-1]['km_at_fill']; curr_km = fills[i]['km_at_fill']
        if prev_km is None or curr_km is None: continue
        period_km = curr_km - prev_km
        if period_km <= 0: continue
        total_km += period_km; total_liters += fills[i]['liters'] or 0
    return (total_liters / total_km) * 100 if total_km > 0 else 0

def get_last_station_price_per_liter(car_id, conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor(); execute_query(cursor, "SELECT fuel_type FROM cars WHERE id = ?", (car_id,)); car = cursor.fetchone()
    car_fuel_type = car['fuel_type'] if car else None
    if not car_fuel_type:
        if close_conn: conn.close()
        return None
    execute_query(cursor, "SELECT price_per_liter FROM fuel WHERE car_id = ? AND source = 'كازية' AND fuel_type = ? AND price_per_liter IS NOT NULL AND price_per_liter > 0 ORDER BY fuel_date DESC, id DESC LIMIT 1", (car_id, car_fuel_type)); row = cursor.fetchone()
    price = None
    if row and row['price_per_liter']:
        try: price = float(row['price_per_liter'])
        except: pass
    if price is None:
        execute_query(cursor, "SELECT price_per_liter FROM fuel WHERE source = 'كازية' AND fuel_type = ? AND price_per_liter IS NOT NULL AND price_per_liter > 0 ORDER BY fuel_date DESC, id DESC LIMIT 1", (car_fuel_type,)); row = cursor.fetchone()
        if row and row['price_per_liter']:
            try: price = float(row['price_per_liter'])
            except: pass
    if close_conn: conn.close()
    return price

def get_car_estimated_cost_per_km(car_id, conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor(); execute_query(cursor, "SELECT expected_consumption FROM cars WHERE id = ?", (car_id,)); car = cursor.fetchone(); expected = car['expected_consumption'] if car else None
    if not expected:
        if close_conn: conn.close()
        return None
    try: expected = float(expected)
    except:
        if close_conn: conn.close()
        return None
    if expected <= 0:
        if close_conn: conn.close()
        return None
    price = get_last_station_price_per_liter(car_id, conn)
    if price is None:
        if close_conn: conn.close()
        return None
    if close_conn: conn.close()
    return (expected * price) / 100

def get_requesting_parties(active_only=True, conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor()
    if active_only: execute_query(cursor, "SELECT * FROM requesting_parties WHERE active=1 ORDER BY name")
    else: execute_query(cursor, "SELECT * FROM requesting_parties ORDER BY name")
    parties = cursor.fetchall()
    if close_conn: conn.close()
    return parties

def get_last_daily_check(car_id, conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor(); execute_query(cursor, "SELECT dc.*, d.name as driver_name FROM daily_checks dc JOIN drivers d ON dc.driver_id = d.id WHERE dc.car_id = ? ORDER BY dc.id DESC LIMIT 1", (car_id,)); check = cursor.fetchone()
    if close_conn: conn.close()
    if not check: return None
    try: check_datetime = datetime.strptime(f"{check['check_date']} {check['check_time']}", "%Y-%m-%d %H:%M:%S")
    except:
        try: check_datetime = datetime.strptime(f"{check['check_date']} {check['check_time']}", "%Y-%m-%d %H:%M")
        except: return None
    hours_since = (datetime.now() - check_datetime).total_seconds() / 3600
    if hours_since < 3: return dict(check)
    return None

def needs_daily_check(car_id, conn=None): return get_last_daily_check(car_id, conn) is None

def get_busy_driver_ids(conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor(); execute_query(cursor, "SELECT DISTINCT driver_id FROM trips WHERE end_km IS NULL AND trip_source = 'driver_app'"); rows = cursor.fetchall()
    if close_conn: conn.close()
    return {row['driver_id'] for row in rows}

def get_busy_car_ids(conn=None):
    close_conn = False
    if conn is None: conn = get_connection(); close_conn = True
    cursor = conn.cursor(); execute_query(cursor, "SELECT DISTINCT car_id FROM trips WHERE end_km IS NULL AND trip_source = 'driver_app'"); rows = cursor.fetchall()
    if close_conn: conn.close()
    return {row['car_id'] for row in rows}
    # ============== دوال إنشاء الجداول ==============
def ensure_all_tables():
    try:
        conn = get_connection(); cursor = conn.cursor()
        if is_postgres():
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS cars (id SERIAL PRIMARY KEY, name TEXT NOT NULL, plate_number TEXT, car_type TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            execute_query(cursor, "ALTER TABLE cars ADD COLUMN IF NOT EXISTS expected_consumption REAL DEFAULT 0")
            execute_query(cursor, "ALTER TABLE cars ADD COLUMN IF NOT EXISTS fuel_type TEXT DEFAULT ''")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS drivers (id SERIAL PRIMARY KEY, name TEXT NOT NULL, phone TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS requesting_parties (id SERIAL PRIMARY KEY, name TEXT NOT NULL, notes TEXT, active INTEGER DEFAULT 1)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS trips (id SERIAL PRIMARY KEY, trip_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER NOT NULL, destination TEXT, start_location TEXT, trip_type TEXT, requester TEXT, requesting_party_id INTEGER, exit_time TEXT, return_time TEXT, start_km REAL, end_km REAL, distance REAL, notes TEXT, start_image TEXT, end_image TEXT, start_time TEXT, end_time TEXT, trip_source TEXT DEFAULT 'manual', driver_confirmed INTEGER DEFAULT 0, admin_seen INTEGER DEFAULT 0, admin_seen_at TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS fuel (id SERIAL PRIMARY KEY, fuel_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER, fuel_type TEXT, source TEXT, payment_status TEXT, liters REAL, price_per_liter REAL, total_cost REAL, final_cost REAL, km_at_fill REAL, notes TEXT, car_image TEXT, station_image TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS maintenance (id SERIAL PRIMARY KEY, date TEXT NOT NULL, car_id INTEGER NOT NULL, type TEXT, category TEXT, km_at_service REAL, cost REAL, invoice_number TEXT, invoice_image TEXT, workshop TEXT, notes TEXT, next_service_km REAL, next_service_date TEXT, currency TEXT DEFAULT 'ل.س')""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS travel_missions (id SERIAL PRIMARY KEY, date TEXT NOT NULL, driver_id INTEGER NOT NULL, description TEXT, distance_km REAL DEFAULT 0, is_travel INTEGER DEFAULT 0, departure_time TEXT, return_time TEXT, notes TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS overtime (id SERIAL PRIMARY KEY, date TEXT NOT NULL, driver_id INTEGER NOT NULL, morning_hours REAL DEFAULT 0, evening_hours REAL DEFAULT 0, notes TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS daily_checks (id SERIAL PRIMARY KEY, check_date TEXT NOT NULL, check_time TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER NOT NULL, km_at_check REAL, oil_checked INTEGER DEFAULT 0, radiator_checked INTEGER DEFAULT 0, tires_checked INTEGER DEFAULT 0, battery_checked INTEGER DEFAULT 0, belts_checked INTEGER DEFAULT 0, oil_notes TEXT, radiator_notes TEXT, tires_notes TEXT, battery_notes TEXT, belts_notes TEXT, oil_image TEXT, radiator_image TEXT, tires_image TEXT, battery_image TEXT, belts_image TEXT, general_notes TEXT, admin_notes TEXT, has_issue INTEGER DEFAULT 0, created_at TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS pending_fuel (id SERIAL PRIMARY KEY, driver_id INTEGER, car_id INTEGER, fuel_date TEXT, fuel_time TEXT, liters REAL, price_per_liter REAL, total_cost REAL, final_cost REAL, km_at_fill REAL, station_image TEXT, car_image TEXT, notes TEXT, status TEXT DEFAULT 'pending', created_at TEXT)""")
        else:
            cursor.execute("""CREATE TABLE IF NOT EXISTS cars (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, plate_number TEXT, car_type TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            cursor.execute("PRAGMA table_info(cars)"); cars_cols = {r[1] for r in cursor.fetchall()}
            if 'expected_consumption' not in cars_cols: cursor.execute("ALTER TABLE cars ADD COLUMN expected_consumption REAL DEFAULT 0")
            if 'fuel_type' not in cars_cols: cursor.execute("ALTER TABLE cars ADD COLUMN fuel_type TEXT DEFAULT ''")
            cursor.execute("""CREATE TABLE IF NOT EXISTS drivers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS requesting_parties (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, notes TEXT, active INTEGER DEFAULT 1)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS trips (id INTEGER PRIMARY KEY AUTOINCREMENT, trip_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER NOT NULL, destination TEXT, start_location TEXT, trip_type TEXT, requester TEXT, requesting_party_id INTEGER, exit_time TEXT, return_time TEXT, start_km REAL, end_km REAL, distance REAL, notes TEXT)""")
            cursor.execute("PRAGMA table_info(trips)"); trips_cols = {r[1] for r in cursor.fetchall()}
            if 'requesting_party_id' not in trips_cols: cursor.execute("ALTER TABLE trips ADD COLUMN requesting_party_id INTEGER")
            cursor.execute("""CREATE TABLE IF NOT EXISTS fuel (id INTEGER PRIMARY KEY AUTOINCREMENT, fuel_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER, fuel_type TEXT, source TEXT, payment_status TEXT, liters REAL, price_per_liter REAL, total_cost REAL, final_cost REAL, km_at_fill REAL, notes TEXT)""")
            cursor.execute("PRAGMA table_info(fuel)"); fuel_cols = {r[1] for r in cursor.fetchall()}
            if 'car_image' not in fuel_cols: cursor.execute("ALTER TABLE fuel ADD COLUMN car_image TEXT")
            if 'station_image' not in fuel_cols: cursor.execute("ALTER TABLE fuel ADD COLUMN station_image TEXT")
            cursor.execute("""CREATE TABLE IF NOT EXISTS maintenance (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, car_id INTEGER NOT NULL, type TEXT, category TEXT, km_at_service REAL, cost REAL, invoice_number TEXT, invoice_image TEXT, workshop TEXT, notes TEXT, next_service_km REAL, next_service_date TEXT)""")
            cursor.execute("PRAGMA table_info(maintenance)"); maint_cols = {r[1] for r in cursor.fetchall()}
            if 'currency' not in maint_cols: cursor.execute("ALTER TABLE maintenance ADD COLUMN currency TEXT DEFAULT 'ل.س'")
            cursor.execute("""CREATE TABLE IF NOT EXISTS travel_missions (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, driver_id INTEGER NOT NULL, description TEXT, distance_km REAL DEFAULT 0, is_travel INTEGER DEFAULT 0, departure_time TEXT, return_time TEXT, notes TEXT)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS overtime (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, driver_id INTEGER NOT NULL, morning_hours REAL DEFAULT 0, evening_hours REAL DEFAULT 0, notes TEXT)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS daily_checks (id INTEGER PRIMARY KEY AUTOINCREMENT, check_date TEXT NOT NULL, check_time TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER NOT NULL, km_at_check REAL, oil_checked INTEGER DEFAULT 0, radiator_checked INTEGER DEFAULT 0, tires_checked INTEGER DEFAULT 0, battery_checked INTEGER DEFAULT 0, belts_checked INTEGER DEFAULT 0, oil_notes TEXT, radiator_notes TEXT, tires_notes TEXT, battery_notes TEXT, belts_notes TEXT, oil_image TEXT, radiator_image TEXT, tires_image TEXT, battery_image TEXT, belts_image TEXT, general_notes TEXT, admin_notes TEXT, has_issue INTEGER DEFAULT 0, created_at TEXT)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS pending_fuel (id INTEGER PRIMARY KEY AUTOINCREMENT, driver_id INTEGER, car_id INTEGER, fuel_date TEXT, fuel_time TEXT, liters REAL, price_per_liter REAL, total_cost REAL, final_cost REAL, km_at_fill REAL, station_image TEXT, car_image TEXT, notes TEXT, status TEXT DEFAULT 'pending', created_at TEXT)""")
        conn.commit(); conn.close(); return True
    except Exception as e: print(f"⚠ خطأ في ensure_all: {e}"); return False

def migrate_requesting_parties():
    try:
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "SELECT DISTINCT requester FROM trips WHERE requester IS NOT NULL AND requester != '' AND requesting_party_id IS NULL")
        for row in cursor.fetchall():
            name = row['requester'].strip() if row['requester'] else ''
            if not name: continue
            execute_query(cursor, "SELECT id FROM requesting_parties WHERE name=?", (name,)); existing = cursor.fetchone()
            if existing: party_id = existing['id']
            else:
                execute_query(cursor, "INSERT INTO requesting_parties (name, notes, active) VALUES (?, ?, 1)", (name, 'مستورد'))
                if is_postgres(): execute_query(cursor, "SELECT MAX(id) as id FROM requesting_parties")
                else: execute_query(cursor, "SELECT last_insert_rowid() as id")
                party_id = cursor.fetchone()['id']
            if party_id: execute_query(cursor, "UPDATE trips SET requesting_party_id=? WHERE requester=? AND requesting_party_id IS NULL", (party_id, row['requester']))
        conn.commit(); conn.close(); return True
    except Exception: return False

ensure_all_tables()
migrate_requesting_parties()

# ============== دوال الذكاء الاصطناعي وتلغرام ==============
def send_telegram_message(chat_id, text):
    if not TELEGRAM_TOKEN or TELEGRAM_TOKEN == 'ضع_توكن_البوت_هنا': 
        print("⚠ لم يتم إرسال رسالة تلغرام: التوكن غير مضبوط"); return
    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    try: 
        res = requests.post(url, json={'chat_id': chat_id, 'text': text}, timeout=10)
        if res.status_code != 200: print(f"⚠ فشل إرسال رسالة تلغرام: {res.text}")
    except Exception as e: print(f"⚠ خطأ في الاتصال بتلغرام: {e}")

def extract_data_from_images(car_img_b64=None, station_img_b64=None):
    if not GEMINI_API_KEY or GEMINI_API_KEY == 'ضع_مفتاح_جيميني_هنا': return {'km': None, 'liters': None, 'error': 'no key'}
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        contents = ["أنت مساعد خبير لقراءة عدادات السيارات وفواتير الوقود. المطلوب منك استخراج أرقام دقيقة."]
        if car_img_b64:
            data = car_img_b64.split(',')[1] if ',' in car_img_b64 else car_img_b64
            contents.append({"mime_type": "image/jpeg", "data": base64.b64decode(data)})
            contents.append("هذه صورة عداد السيارة (Odometer). استخرج قراءة العداد بالكيلومترات كرقم صحيح.")
        if station_img_b64:
            data = station_img_b64.split(',')[1] if ',' in station_img_b64 else station_img_b64
            contents.append({"mime_type": "image/jpeg", "data": base64.b64decode(data)})
            contents.append("هذه صورة عداد الكازية. استخرج عدد اللترات (Liters) كرقم عشري.")
        contents.append("الرجاء إرجاع النتيجة بصيغة JSON فقط كالتالي: {\"km\": 51756, \"liters\": 40.5}. إذا لم تجد أحد الرقمين اجعله null.")
        response = client.models.generate_content(model='gemini-1.5-flash', contents=contents)
        text = response.text.replace('```json', '').replace('```', '').strip()
        data = json.loads(text)
        return {'km': data.get('km'), 'liters': data.get('liters'), 'error': None}
    except Exception as e: print(f"⚠ خطأ في OCR: {e}"); return {'km': None, 'liters': None, 'error': str(e)}

@app.context_processor
def inject_permissions():
    pending_fuel_count = 0
    if is_admin():
        try:
            conn = get_connection(); cursor = conn.cursor()
            execute_query(cursor, "SELECT COUNT(*) as cnt FROM pending_fuel WHERE status='pending'")
            pending_fuel_count = cursor.fetchone()['cnt']; conn.close()
        except: pass
    return {'is_admin': is_admin(), 'is_viewer': is_viewer(), 'pending_fuel_count': pending_fuel_count}
    # ============== المسارات الأساسية ==============
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        password = request.form.get('password', ''); login_type = request.form.get('login_type', 'viewer')
        if login_type == 'admin':
            if password == get_admin_password(): session['role'] = 'admin'; return redirect(url_for('index'))
            else: flash('كلمة سر المسؤول غير صحيحة', 'error')
        else:
            if password == get_viewer_password(): session['role'] = 'viewer'; return redirect(url_for('index'))
            else: flash('كلمة السر غير صحيحة', 'error')
    return render_template('login.html')

@app.route('/logout')
def logout(): session.clear(); flash('تم تسجيل الخروج', 'success'); return redirect(url_for('login'))

@app.route('/settings/passwords', methods=['GET', 'POST'])
@login_required
def settings_passwords():
    if request.method == 'POST':
        current_password = request.form.get('current_password', ''); new_admin = request.form.get('new_admin_password', '').strip(); new_viewer = request.form.get('new_viewer_password', '').strip()
        passwords = load_passwords()
        if is_admin():
            if current_password != passwords.get('admin_password'): flash('كلمة السر الحالية غير صحيحة', 'error'); return redirect(url_for('settings_passwords'))
            if new_admin: passwords['admin_password'] = new_admin
            if new_viewer: passwords['viewer_password'] = new_viewer
            save_passwords(passwords); return redirect(url_for('restart_app'))
        else:
            if current_password != passwords.get('viewer_password'): flash('كلمة السر الحالية غير صحيحة', 'error'); return redirect(url_for('settings_passwords'))
            if new_viewer: passwords['viewer_password'] = new_viewer
            save_passwords(passwords); return redirect(url_for('restart_app'))
    return render_template('settings_passwords.html')

@app.route('/restart')
@login_required
def restart_app(): return render_template('restart.html')

@app.route('/restart/do')
@login_required
def do_restart():
    try: subprocess.Popen([sys.executable] + sys.argv, close_fds=True)
    except: pass
    return jsonify({'status': 'ok'})

@app.route('/')
@login_required
def index():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT COUNT(*) as cnt FROM cars WHERE active=1"); cars_count = cursor.fetchone()['cnt']
    execute_query(cursor, "SELECT COUNT(*) as cnt FROM drivers WHERE active=1"); drivers_count = cursor.fetchone()['cnt']
    now = datetime.now(); month_str = now.strftime('%Y-%m')
    execute_query(cursor, "SELECT COUNT(*) as cnt FROM trips WHERE trip_date LIKE ?", (month_str + '%',)); trips_month = cursor.fetchone()['cnt']
    execute_query(cursor, "SELECT SUM(distance) as total FROM trips WHERE trip_date LIKE ?", (month_str + '%',)); km_row = cursor.fetchone(); km_month = km_row['total'] if km_row and km_row['total'] else 0
    execute_query(cursor, "SELECT t.*, c.name as car_name, d.name as driver_name FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id ORDER BY t.id DESC LIMIT 10")
    recent_trips = cursor.fetchall()
    conn.close()
    stats = {'cars': cars_count, 'drivers': drivers_count, 'trips_month': trips_month, 'km_month': round(km_month, 1)}
    return render_template('index.html', stats=stats, trips_by_car=recent_trips, alerts=get_maintenance_alerts())

@app.route('/admin/pending')
@login_required
def admin_pending_trips():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT t.*, c.name as car_name, d.name as driver_name FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id WHERE t.trip_source = 'driver_app' AND (t.admin_seen IS NULL OR t.admin_seen = 0) ORDER BY t.id DESC")
    pending_trips = cursor.fetchall(); conn.close()
    return render_template('admin_pending.html', pending_trips=pending_trips)

@app.route('/add_trip', methods=['GET', 'POST'])
@admin_required
def add_trip():
    conn = get_connection(); cursor = conn.cursor()
    if request.method == 'POST':
        trip_date = request.form['trip_date']; car_id = request.form['car_id']; driver_id = request.form['driver_id']
        destination = request.form.get('destination', ''); start_km = request.form.get('start_km') or None; end_km = request.form.get('end_km') or None; notes = request.form.get('notes', ''); distance = None
        if start_km and end_km:
            try: distance = float(end_km) - float(start_km)
            except: distance = None
        execute_query(cursor, "INSERT INTO trips (trip_date, car_id, driver_id, destination, start_km, end_km, distance, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (trip_date, car_id, driver_id, destination, start_km, end_km, distance, notes))
        conn.commit(); conn.close(); return redirect(url_for('index'))
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name"); cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name"); drivers_list = cursor.fetchall(); conn.close()
    return render_template('add_trip.html', cars=cars_list, drivers=drivers_list)

@app.route('/trips/delete/<int:trip_id>')
@admin_required
def delete_trip(trip_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM trips WHERE id=?", (trip_id,)); conn.commit(); conn.close()
    return redirect(url_for('index'))

@app.route('/cars')
@login_required
def cars():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY id"); cars_list = cursor.fetchall(); conn.close()
    return render_template('cars.html', cars=cars_list)

@app.route('/cars/add', methods=['POST'])
@admin_required
def add_car():
    name = request.form['name']; plate = request.form.get('plate_number', ''); car_type = request.form.get('car_type', ''); notes = request.form.get('notes', ''); fuel_type = request.form.get('fuel_type', '')
    expected_consumption = request.form.get('expected_consumption') or 0
    try: expected_consumption = float(expected_consumption)
    except: expected_consumption = 0
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO cars (name, plate_number, car_type, notes, fuel_type, expected_consumption) VALUES (?, ?, ?, ?, ?, ?)", (name, plate, car_type, notes, fuel_type, expected_consumption))
    conn.commit(); conn.close(); return redirect(url_for('cars'))

@app.route('/cars/edit/<int:car_id>', methods=['GET', 'POST'])
@admin_required
def edit_car(car_id):
    conn = get_connection(); cursor = conn.cursor()
    if request.method == 'POST':
        name = request.form['name']; plate = request.form.get('plate_number', ''); car_type = request.form.get('car_type', ''); notes = request.form.get('notes', ''); fuel_type = request.form.get('fuel_type', '')
        expected_consumption = request.form.get('expected_consumption') or 0
        try: expected_consumption = float(expected_consumption)
        except: expected_consumption = 0
        execute_query(cursor, "UPDATE cars SET name=?, plate_number=?, car_type=?, notes=?, fuel_type=?, expected_consumption=? WHERE id=?", (name, plate, car_type, notes, fuel_type, expected_consumption, car_id))
        conn.commit(); conn.close(); return redirect(url_for('cars'))
    execute_query(cursor, "SELECT * FROM cars WHERE id=?", (car_id,)); car = cursor.fetchone(); conn.close()
    return render_template('cars.html', cars=[], edit_car=car)

@app.route('/cars/delete/<int:car_id>')
@admin_required
def delete_car(car_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM cars WHERE id=?", (car_id,)); conn.commit(); conn.close()
    return redirect(url_for('cars'))

@app.route('/drivers')
@login_required
def drivers():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY id"); drivers_list = cursor.fetchall()
    execute_query(cursor, "SELECT tm.*, d.name as driver_name FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id ORDER BY tm.date DESC, tm.id DESC"); all_missions = cursor.fetchall()
    execute_query(cursor, "SELECT o.*, d.name as driver_name FROM overtime o JOIN drivers d ON o.driver_id = d.id ORDER BY o.date DESC, o.id DESC"); all_overtime = cursor.fetchall()
    missions_by_driver = {}; overtime_by_driver = {}
    for m in all_missions: missions_by_driver.setdefault(m['driver_id'], []).append(dict(m))
    for o in all_overtime: overtime_by_driver.setdefault(o['driver_id'], []).append(dict(o))
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name"); active_drivers = cursor.fetchall(); conn.close()
    return render_template('drivers.html', drivers=drivers_list, active_drivers=active_drivers, missions_by_driver=missions_by_driver, overtime_by_driver=overtime_by_driver)

@app.route('/drivers/add', methods=['POST'])
@admin_required
def add_driver():
    name = request.form['name']; phone = request.form.get('phone', ''); notes = request.form.get('notes', '')
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO drivers (name, phone, notes) VALUES (?, ?, ?)", (name, phone, notes))
    conn.commit(); conn.close(); return redirect(url_for('drivers'))

@app.route('/drivers/delete/<int:driver_id>')
@admin_required
def delete_driver(driver_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM drivers WHERE id=?", (driver_id,)); conn.commit(); conn.close()
    return redirect(url_for('drivers'))

@app.route('/requesting_parties')
@login_required
def requesting_parties():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM requesting_parties ORDER BY id"); parties = cursor.fetchall(); conn.close()
    return render_template('requesting_parties.html', parties=parties)

@app.route('/requesting_parties/add', methods=['POST'])
@admin_required
def add_requesting_party():
    name = request.form['name'].strip(); notes = request.form.get('notes', '')
    if not name: flash('اسم الجهة مطلوب', 'error'); return redirect(url_for('requesting_parties'))
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO requesting_parties (name, notes, active) VALUES (?, ?, 1)", (name, notes))
    conn.commit(); conn.close(); return redirect(url_for('requesting_parties'))

@app.route('/requesting_parties/delete/<int:party_id>')
@admin_required
def delete_requesting_party(party_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM requesting_parties WHERE id=?", (party_id,)); conn.commit(); conn.close()
    return redirect(url_for('requesting_parties'))

@app.route('/requesting_parties/toggle/<int:party_id>')
@admin_required
def toggle_requesting_party(party_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT active FROM requesting_parties WHERE id=?", (party_id,)); row = cursor.fetchone()
    if not row: conn.close(); return redirect(url_for('requesting_parties'))
    new_active = 0 if row['active'] == 1 else 1
    execute_query(cursor, "UPDATE requesting_parties SET active=? WHERE id=?", (new_active, party_id)); conn.commit(); conn.close()
    return redirect(url_for('requesting_parties'))

@app.route('/daily_checks')
@login_required
def daily_checks_list():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT dc.*, c.name as car_name, d.name as driver_name FROM daily_checks dc JOIN cars c ON dc.car_id = c.id JOIN drivers d ON dc.driver_id = d.id ORDER BY dc.id DESC")
    checks = cursor.fetchall(); conn.close()
    return render_template('daily_checks.html', checks=checks)

@app.route('/daily_checks/view/<int:check_id>')
@login_required
def daily_check_view(check_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT dc.*, c.name as car_name, d.name as driver_name FROM daily_checks dc JOIN cars c ON dc.car_id = c.id JOIN drivers d ON dc.driver_id = d.id WHERE dc.id = ?", (check_id,))
    check = cursor.fetchone(); conn.close()
    if not check: return redirect(url_for('daily_checks_list'))
    return render_template('daily_check_view.html', check=check)

@app.route('/daily_checks/delete/<int:check_id>')
@admin_required
def daily_check_delete(check_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM daily_checks WHERE id=?", (check_id,)); conn.commit(); conn.close()
    return redirect(url_for('daily_checks_list'))
    @app.route('/travel_missions')
@login_required
def travel_missions():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT tm.*, d.name as driver_name FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id ORDER BY tm.date DESC, tm.id DESC"); missions = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name"); drivers_list = cursor.fetchall(); conn.close()
    return render_template('travel_missions.html', missions=missions, drivers=drivers_list)

@app.route('/travel_missions/add', methods=['POST'])
@admin_required
def add_travel_mission():
    date = request.form['date']; driver_id = request.form['driver_id']; description = request.form.get('description', ''); distance_km = request.form.get('distance_km') or 0
    try: distance_km = float(distance_km)
    except: distance_km = 0
    is_travel = 1 if distance_km >= 75 else 0
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO travel_missions (date, driver_id, description, distance_km, is_travel) VALUES (?, ?, ?, ?, ?)", (date, driver_id, description, distance_km, is_travel))
    conn.commit(); conn.close(); return redirect(url_for('drivers'))

@app.route('/overtime')
@login_required
def overtime():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT o.*, d.name as driver_name FROM overtime o JOIN drivers d ON o.driver_id = d.id ORDER BY o.date DESC, o.id DESC"); overtimes = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name"); drivers_list = cursor.fetchall(); conn.close()
    return render_template('overtime.html', overtimes=overtimes, drivers=drivers_list)

@app.route('/overtime/add', methods=['POST'])
@admin_required
def add_overtime():
    date = request.form['date']; driver_id = request.form['driver_id']; morning_hours = request.form.get('morning_hours') or 0; evening_hours = request.form.get('evening_hours') or 0; notes = request.form.get('notes', '')
    try: morning_hours = float(morning_hours)
    except: morning_hours = 0
    try: evening_hours = float(evening_hours)
    except: evening_hours = 0
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO overtime (date, driver_id, morning_hours, evening_hours, notes) VALUES (?, ?, ?, ?, ?)", (date, driver_id, morning_hours, evening_hours, notes))
    conn.commit(); conn.close(); return redirect(url_for('drivers'))

@app.route('/fuel')
@login_required
def fuel():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT f.*, c.name as car_name, d.name as driver_name FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id ORDER BY f.id DESC LIMIT 50")
    fuel_records = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name"); cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name"); drivers_list = cursor.fetchall(); conn.close()
    return render_template('fuel.html', fuel_records=fuel_records, cars=cars_list, drivers=drivers_list)

@app.route('/fuel/add', methods=['POST'])
@admin_required
def add_fuel():
    fuel_date = request.form['fuel_date']; car_id = request.form['car_id']; driver_id = request.form.get('driver_id') or None
    fuel_type = request.form.get('fuel_type', ''); source = request.form.get('source', ''); payment_status = request.form.get('payment_status', '')
    liters = request.form.get('liters') or None; price = request.form.get('price_per_liter') or None; total = request.form.get('total_cost') or None
    final = request.form.get('final_cost') or None; km = request.form.get('km_at_fill') or None; notes = request.form.get('notes', '')
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO fuel (fuel_date, car_id, driver_id, fuel_type, source, payment_status, liters, price_per_liter, total_cost, final_cost, km_at_fill, notes) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (fuel_date, car_id, driver_id, fuel_type, source, payment_status, liters, price, total, final, km, notes))
    conn.commit(); conn.close(); return redirect(url_for('fuel'))

@app.route('/fuel/delete/<int:fuel_id>')
@admin_required
def delete_fuel(fuel_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM fuel WHERE id=?", (fuel_id,)); conn.commit(); conn.close()
    return redirect(url_for('fuel'))

@app.route('/maintenance')
@login_required
def maintenance():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT m.*, c.name as car_name FROM maintenance m JOIN cars c ON m.car_id = c.id ORDER BY m.id DESC LIMIT 100"); records = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name"); cars_list = cursor.fetchall(); conn.close()
    return render_template('maintenance.html', records=records, cars=cars_list)

@app.route('/maintenance/add', methods=['POST'])
@admin_required
def add_maintenance():
    date = request.form['date']; car_id = request.form['car_id']; mtype = request.form['type']; category = request.form.get('category', 'دورية')
    km = request.form.get('km_at_service') or None; cost = request.form.get('cost') or None; currency = request.form.get('currency', 'ل.س')
    invoice_number = request.form.get('invoice_number', ''); workshop = request.form.get('workshop', ''); notes = request.form.get('notes', '')
    next_km = request.form.get('next_service_km') or None; next_date = request.form.get('next_service_date') or None
    invoice_image = None
    if 'invoice_image' in request.files:
        file = request.files['invoice_image']
        if file and file.filename and allowed_file(file.filename):
            filename = secure_filename(file.filename); timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_'); filename = timestamp + filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename)); invoice_image = filename
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO maintenance (date, car_id, type, category, km_at_service, cost, currency, invoice_number, invoice_image, workshop, notes, next_service_km, next_service_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (date, car_id, mtype, category, km, cost, currency, invoice_number, invoice_image, workshop, notes, next_km, next_date))
    conn.commit(); conn.close(); return redirect(url_for('maintenance'))

@app.route('/maintenance/delete/<int:m_id>')
@admin_required
def delete_maintenance(m_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM maintenance WHERE id=?", (m_id,)); conn.commit(); conn.close()
    return redirect(url_for('maintenance'))

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    if not filename: return jsonify({'error': 'not found'}), 404
    if SUPABASE_URL and SUPABASE_BUCKET:
        local_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        if os.path.exists(local_path): return send_from_directory(app.config['UPLOAD_FOLDER'], filename)
        return redirect(f"{SUPABASE_URL}/storage/v1/object/public/{SUPABASE_BUCKET}/{filename}")
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)
    # ============== التقارير ==============
@app.route('/reports')
@login_required
def reports():
    month = int(request.args.get('month', datetime.now().month)); year = int(request.args.get('year', datetime.now().year))
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY name"); cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY name"); drivers_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM requesting_parties ORDER BY name"); parties_list = cursor.fetchall()
    years = list(range(2024, datetime.now().year + 2)); report_data = None
    if request.args.get('month'):
        month_str = f"{year:04d}-{month:02d}"
        execute_query(cursor, "SELECT t.*, c.name as car_name, d.name as driver_name FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id WHERE t.trip_date LIKE ?", (month_str + '%',)); trips = cursor.fetchall()
        execute_query(cursor, "SELECT f.*, c.name as car_name FROM fuel f JOIN cars c ON f.car_id = c.id WHERE f.fuel_date LIKE ?", (month_str + '%',)); fuels = cursor.fetchall()
        execute_query(cursor, "SELECT m.*, c.name as car_name FROM maintenance m JOIN cars c ON m.car_id = c.id WHERE m.date LIKE ?", (month_str + '%',)); maints = cursor.fetchall()
        report_data = {'summary': {
            'total_trips': len(trips), 'total_km': sum(t['distance'] or 0 for t in trips),
            'total_liters': sum(f['liters'] or 0 for f in fuels), 'total_fuel_cost': sum(get_effective_cost(f) for f in fuels),
            'total_maint_cost_syp': sum(m['cost'] or 0 for m in maints if m.get('currency', 'ل.س') == 'ل.س'),
            'total_maint_cost_usd': sum(m['cost'] or 0 for m in maints if m.get('currency') == 'دولار'),
        }}
    conn.close()
    return render_template('reports.html', cars=cars_list, drivers=drivers_list, parties=parties_list, years=years, selected_month=month, selected_year=year, report_data=report_data)

@app.route('/reports/trips')
@login_required
def reports_trips():
    car_id = request.args.get('car_id') or ''; driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''; date_to = request.args.get('date_to') or ''
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY name"); cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY name"); drivers_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM requesting_parties ORDER BY name"); parties_list = cursor.fetchall()
    query = "SELECT t.*, c.name as car_name, d.name as driver_name FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id WHERE 1=1"
    params = []
    if car_id: query += " AND t.car_id = ?"; params.append(car_id)
    if driver_id: query += " AND t.driver_id = ?"; params.append(driver_id)
    if date_from: query += " AND t.trip_date >= ?"; params.append(date_from)
    if date_to: query += " AND t.trip_date <= ?"; params.append(date_to)
    query += " ORDER BY t.trip_date DESC, t.id DESC"
    execute_query(cursor, query, params); trips = cursor.fetchall()
    summary = {'total_trips': len(trips), 'total_km': sum(t['distance'] or 0 for t in trips)}
    conn.close()
    return render_template('reports_trips.html', cars=cars_list, drivers=drivers_list, parties=parties_list, trips=trips, summary=summary, selected_car=car_id, selected_driver=driver_id, date_from=date_from, date_to=date_to, has_filter=bool(car_id or driver_id or date_from or date_to))

@app.route('/reports/fuel')
@login_required
def reports_fuel():
    car_id = request.args.get('car_id') or ''; driver_id = request.args.get('driver_id') or ''; date_from = request.args.get('date_from') or ''; date_to = request.args.get('date_to') or ''; payment_filter = request.args.get('payment_filter') or ''
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY name"); cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY name"); drivers_list = cursor.fetchall()
    query = "SELECT f.*, c.name as car_name, d.name as driver_name FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id WHERE 1=1"
    params = []
    if car_id: query += " AND f.car_id = ?"; params.append(car_id)
    if driver_id: query += " AND f.driver_id = ?"; params.append(driver_id)
    if date_from: query += " AND f.fuel_date >= ?"; params.append(date_from)
    if date_to: query += " AND f.fuel_date <= ?"; params.append(date_to)
    if payment_filter:
        if payment_filter == 'خزان الشركة': query += " AND f.source = ?"; params.append('خزان الشركة')
        elif payment_filter == 'كازية': query += " AND f.source = ?"; params.append('كازية')
        else: query += " AND f.payment_status = ?"; params.append(payment_filter)
    query += " ORDER BY f.fuel_date DESC, f.id DESC"
    execute_query(cursor, query, params); fuel_records = cursor.fetchall()
    summary = {'total_records': len(fuel_records), 'total_liters': sum(f['liters'] or 0 for f in fuel_records), 'total_cost': sum(get_effective_cost(f) for f in fuel_records)}
    conn.close()
    return render_template('reports_fuel.html', cars=cars_list, drivers=drivers_list, fuel_records=fuel_records, summary=summary, selected_car=car_id, selected_driver=driver_id, date_from=date_from, date_to=date_to, payment_filter=payment_filter, has_filter=bool(car_id or driver_id or date_from or date_to or payment_filter))

# ============== تصدير الإكسيل ==============
@app.route('/reports/export')
@login_required
def export_excel():
    month = int(request.args.get('month', datetime.now().month)); year = int(request.args.get('year', datetime.now().year))
    month_str = f"{year:04d}-{month:02d}"
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT t.*, c.name as car_name, d.name as driver_name FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id WHERE t.trip_date LIKE ?", (month_str + '%',)); trips = cursor.fetchall()
    execute_query(cursor, "SELECT f.*, c.name as car_name FROM fuel f JOIN cars c ON f.car_id = c.id WHERE f.fuel_date LIKE ?", (month_str + '%',)); fuels = cursor.fetchall()
    execute_query(cursor, "SELECT m.*, c.name as car_name FROM maintenance m JOIN cars c ON m.car_id = c.id WHERE m.date LIKE ?", (month_str + '%',)); maints = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM cars ORDER BY name"); cars_list = cursor.fetchall()
    conn.close()
    
    wb = Workbook(); ws_sum = wb.active; ws_sum.title = "الملخص العام"; ws_sum.sheet_view.rightToLeft = True
    ws_sum.append(['التقرير الشهري الشامل', month_str]); ws_sum.append([]); ws_sum.append(['الإحصائية', 'القيمة'])
    ws_sum.append(['إجمالي الرحلات', len(trips)]); ws_sum.append(['إجمالي الكيلومترات', sum(t['distance'] or 0 for t in trips)])
    ws_sum.append(['إجمالي اللترات', sum(f['liters'] or 0 for f in fuels)]); ws_sum.append(['إجمالي كلفة الوقود', sum(get_effective_cost(f) for f in fuels)])
    ws_sum.append(['إجمالي كلفة الصيانة (ل.س)', sum(m['cost'] or 0 for m in maints if m.get('currency', 'ل.س') == 'ل.س')])
    ws_sum.append(['إجمالي كلفة الصيانة (دولار)', sum(m['cost'] or 0 for m in maints if m.get('currency') == 'دولار')])
    
    for c in cars_list:
        c_trips = [t for t in trips if t['car_id'] == c['id']]; c_fuel = [f for f in fuels if f['car_id'] == c['id']]; c_maint = [m for m in maints if m['car_id'] == c['id']]
        if not c_trips and not c_fuel and not c_maint: continue
        ws_c = wb.create_sheet(title=f"سيارة - {c['name'][:20]}"); ws_c.sheet_view.rightToLeft = True
        ws_c.append([f"تقرير السيارة: {c['name']}"]); ws_c.append([])
        ws_c.append(['الرحلات:', 'التاريخ', 'السائق', 'المسافة (كم)'])
        for t in c_trips: ws_c.append(['', t['trip_date'], t['driver_name'], t['distance'] or 0])
        ws_c.append([]); ws_c.append(['الوقود:', 'التاريخ', 'اللترات', 'الكلفة'])
        for f in c_fuel: ws_c.append(['', f['fuel_date'], f['liters'] or 0, get_effective_cost(f)])
        ws_c.append([]); ws_c.append(['الصيانة:', 'التاريخ', 'النوع', 'التكلفة', 'العملة'])
        for m in c_maint: ws_c.append(['', m['date'], m['type'], m['cost'] or 0, m.get('currency', 'ل.س')])
        
    ws3 = wb.create_sheet(f"تفاصيل الرحلات"); ws3.sheet_view.rightToLeft = True
    ws3.append(['التاريخ', 'السيارة', 'السائق', 'الوجهة', 'المسافة (كم)', 'ملاحظات'])
    for t in trips: ws3.append([t['trip_date'], t['car_name'], t['driver_name'], t['destination'] or '', t['distance'] or 0, t['notes'] or ''])
    
    output = io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name=f'qamh_report_{month_str}.xlsx')

@app.route('/reports/trips/export')
@login_required
def export_trips_excel():
    car_id = request.args.get('car_id') or ''; driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''; date_to = request.args.get('date_to') or ''
    conn = get_connection(); cursor = conn.cursor()
    query = "SELECT t.*, c.name as car_name, d.name as driver_name FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id WHERE 1=1"
    params = []
    if car_id: query += " AND t.car_id = ?"; params.append(car_id)
    if driver_id: query += " AND t.driver_id = ?"; params.append(driver_id)
    if date_from: query += " AND t.trip_date >= ?"; params.append(date_from)
    if date_to: query += " AND t.trip_date <= ?"; params.append(date_to)
    query += " ORDER BY t.trip_date DESC, t.id DESC"
    execute_query(cursor, query, params); trips = cursor.fetchall(); conn.close()
    wb = Workbook(); ws = wb.active; ws.title = "Trips Report"; ws.sheet_view.rightToLeft = True
    ws.append(['التاريخ', 'السيارة', 'السائق', 'الوجهة', 'المسافة (كم)', 'ملاحظات'])
    for t in trips: ws.append([t['trip_date'], t['car_name'], t['driver_name'], t['destination'] or '', t['distance'] or 0, t['notes'] or ''])
    output = io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='trips_report.xlsx')

@app.route('/reports/fuel/export')
@login_required
def export_fuel_excel():
    car_id = request.args.get('car_id') or ''; driver_id = request.args.get('driver_id') or ''; date_from = request.args.get('date_from') or ''; date_to = request.args.get('date_to') or ''
    conn = get_connection(); cursor = conn.cursor()
    query = "SELECT f.*, c.name as car_name, d.name as driver_name FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id WHERE 1=1"
    params = []
    if car_id: query += " AND f.car_id = ?"; params.append(car_id)
    if driver_id: query += " AND f.driver_id = ?"; params.append(driver_id)
    if date_from: query += " AND f.fuel_date >= ?"; params.append(date_from)
    if date_to: query += " AND f.fuel_date <= ?"; params.append(date_to)
    query += " ORDER BY f.fuel_date DESC, f.id DESC"
    execute_query(cursor, query, params); fuel_records = cursor.fetchall(); conn.close()
    wb = Workbook(); ws = wb.active; ws.title = "Fuel Report"; ws.sheet_view.rightToLeft = True
    ws.append(['التاريخ', 'السيارة', 'السائق', 'اللترات', 'الكلفة'])
    for f in fuel_records: ws.append([f['fuel_date'], f['car_name'], f['driver_name'], f['liters'] or 0, get_effective_cost(f)])
    output = io.BytesIO(); wb.save(output); output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', as_attachment=True, download_name='fuel_report.xlsx')
    # ============== Driver App API ==============
@app.route('/driver')
def driver_app(): return render_template('driver_app.html')

@app.route('/api/driver/init', methods=['GET'])
def api_driver_init():
    try:
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "SELECT id, name, phone FROM drivers WHERE active=1 ORDER BY name"); drivers_list = [dict(row) for row in cursor.fetchall()]
        execute_query(cursor, "SELECT id, name, plate_number FROM cars WHERE active=1 ORDER BY name"); cars_list = [dict(row) for row in cursor.fetchall()]
        execute_query(cursor, "SELECT id, name FROM requesting_parties WHERE active=1 ORDER BY name"); parties_list = [dict(row) for row in cursor.fetchall()]
        busy_driver_ids = list(get_busy_driver_ids(conn)); busy_car_ids = list(get_busy_car_ids(conn)); conn.close()
        alerts = get_maintenance_alerts(); critical_alerts = [a for a in alerts if a['status'] == 'red']
        return jsonify({'success': True, 'drivers': drivers_list, 'cars': cars_list, 'parties': parties_list, 'busy_driver_ids': busy_driver_ids, 'busy_car_ids': busy_car_ids, 'alerts': critical_alerts})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/alerts', methods=['GET'])
def api_driver_alerts():
    try:
        alerts = get_maintenance_alerts(); critical_alerts = [a for a in alerts if a['status'] == 'red']
        return jsonify({'success': True, 'alerts': critical_alerts})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/last_check/<int:car_id>', methods=['GET'])
def api_driver_last_check(car_id):
    try:
        conn = get_connection(); last_check = get_last_daily_check(car_id, conn); conn.close()
        if last_check:
            return jsonify({'success': True, 'needs_check': False, 'last_check': {'id': last_check['id'], 'check_date': last_check['check_date'], 'check_time': last_check['check_time'], 'driver_name': last_check['driver_name'], 'general_notes': last_check['general_notes'] or '', 'has_issue': last_check['has_issue']}})
        return jsonify({'success': True, 'needs_check': True})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/daily_check', methods=['POST'])
def api_driver_daily_check():
    try:
        data = request.get_json(); car_id = data.get('car_id'); driver_id = data.get('driver_id'); km_at_check = data.get('km_at_check') or None
        general_notes = data.get('general_notes', ''); has_issue = 1 if data.get('has_issue') else 0
        if not car_id or not driver_id: return jsonify({'success': False, 'error': 'الرجاء اختيار السائق والسيارة'}), 400
        now = datetime.now() + timedelta(hours=3); check_date = now.strftime('%Y-%m-%d'); check_time = now.strftime('%H:%M:%S'); created_at = now.strftime('%Y-%m-%d %H:%M:%S')
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "INSERT INTO daily_checks (check_date, check_time, car_id, driver_id, km_at_check, general_notes, has_issue, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", (check_date, check_time, car_id, driver_id, km_at_check, general_notes, has_issue, created_at))
        conn.commit(); conn.close()
        return jsonify({'success': True, 'message': 'تم حفظ الفحص اليومي'})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/start_trip', methods=['POST'])
def api_driver_start_trip():
    try:
        data = request.get_json(); driver_id = data.get('driver_id'); car_id = data.get('car_id'); destination = data.get('destination', '').strip()
        start_km = data.get('start_km')
        if not driver_id or not car_id: return jsonify({'success': False, 'error': 'الرجاء اختيار السائق والسيارة'}), 400
        if not destination: return jsonify({'success': False, 'error': 'الرجاء إدخال الوجهة'}), 400
        if start_km is None or str(start_km).strip() == '': return jsonify({'success': False, 'error': 'الرجاء إدخال قراءة العداد'}), 400
        try:
            start_km = float(start_km)
        except: return jsonify({'success': False, 'error': 'قراءة العداد غير صحيحة'}), 400
        now = datetime.now() + timedelta(hours=3); trip_date = now.strftime('%Y-%m-%d'); start_time = now.strftime('%H:%M:%S')
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "INSERT INTO trips (trip_date, car_id, driver_id, destination, start_km, start_time, trip_source, admin_seen) VALUES (?, ?, ?, ?, ?, ?, 'driver_app', 0)", (trip_date, car_id, driver_id, destination, start_km, start_time))
        conn.commit(); conn.close()
        return jsonify({'success': True, 'message': 'تم بدء الرحلة بنجاح'})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/end_trip', methods=['POST'])
def api_driver_end_trip():
    try:
        data = request.get_json(); trip_id = data.get('trip_id'); end_km = data.get('end_km'); notes = data.get('notes', '')
        if not trip_id: return jsonify({'success': False, 'error': 'رقم الرحلة مفقود'}), 400
        try:
            end_km = float(end_km)
        except: return jsonify({'success': False, 'error': 'قراءة العداد غير صحيحة'}), 400
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "SELECT start_km FROM trips WHERE id=?", (trip_id,)); trip = cursor.fetchone()
        if not trip: conn.close(); return jsonify({'success': False, 'error': 'الرحلة غير موجودة'}), 404
        start_km = trip['start_km']
        distance = None
        if start_km is not None:
            distance = end_km - float(start_km)
        now = datetime.now() + timedelta(hours=3); end_time = now.strftime('%H:%M:%S')
        execute_query(cursor, "UPDATE trips SET end_km=?, end_time=?, distance=?, notes=? WHERE id=?", (end_km, end_time, distance, notes, trip_id))
        conn.commit(); conn.close()
        return jsonify({'success': True, 'message': 'تم إنهاء الرحلة', 'distance': distance})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/active_trip/<int:driver_id>', methods=['GET'])
def api_driver_active_trip(driver_id):
    try:
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "SELECT t.*, c.name as car_name FROM trips t JOIN cars c ON t.car_id = c.id WHERE t.driver_id = ? AND t.trip_source = 'driver_app' AND t.end_km IS NULL ORDER BY t.id DESC LIMIT 1", (driver_id,))
        trip = cursor.fetchone(); conn.close()
        if trip: return jsonify({'success': True, 'has_active': True, 'trip': dict(trip)})
        return jsonify({'success': True, 'has_active': False})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/recent_trips/<int:driver_id>', methods=['GET'])
def api_driver_recent_trips(driver_id):
    try:
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "SELECT t.id, t.trip_date, t.destination, t.start_km, t.end_km, t.distance, t.start_time, t.end_time, c.name as car_name FROM trips t JOIN cars c ON t.car_id = c.id WHERE t.driver_id = ? ORDER BY t.id DESC LIMIT 5", (driver_id,))
        trips = [dict(row) for row in cursor.fetchall()]; conn.close()
        return jsonify({'success': True, 'trips': trips})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/api/driver/pending_count/<int:driver_id>', methods=['GET'])
def api_driver_pending_count(driver_id):
    try:
        conn = get_connection(); cursor = conn.cursor()
        execute_query(cursor, "SELECT COUNT(*) as cnt FROM trips WHERE driver_id = ? AND trip_source = 'driver_app' AND end_km IS NULL", (driver_id,))
        count = cursor.fetchone()['cnt']; conn.close()
        return jsonify({'success': True, 'count': count})
    except Exception as e: return jsonify({'success': False, 'error': str(e)}), 500
        # ============== Telegram Webhook & Pending Fuel ==============
telegram_pending_photos = {}

@app.route('/telegram_webhook', methods=['POST'])
def telegram_webhook():
    data = request.get_json()
    if 'message' in data:
        chat_id = data['message']['chat']['id']; text = data['message'].get('text', ''); photos = data['message'].get('photo', [])
        
        if photos:
            file_id = photos[-1]['file_id']
            file_info_url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/getFile?file_id={file_id}"
            
            try:
                file_info = requests.get(file_info_url).json()
                file_path = file_info['result']['file_path']
                img_url = f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_path}"
                img_data = requests.get(img_url).content
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"tg_fuel_{chat_id}_{timestamp}.jpg"
                
                uploaded_filename = upload_image_to_supabase(filename, img_data)
                if not uploaded_filename:
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    with open(filepath, 'wb') as f: f.write(img_data)
                    uploaded_filename = filename

                img_b64 = base64.b64encode(img_data).decode('utf-8')
                
                if chat_id not in telegram_pending_photos:
                    telegram_pending_photos[chat_id] = {'car_image': uploaded_filename, 'car_b64': img_b64, 'station_image': None, 'station_b64': None, 'text': text}
                    send_telegram_message(chat_id, "✅ تم استلام صورة عداد السيارة. يرجى إرسال صورة عداد الكازية الآن لإكمال الطلب.")
                else:
                    telegram_pending_photos[chat_id]['station_image'] = uploaded_filename
                    telegram_pending_photos[chat_id]['station_b64'] = img_b64
                    
                    session = telegram_pending_photos[chat_id]
                    extracted = extract_data_from_images(car_img_b64=session['car_b64'], station_img_b64=session['station_b64'])
                    
                    conn = get_connection(); cursor = conn.cursor()
                    now = datetime.now() + timedelta(hours=3)
                    
                    execute_query(cursor, "INSERT INTO pending_fuel (driver_id, car_id, fuel_date, fuel_time, liters, km_at_fill, status, created_at, notes, car_image, station_image) VALUES (NULL, NULL, ?, ?, ?, ?, 'pending', ?, ?, ?, ?)", (now.strftime('%Y-%m-%d'), now.strftime('%H:%M:%S'), extracted.get('liters'), extracted.get('km'), now.strftime('%Y-%m-%d %H:%M:%S'), f"صورتين مرسلتين من تلغرام. النص: {text}", session['car_image'], session['station_image']))
                    conn.commit(); conn.close()
                    
                    del telegram_pending_photos[chat_id]
                    
                    send_telegram_message(chat_id, f"✅ تم استلام الصورتين بنجاح! العداد المقروء: {extracted.get('km') or 'غير واضح'}. اللترات: {extracted.get('liters') or 'غير واضحة'}. بانتظار مراجعة الإدارة.")
            except Exception as e:
                send_telegram_message(chat_id, f"❌ خطأ في معالجة الصورة: {str(e)}")
        elif text:
            send_telegram_message(chat_id, "ℹ️ يرجى إرسال صورة عداد السيارة أولاً، ثم صورة عداد الكازية.")
            
    return jsonify({'status': 'ok'})

@app.route('/admin/pending_fuel')
@login_required
def admin_pending_fuel():
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT pf.*, d.name as driver_name, c.name as car_name FROM pending_fuel pf LEFT JOIN drivers d ON pf.driver_id = d.id LEFT JOIN cars c ON pf.car_id = c.id WHERE pf.status = 'pending' ORDER BY pf.id DESC")
    pending_fuels = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name"); cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name"); drivers_list = cursor.fetchall(); conn.close()
    return render_template('admin_pending_fuel.html', pending_fuels=pending_fuels, cars=cars_list, drivers=drivers_list)

@app.route('/admin/pending_fuel/approve/<int:req_id>', methods=['POST'])
@admin_required
def approve_pending_fuel(req_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM pending_fuel WHERE id=?", (req_id,)); req = cursor.fetchone()
    if req:
        fuel_date = request.form.get('fuel_date') or req['fuel_date']; car_id = request.form.get('car_id'); driver_id = request.form.get('driver_id') or None
        fuel_type = request.form.get('fuel_type', 'بنزين'); source = request.form.get('source', 'كازية'); payment_status = request.form.get('payment_status', '')
        liters = request.form.get('liters') or None; price = request.form.get('price_per_liter') or None; total = request.form.get('total_cost') or None
        final = request.form.get('final_cost') or None; km = request.form.get('km_at_fill') or None; notes = request.form.get('notes', '')
        
        car_image = req['car_image'] if 'car_image' in req.keys() else None
        station_image = req['station_image'] if 'station_image' in req.keys() else None

        execute_query(cursor, "INSERT INTO fuel (fuel_date, car_id, driver_id, fuel_type, source, payment_status, liters, price_per_liter, total_cost, final_cost, km_at_fill, notes, car_image, station_image) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)", (fuel_date, car_id, driver_id, fuel_type, source, payment_status, liters, price, total, final, km, notes, car_image, station_image))
        execute_query(cursor, "UPDATE pending_fuel SET status='approved' WHERE id=?", (req_id,))
        conn.commit(); flash('تمت الموافقة وإدخال التعبئة في النظام بنجاح', 'success')
    else: flash('الطلب غير موجود', 'error')
    conn.close(); return redirect(url_for('admin_pending_fuel'))

@app.route('/admin/pending_fuel/reject/<int:req_id>')
@admin_required
def reject_pending_fuel(req_id):
    conn = get_connection(); cursor = conn.cursor()
    execute_query(cursor, "UPDATE pending_fuel SET status='rejected' WHERE id=?", (req_id,))
    conn.commit(); conn.close(); flash('تم رفض الطلب', 'success'); return redirect(url_for('admin_pending_fuel'))

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 60); print(f"  Database: {'PostgreSQL' if is_postgres() else 'SQLite'}")
    print("=" * 60)
    app.run(host='0.0.0.0', port=port, debug=False)
