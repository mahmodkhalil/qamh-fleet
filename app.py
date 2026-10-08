from flask import Flask, render_template, request, redirect, url_for, flash, send_file, session, send_from_directory, jsonify
from werkzeug.utils import secure_filename
from functools import wraps
import sqlite3
import json
import sys
import subprocess
import os
import base64
from datetime import datetime
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

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

# ============== مسار البيانات ==============
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


def load_passwords():
    if not os.path.exists(PASSWORDS_FILE):
        default = {'admin_password': 'admin2026', 'viewer_password': 'qamh2026'}
        with open(PASSWORDS_FILE, 'w', encoding='utf-8') as f:
            json.dump(default, f, ensure_ascii=False, indent=4)
        return default
    try:
        with open(PASSWORDS_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {'admin_password': 'admin2026', 'viewer_password': 'qamh2026'}


def save_passwords(data):
    with open(PASSWORDS_FILE, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


def get_admin_password():
    return load_passwords().get('admin_password', 'admin2026')


def get_viewer_password():
    return load_passwords().get('viewer_password', 'qamh2026')


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def get_connection():
    if DATABASE_URL:
        conn = psycopg2.connect(DATABASE_URL, cursor_factory=RealDictCursor)
        return conn
    else:
        conn = sqlite3.connect(DB_PATH)
        conn.row_factory = sqlite3.Row
        return conn


def is_postgres():
    return DATABASE_URL is not None


def execute_query(cursor, query, params=()):
    if is_postgres():
        query = query.replace('?', '%s')
    cursor.execute(query, params)
    return cursor


def is_admin():
    return session.get('role') == 'admin'


def is_viewer():
    return session.get('role') in ('admin', 'viewer')


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not is_viewer():
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not is_admin():
            flash('هذا الإجراء متاح للمسؤول فقط', 'error')
            return redirect(url_for('index'))
        return f(*args, **kwargs)
    return decorated_function


def get_car_current_km(car_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT MAX(end_km) as max_km FROM trips WHERE car_id=? AND end_km IS NOT NULL", (car_id,))
    row = cursor.fetchone()
    trip_km = row['max_km'] if row else None
    execute_query(cursor, "SELECT MAX(km_at_service) as max_km FROM maintenance WHERE car_id=? AND km_at_service IS NOT NULL", (car_id,))
    row = cursor.fetchone()
    maint_km = row['max_km'] if row else None
    conn.close()
    candidates = [x for x in [trip_km, maint_km] if x is not None]
    return max(candidates) if candidates else 0


def get_maintenance_alerts():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT m.*, c.name as car_name
        FROM maintenance m JOIN cars c ON m.car_id = c.id
        WHERE m.next_service_km IS NOT NULL AND m.category = 'دورية'
        ORDER BY m.car_id, m.km_at_service DESC
    """)
    all_maint = cursor.fetchall()
    conn.close()
    latest_per_car = {}
    for m in all_maint:
        car_id = m['car_id']
        if car_id not in latest_per_car:
            latest_per_car[car_id] = m
    alerts = []
    for car_id, m in latest_per_car.items():
        current_km = get_car_current_km(car_id)
        next_km = m['next_service_km']
        remaining = next_km - current_km
        if remaining <= 0:
            status = 'red'
            message = f"تجاوزت الصيانة بـ {abs(remaining):.0f} كم"
        elif remaining <= 500:
            status = 'yellow'
            message = f"باقي {remaining:.0f} كم للصيانة"
        else:
            continue
        alerts.append({
            'car_name': m['car_name'], 'car_id': car_id, 'type': m['type'],
            'current_km': current_km, 'next_km': next_km, 'remaining': remaining,
            'status': status, 'message': message, 'last_date': m['date']
        })
    alerts.sort(key=lambda x: x['remaining'])
    return alerts


def should_count_cost(f):
    if f['source'] == 'خزان الشركة':
        return False
    ps = f['payment_status'] or ''
    if ps == 'مدفوعه من قسم الحركه' or ps == 'مدفوعة من قسم الحركة':
        return True
    return False


def get_effective_cost(f):
    if not should_count_cost(f):
        return 0
    if f['final_cost']:
        return f['final_cost']
    return f['total_cost'] or 0


def get_estimated_cost(f):
    if f['final_cost']:
        return f['final_cost']
    if f['total_cost']:
        return f['total_cost']
    if f['price_per_liter'] and f['liters']:
        try:
            return float(f['price_per_liter']) * float(f['liters'])
        except (ValueError, TypeError):
            return 0
    return 0


def get_car_cost_per_km(car_id, conn=None):
    close_conn = False
    if conn is None:
        conn = get_connection()
        close_conn = True
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT * FROM fuel 
        WHERE car_id=? AND km_at_fill IS NOT NULL 
        ORDER BY km_at_fill ASC, fuel_date ASC
    """, (car_id,))
    fills = cursor.fetchall()
    if close_conn:
        conn.close()
    total_km = 0
    total_cost = 0
    for i in range(1, len(fills)):
        prev_km = fills[i-1]['km_at_fill']
        curr_km = fills[i]['km_at_fill']
        if prev_km is None or curr_km is None:
            continue
        period_km = curr_km - prev_km
        if period_km <= 0:
            continue
        period_cost = get_estimated_cost(fills[i])
        total_km += period_km
        total_cost += period_cost
    if total_km <= 0:
        return 0
    return total_cost / total_km


def get_car_liters_per_100km(car_id, conn=None):
    close_conn = False
    if conn is None:
        conn = get_connection()
        close_conn = True
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT * FROM fuel 
        WHERE car_id=? AND km_at_fill IS NOT NULL 
        ORDER BY km_at_fill ASC, fuel_date ASC
    """, (car_id,))
    fills = cursor.fetchall()
    if close_conn:
        conn.close()
    total_km = 0
    total_liters = 0
    for i in range(1, len(fills)):
        prev_km = fills[i-1]['km_at_fill']
        curr_km = fills[i]['km_at_fill']
        if prev_km is None or curr_km is None:
            continue
        period_km = curr_km - prev_km
        if period_km <= 0:
            continue
        period_liters = fills[i]['liters'] or 0
        total_km += period_km
        total_liters += period_liters
    if total_km <= 0:
        return 0
    return (total_liters / total_km) * 100


def ensure_trips_columns():
    try:
        conn = get_connection()
        cursor = conn.cursor()
        if is_postgres():
            columns_to_add = [
                ('start_image', 'TEXT'),
                ('end_image', 'TEXT'),
                ('start_time', 'TEXT'),
                ('end_time', 'TEXT'),
                ('trip_source', "TEXT DEFAULT 'manual'"),
                ('driver_confirmed', 'INTEGER DEFAULT 0'),
                ('admin_seen', 'INTEGER DEFAULT 0'),
                ('admin_seen_at', 'TEXT'),
            ]
            for col_name, col_type in columns_to_add:
                try:
                    execute_query(cursor, f"ALTER TABLE trips ADD COLUMN IF NOT EXISTS {col_name} {col_type}")
                except Exception as e:
                    print(f"⚠ خطأ في إضافة {col_name}: {e}")
        else:
            cursor.execute("PRAGMA table_info(trips)")
            existing = {row[1] for row in cursor.fetchall()}
            new_columns = [
                ("start_image", "TEXT"), ("end_image", "TEXT"),
                ("start_time", "TEXT"), ("end_time", "TEXT"),
                ("trip_source", "TEXT DEFAULT 'manual'"),
                ("driver_confirmed", "INTEGER DEFAULT 0"),
                ("admin_seen", "INTEGER DEFAULT 0"),
                ("admin_seen_at", "TEXT"),
            ]
            for col_name, col_type in new_columns:
                if col_name not in existing:
                    try:
                        cursor.execute(f"ALTER TABLE trips ADD COLUMN {col_name} {col_type}")
                    except Exception as e:
                        print(f"⚠ خطأ في إضافة {col_name}: {e}")
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        print(f"⚠ خطأ في ensure_trips_columns: {e}")
        return False


def ensure_all_tables():
    try:
        conn = get_connection()
        cursor = conn.cursor()
        if is_postgres():
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS cars (id SERIAL PRIMARY KEY, name TEXT NOT NULL, plate_number TEXT, car_type TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS drivers (id SERIAL PRIMARY KEY, name TEXT NOT NULL, phone TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS trips (id SERIAL PRIMARY KEY, trip_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER NOT NULL, destination TEXT, start_location TEXT, trip_type TEXT, requester TEXT, exit_time TEXT, return_time TEXT, start_km REAL, end_km REAL, distance REAL, notes TEXT, start_image TEXT, end_image TEXT, start_time TEXT, end_time TEXT, trip_source TEXT DEFAULT 'manual', driver_confirmed INTEGER DEFAULT 0, admin_seen INTEGER DEFAULT 0, admin_seen_at TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS fuel (id SERIAL PRIMARY KEY, fuel_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER, fuel_type TEXT, source TEXT, payment_status TEXT, liters REAL, price_per_liter REAL, total_cost REAL, final_cost REAL, km_at_fill REAL, notes TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS maintenance (id SERIAL PRIMARY KEY, date TEXT NOT NULL, car_id INTEGER NOT NULL, type TEXT, category TEXT, km_at_service REAL, cost REAL, invoice_number TEXT, invoice_image TEXT, workshop TEXT, notes TEXT, next_service_km REAL, next_service_date TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS travel_missions (id SERIAL PRIMARY KEY, date TEXT NOT NULL, driver_id INTEGER NOT NULL, description TEXT, distance_km REAL DEFAULT 0, is_travel INTEGER DEFAULT 0, departure_time TEXT, return_time TEXT, notes TEXT)""")
            execute_query(cursor, """CREATE TABLE IF NOT EXISTS overtime (id SERIAL PRIMARY KEY, date TEXT NOT NULL, driver_id INTEGER NOT NULL, morning_hours REAL DEFAULT 0, evening_hours REAL DEFAULT 0, notes TEXT)""")
        else:
            cursor.execute("""CREATE TABLE IF NOT EXISTS cars (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, plate_number TEXT, car_type TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS drivers (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, phone TEXT, notes TEXT, active INTEGER DEFAULT 1)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS trips (id INTEGER PRIMARY KEY AUTOINCREMENT, trip_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER NOT NULL, destination TEXT, start_location TEXT, trip_type TEXT, requester TEXT, exit_time TEXT, return_time TEXT, start_km REAL, end_km REAL, distance REAL, notes TEXT)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS fuel (id INTEGER PRIMARY KEY AUTOINCREMENT, fuel_date TEXT NOT NULL, car_id INTEGER NOT NULL, driver_id INTEGER, fuel_type TEXT, source TEXT, payment_status TEXT, liters REAL, price_per_liter REAL, total_cost REAL, final_cost REAL, km_at_fill REAL, notes TEXT)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS maintenance (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, car_id INTEGER NOT NULL, type TEXT, category TEXT, km_at_service REAL, cost REAL, invoice_number TEXT, invoice_image TEXT, workshop TEXT, notes TEXT, next_service_km REAL, next_service_date TEXT)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS travel_missions (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, driver_id INTEGER NOT NULL, description TEXT, distance_km REAL DEFAULT 0, is_travel INTEGER DEFAULT 0, departure_time TEXT, return_time TEXT, notes TEXT)""")
            cursor.execute("""CREATE TABLE IF NOT EXISTS overtime (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL, driver_id INTEGER NOT NULL, morning_hours REAL DEFAULT 0, evening_hours REAL DEFAULT 0, notes TEXT)""")
        conn.commit()
        conn.close()
        print(f"✓ تم إنشاء/التحقق من كل الجداول ({'PostgreSQL' if is_postgres() else 'SQLite'})")
        return True
    except Exception as e:
        print(f"⚠ خطأ في ensure_all_tables: {e}")
        return False


ensure_all_tables()
ensure_trips_columns()


@app.context_processor
def inject_permissions():
    return {'is_admin': is_admin(), 'is_viewer': is_viewer()}


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        password = request.form.get('password', '')
        login_type = request.form.get('login_type', 'viewer')
        if login_type == 'admin':
            if password == get_admin_password():
                session['role'] = 'admin'
                flash('مرحباً بك كمسؤول', 'success')
                return redirect(url_for('index'))
            else:
                flash('كلمة سر المسؤول غير صحيحة', 'error')
        else:
            if password == get_viewer_password():
                session['role'] = 'viewer'
                flash('مرحباً بك', 'success')
                return redirect(url_for('index'))
            else:
                flash('كلمة السر غير صحيحة', 'error')
    return render_template('login.html')


@app.route('/logout')
def logout():
    session.clear()
    flash('تم تسجيل الخروج', 'success')
    return redirect(url_for('login'))


@app.route('/lang/<lang>')
def set_lang(lang):
    session['lang'] = lang
    return redirect(request.referrer or '/')


@app.route('/settings/passwords', methods=['GET', 'POST'])
@login_required
def settings_passwords():
    if request.method == 'POST':
        current_password = request.form.get('current_password', '')
        new_admin = request.form.get('new_admin_password', '').strip()
        new_viewer = request.form.get('new_viewer_password', '').strip()
        passwords = load_passwords()
        if is_admin():
            if current_password != passwords.get('admin_password'):
                flash('كلمة السر الحالية غير صحيحة', 'error')
                return redirect(url_for('settings_passwords'))
            if new_admin and len(new_admin) < 6:
                flash('كلمة سر المسؤول الجديدة يجب أن تكون 6 أحرف على الأقل', 'error')
                return redirect(url_for('settings_passwords'))
            if new_viewer and len(new_viewer) < 6:
                flash('كلمة سر المشاهد الجديدة يجب أن تكون 6 أحرف على الأقل', 'error')
                return redirect(url_for('settings_passwords'))
            if new_admin:
                passwords['admin_password'] = new_admin
            if new_viewer:
                passwords['viewer_password'] = new_viewer
            save_passwords(passwords)
            flash('تم حفظ كلمات السر بنجاح. جاري إعادة تشغيل التطبيق...', 'success')
            return redirect(url_for('restart_app'))
        else:
            if current_password != passwords.get('viewer_password'):
                flash('كلمة السر الحالية غير صحيحة', 'error')
                return redirect(url_for('settings_passwords'))
            if not new_viewer:
                flash('الرجاء إدخال كلمة السر الجديدة', 'error')
                return redirect(url_for('settings_passwords'))
            if len(new_viewer) < 6:
                flash('كلمة السر الجديدة يجب أن تكون 6 أحرف على الأقل', 'error')
                return redirect(url_for('settings_passwords'))
            passwords['viewer_password'] = new_viewer
            save_passwords(passwords)
            flash('تم حفظ كلمة السر بنجاح. جاري إعادة تشغيل التطبيق...', 'success')
            return redirect(url_for('restart_app'))
    return render_template('settings_passwords.html')


@app.route('/restart')
@login_required
def restart_app():
    return render_template('restart.html')


@app.route('/restart/do')
@login_required
def do_restart():
    try:
        subprocess.Popen([sys.executable] + sys.argv, close_fds=True)
    except Exception:
        pass
    return jsonify({'status': 'ok'})


@app.route('/')
@login_required
def index():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT COUNT(*) as cnt FROM cars WHERE active=1")
    cars_count = cursor.fetchone()['cnt']
    execute_query(cursor, "SELECT COUNT(*) as cnt FROM drivers WHERE active=1")
    drivers_count = cursor.fetchone()['cnt']
    now = datetime.now()
    month_str = now.strftime('%Y-%m')
    execute_query(cursor, "SELECT COUNT(*) as cnt FROM trips WHERE trip_date LIKE ?", (month_str + '%',))
    trips_month = cursor.fetchone()['cnt']
    execute_query(cursor, "SELECT SUM(distance) as total FROM trips WHERE trip_date LIKE ?", (month_str + '%',))
    km_row = cursor.fetchone()
    km_month = km_row['total'] if km_row and km_row['total'] else 0
    execute_query(cursor, """
        SELECT t.*, c.name as car_name, d.name as driver_name
        FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id
        ORDER BY COALESCE(t.end_km, 0) DESC, t.trip_date DESC, 
                 COALESCE(t.exit_time, '00:00') DESC, t.id DESC
    """)
    all_recent_trips = cursor.fetchall()
    execute_query(cursor, """
        SELECT COUNT(*) as cnt FROM trips 
        WHERE trip_source = 'driver_app' AND (admin_seen IS NULL OR admin_seen = 0)
    """)
    pending_trips_count = cursor.fetchone()['cnt']
    conn.close()
    trips_by_car = {}
    for t in all_recent_trips:
        car_name = t['car_name']
        if car_name not in trips_by_car:
            trips_by_car[car_name] = []
        trips_by_car[car_name].append(t)
    trips_by_car = dict(sorted(trips_by_car.items()))
    stats = {'cars': cars_count, 'drivers': drivers_count, 'trips_month': trips_month, 'km_month': round(km_month, 1)}
    alerts = get_maintenance_alerts()
    return render_template('index.html', stats=stats, trips_by_car=trips_by_car, alerts=alerts,
                         pending_trips_count=pending_trips_count)


@app.route('/api/admin/pending_count')
@login_required
def api_admin_pending_count():
    try:
        conn = get_connection()
        cursor = conn.cursor()
        execute_query(cursor, """
            SELECT COUNT(*) as cnt FROM trips 
            WHERE trip_source = 'driver_app' AND (admin_seen IS NULL OR admin_seen = 0)
        """)
        count = cursor.fetchone()['cnt']
        conn.close()
        return jsonify({'success': True, 'count': count})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/admin/pending')
@login_required
def admin_pending_trips():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT t.*, c.name as car_name, d.name as driver_name
        FROM trips t 
        JOIN cars c ON t.car_id = c.id 
        JOIN drivers d ON t.driver_id = d.id
        WHERE t.trip_source = 'driver_app' 
          AND (t.admin_seen IS NULL OR t.admin_seen = 0)
        ORDER BY t.id DESC
    """)
    pending_trips = cursor.fetchall()
    conn.close()
    return render_template('admin_pending.html', pending_trips=pending_trips)


@app.route('/admin/pending/mark_seen/<int:trip_id>', methods=['POST'])
@login_required
def admin_mark_seen(trip_id):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        execute_query(cursor, "UPDATE trips SET admin_seen = 1, admin_seen_at = ? WHERE id = ?", (now, trip_id))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/admin/pending/mark_all_seen', methods=['POST'])
@login_required
def admin_mark_all_seen():
    try:
        conn = get_connection()
        cursor = conn.cursor()
        now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        execute_query(cursor, """
            UPDATE trips SET admin_seen = 1, admin_seen_at = ? 
            WHERE trip_source = 'driver_app' AND (admin_seen IS NULL OR admin_seen = 0)
        """, (now,))
        conn.commit()
        conn.close()
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/add_trip', methods=['GET', 'POST'])
@admin_required
def add_trip():
    conn = get_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        trip_date = request.form['trip_date']
        car_id = request.form['car_id']
        driver_id = request.form['driver_id']
        destination = request.form.get('destination', '')
        start_location = request.form.get('start_location', '')
        trip_type = request.form.get('trip_type', '')
        requester = request.form.get('requester', '')
        exit_time = request.form.get('exit_time', '')
        return_time = request.form.get('return_time', '')
        start_km = request.form.get('start_km') or None
        end_km = request.form.get('end_km') or None
        notes = request.form.get('notes', '')
        distance = None
        if start_km and end_km:
            try:
                distance = float(end_km) - float(start_km)
            except ValueError:
                distance = None
        execute_query(cursor, """
            INSERT INTO trips (trip_date, car_id, driver_id, destination, start_location, trip_type, requester, exit_time, return_time, start_km, end_km, distance, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (trip_date, car_id, driver_id, destination, start_location, trip_type, requester, exit_time, return_time, start_km, end_km, distance, notes))
        conn.commit()
        conn.close()
        flash('تم حفظ الرحلة بنجاح', 'success')
        return redirect(url_for('index'))
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name")
    cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    drivers_list = cursor.fetchall()
    conn.close()
    return render_template('add_trip.html', cars=cars_list, drivers=drivers_list)


@app.route('/trips/edit/<int:trip_id>', methods=['GET', 'POST'])
@admin_required
def edit_trip(trip_id):
    conn = get_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        trip_date = request.form['trip_date']
        car_id = request.form['car_id']
        driver_id = request.form['driver_id']
        destination = request.form.get('destination', '')
        start_location = request.form.get('start_location', '')
        trip_type = request.form.get('trip_type', '')
        requester = request.form.get('requester', '')
        exit_time = request.form.get('exit_time', '')
        return_time = request.form.get('return_time', '')
        start_km = request.form.get('start_km') or None
        end_km = request.form.get('end_km') or None
        notes = request.form.get('notes', '')
        distance = None
        if start_km and end_km:
            try:
                distance = float(end_km) - float(start_km)
            except ValueError:
                distance = None
        execute_query(cursor, """
            UPDATE trips SET trip_date=?, car_id=?, driver_id=?, destination=?, 
            start_location=?, trip_type=?, requester=?,
            exit_time=?, return_time=?, start_km=?, end_km=?, distance=?, notes=?
            WHERE id=?
        """, (trip_date, car_id, driver_id, destination, start_location, trip_type, requester, exit_time, return_time, start_km, end_km, distance, notes, trip_id))
        conn.commit()
        conn.close()
        flash('تم تعديل الرحلة', 'success')
        return redirect(url_for('index'))
    execute_query(cursor, "SELECT * FROM trips WHERE id=?", (trip_id,))
    trip = cursor.fetchone()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name")
    cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    drivers_list = cursor.fetchall()
    conn.close()
    return render_template('add_trip.html', cars=cars_list, drivers=drivers_list, edit_trip=trip)


@app.route('/trips/delete/<int:trip_id>')
@admin_required
def delete_trip(trip_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM trips WHERE id=?", (trip_id,))
    conn.commit()
    conn.close()
    flash('تم حذف الرحلة', 'success')
    return redirect(url_for('index'))


@app.route('/cars')
@login_required
def cars():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY id")
    cars_list = cursor.fetchall()
    conn.close()
    return render_template('cars.html', cars=cars_list)


@app.route('/cars/add', methods=['POST'])
@admin_required
def add_car():
    name = request.form['name']
    plate = request.form.get('plate_number', '')
    car_type = request.form.get('car_type', '')
    notes = request.form.get('notes', '')
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO cars (name, plate_number, car_type, notes) VALUES (?, ?, ?, ?)",
                   (name, plate, car_type, notes))
    conn.commit()
    conn.close()
    flash('تم إضافة السيارة', 'success')
    return redirect(url_for('cars'))


@app.route('/cars/edit/<int:car_id>', methods=['GET', 'POST'])
@admin_required
def edit_car(car_id):
    conn = get_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        name = request.form['name']
        plate = request.form.get('plate_number', '')
        car_type = request.form.get('car_type', '')
        notes = request.form.get('notes', '')
        execute_query(cursor, "UPDATE cars SET name=?, plate_number=?, car_type=?, notes=? WHERE id=?",
                       (name, plate, car_type, notes, car_id))
        conn.commit()
        conn.close()
        flash('تم تعديل السيارة', 'success')
        return redirect(url_for('cars'))
    execute_query(cursor, "SELECT * FROM cars WHERE id=?", (car_id,))
    car = cursor.fetchone()
    conn.close()
    return render_template('cars.html', cars=[], edit_car=car)


@app.route('/cars/delete/<int:car_id>')
@admin_required
def delete_car(car_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM cars WHERE id=?", (car_id,))
    conn.commit()
    conn.close()
    flash('تم حذف السيارة', 'success')
    return redirect(url_for('cars'))


@app.route('/drivers')
@login_required
def drivers():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY id")
    drivers_list = cursor.fetchall()
    execute_query(cursor, """
        SELECT tm.*, d.name as driver_name
        FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id
        ORDER BY tm.date DESC, tm.id DESC
    """)
    all_missions = cursor.fetchall()
    execute_query(cursor, """
        SELECT o.*, d.name as driver_name
        FROM overtime o JOIN drivers d ON o.driver_id = d.id
        ORDER BY o.date DESC, o.id DESC
    """)
    all_overtime = cursor.fetchall()
    missions_by_driver = {}
    for m in all_missions:
        did = m['driver_id']
        missions_by_driver.setdefault(did, []).append(dict(m))
    overtime_by_driver = {}
    for o in all_overtime:
        did = o['driver_id']
        overtime_by_driver.setdefault(did, []).append(dict(o))
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    active_drivers = cursor.fetchall()
    conn.close()
    return render_template('drivers.html',
                         drivers=drivers_list,
                         active_drivers=active_drivers,
                         missions_by_driver=missions_by_driver,
                         overtime_by_driver=overtime_by_driver)


@app.route('/drivers/add', methods=['POST'])
@admin_required
def add_driver():
    name = request.form['name']
    phone = request.form.get('phone', '')
    notes = request.form.get('notes', '')
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "INSERT INTO drivers (name, phone, notes) VALUES (?, ?, ?)", (name, phone, notes))
    conn.commit()
    conn.close()
    flash('تم إضافة السائق', 'success')
    return redirect(url_for('drivers'))


@app.route('/drivers/edit/<int:driver_id>', methods=['GET', 'POST'])
@admin_required
def edit_driver(driver_id):
    conn = get_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        name = request.form['name']
        phone = request.form.get('phone', '')
        notes = request.form.get('notes', '')
        execute_query(cursor, "UPDATE drivers SET name=?, phone=?, notes=? WHERE id=?",
                       (name, phone, notes, driver_id))
        conn.commit()
        conn.close()
        flash('تم تعديل السائق', 'success')
        return redirect(url_for('drivers'))
    execute_query(cursor, "SELECT * FROM drivers WHERE id=?", (driver_id,))
    driver = cursor.fetchone()
    conn.close()
    return render_template('drivers.html', drivers=[], edit_driver=driver)


@app.route('/drivers/delete/<int:driver_id>')
@admin_required
def delete_driver(driver_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM drivers WHERE id=?", (driver_id,))
    conn.commit()
    conn.close()
    flash('تم حذف السائق', 'success')
    return redirect(url_for('drivers'))


@app.route('/travel_missions')
@login_required
def travel_missions():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT tm.*, d.name as driver_name
        FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id
        ORDER BY tm.date DESC, tm.id DESC
    """)
    missions = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    drivers_list = cursor.fetchall()
    conn.close()
    return render_template('travel_missions.html', missions=missions, drivers=drivers_list)


@app.route('/travel_missions/add', methods=['POST'])
@admin_required
def add_travel_mission():
    date = request.form['date']
    driver_id = request.form['driver_id']
    description = request.form.get('description', '')
    distance_km = request.form.get('distance_km') or 0
    departure_time = request.form.get('departure_time', '')
    return_time = request.form.get('return_time', '')
    notes = request.form.get('notes', '')
    try:
        distance_km = float(distance_km)
    except (ValueError, TypeError):
        distance_km = 0
    is_travel = 1 if distance_km >= 75 else 0
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        INSERT INTO travel_missions (date, driver_id, description, distance_km, is_travel, departure_time, return_time, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (date, driver_id, description, distance_km, is_travel, departure_time, return_time, notes))
    conn.commit()
    conn.close()
    flash('تم حفظ مهمة السفر', 'success')
    return redirect(url_for('drivers'))


@app.route('/travel_missions/delete/<int:m_id>')
@admin_required
def delete_travel_mission(m_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM travel_missions WHERE id=?", (m_id,))
    conn.commit()
    conn.close()
    flash('تم حذف مهمة السفر', 'success')
    return redirect(url_for('drivers'))


@app.route('/overtime')
@login_required
def overtime():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT o.*, d.name as driver_name
        FROM overtime o JOIN drivers d ON o.driver_id = d.id
        ORDER BY o.date DESC, o.id DESC
    """)
    overtimes = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    drivers_list = cursor.fetchall()
    conn.close()
    return render_template('overtime.html', overtimes=overtimes, drivers=drivers_list)


@app.route('/overtime/add', methods=['POST'])
@admin_required
def add_overtime():
    date = request.form['date']
    driver_id = request.form['driver_id']
    morning_hours = request.form.get('morning_hours') or 0
    evening_hours = request.form.get('evening_hours') or 0
    notes = request.form.get('notes', '')
    try:
        morning_hours = float(morning_hours)
    except (ValueError, TypeError):
        morning_hours = 0
    try:
        evening_hours = float(evening_hours)
    except (ValueError, TypeError):
        evening_hours = 0
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        INSERT INTO overtime (date, driver_id, morning_hours, evening_hours, notes)
        VALUES (?, ?, ?, ?, ?)
    """, (date, driver_id, morning_hours, evening_hours, notes))
    conn.commit()
    conn.close()
    flash('تم حفظ الإضافي', 'success')
    return redirect(url_for('drivers'))


@app.route('/overtime/delete/<int:o_id>')
@admin_required
def delete_overtime(o_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM overtime WHERE id=?", (o_id,))
    conn.commit()
    conn.close()
    flash('تم حذف الإضافي', 'success')
    return redirect(url_for('drivers'))


@app.route('/fuel')
@login_required
def fuel():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT f.*, c.name as car_name, d.name as driver_name
        FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id
        ORDER BY f.id DESC LIMIT 50
    """)
    fuel_records = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name")
    cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    drivers_list = cursor.fetchall()
    conn.close()
    return render_template('fuel.html', fuel_records=fuel_records, cars=cars_list, drivers=drivers_list)


@app.route('/fuel/add', methods=['POST'])
@admin_required
def add_fuel():
    fuel_date = request.form['fuel_date']
    car_id = request.form['car_id']
    driver_id = request.form.get('driver_id') or None
    fuel_type = request.form.get('fuel_type', '')
    source = request.form.get('source', '')
    payment_status = request.form.get('payment_status', '')
    liters = request.form.get('liters') or None
    price = request.form.get('price_per_liter') or None
    total = request.form.get('total_cost') or None
    final = request.form.get('final_cost') or None
    km = request.form.get('km_at_fill') or None
    notes = request.form.get('notes', '')
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        INSERT INTO fuel (fuel_date, car_id, driver_id, fuel_type, source, payment_status, liters, price_per_liter, total_cost, final_cost, km_at_fill, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (fuel_date, car_id, driver_id, fuel_type, source, payment_status, liters, price, total, final, km, notes))
    conn.commit()
    conn.close()
    flash('تم حفظ تعبئة الوقود', 'success')
    return redirect(url_for('fuel'))


@app.route('/fuel/edit/<int:fuel_id>', methods=['GET', 'POST'])
@admin_required
def edit_fuel(fuel_id):
    conn = get_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        fuel_date = request.form['fuel_date']
        car_id = request.form['car_id']
        driver_id = request.form.get('driver_id') or None
        fuel_type = request.form.get('fuel_type', '')
        source = request.form.get('source', '')
        payment_status = request.form.get('payment_status', '')
        liters = request.form.get('liters') or None
        price = request.form.get('price_per_liter') or None
        total = request.form.get('total_cost') or None
        final = request.form.get('final_cost') or None
        km = request.form.get('km_at_fill') or None
        notes = request.form.get('notes', '')
        execute_query(cursor, """
            UPDATE fuel SET fuel_date=?, car_id=?, driver_id=?, fuel_type=?, source=?, payment_status=?,
            liters=?, price_per_liter=?, total_cost=?, final_cost=?, km_at_fill=?, notes=?
            WHERE id=?
        """, (fuel_date, car_id, driver_id, fuel_type, source, payment_status, liters, price, total, final, km, notes, fuel_id))
        conn.commit()
        conn.close()
        flash('تم تعديل تعبئة الوقود', 'success')
        return redirect(url_for('fuel'))
    execute_query(cursor, "SELECT * FROM fuel WHERE id=?", (fuel_id,))
    record = cursor.fetchone()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name")
    cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    drivers_list = cursor.fetchall()
    conn.close()
    return render_template('fuel.html', fuel_records=[], cars=cars_list, drivers=drivers_list, edit_fuel=record)


@app.route('/fuel/delete/<int:fuel_id>')
@admin_required
def delete_fuel(fuel_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM fuel WHERE id=?", (fuel_id,))
    conn.commit()
    conn.close()
    flash('تم حذف التعبئة', 'success')
    return redirect(url_for('fuel'))


@app.route('/maintenance')
@login_required
def maintenance():
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        SELECT m.*, c.name as car_name
        FROM maintenance m JOIN cars c ON m.car_id = c.id
        ORDER BY m.id DESC LIMIT 100
    """)
    records = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name")
    cars_list = cursor.fetchall()
    conn.close()
    return render_template('maintenance.html', records=records, cars=cars_list)


@app.route('/maintenance/add', methods=['POST'])
@admin_required
def add_maintenance():
    date = request.form['date']
    car_id = request.form['car_id']
    mtype = request.form['type']
    category = request.form.get('category', 'دورية')
    km = request.form.get('km_at_service') or None
    cost = request.form.get('cost') or None
    invoice_number = request.form.get('invoice_number', '')
    workshop = request.form.get('workshop', '')
    notes = request.form.get('notes', '')
    next_km = request.form.get('next_service_km') or None
    next_date = request.form.get('next_service_date') or None
    invoice_image = None
    if 'invoice_image' in request.files:
        file = request.files['invoice_image']
        if file and file.filename and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_')
            filename = timestamp + filename
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            invoice_image = filename
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, """
        INSERT INTO maintenance (date, car_id, type, category, km_at_service, cost, 
        invoice_number, invoice_image, workshop, notes, next_service_km, next_service_date)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (date, car_id, mtype, category, km, cost, invoice_number, invoice_image, workshop, notes, next_km, next_date))
    conn.commit()
    conn.close()
    flash('تم حفظ الصيانة', 'success')
    return redirect(url_for('maintenance'))


@app.route('/maintenance/edit/<int:m_id>', methods=['GET', 'POST'])
@admin_required
def edit_maintenance(m_id):
    conn = get_connection()
    cursor = conn.cursor()
    if request.method == 'POST':
        date = request.form['date']
        car_id = request.form['car_id']
        mtype = request.form['type']
        category = request.form.get('category', 'دورية')
        km = request.form.get('km_at_service') or None
        cost = request.form.get('cost') or None
        invoice_number = request.form.get('invoice_number', '')
        workshop = request.form.get('workshop', '')
        notes = request.form.get('notes', '')
        next_km = request.form.get('next_service_km') or None
        next_date = request.form.get('next_service_date') or None
        execute_query(cursor, "SELECT invoice_image FROM maintenance WHERE id=?", (m_id,))
        old = cursor.fetchone()
        invoice_image = old['invoice_image'] if old else None
        if 'invoice_image' in request.files:
            file = request.files['invoice_image']
            if file and file.filename and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S_')
                filename = timestamp + filename
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                invoice_image = filename
        execute_query(cursor, """
            UPDATE maintenance SET date=?, car_id=?, type=?, category=?, km_at_service=?, 
            cost=?, invoice_number=?, invoice_image=?, workshop=?, notes=?, 
            next_service_km=?, next_service_date=?
            WHERE id=?
        """, (date, car_id, mtype, category, km, cost, invoice_number, invoice_image, workshop, notes, next_km, next_date, m_id))
        conn.commit()
        conn.close()
        flash('تم تعديل الصيانة', 'success')
        return redirect(url_for('maintenance'))
    execute_query(cursor, "SELECT * FROM maintenance WHERE id=?", (m_id,))
    record = cursor.fetchone()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name")
    cars_list = cursor.fetchall()
    conn.close()
    return render_template('maintenance.html', records=[], cars=cars_list, edit_maintenance=record)


@app.route('/maintenance/delete/<int:m_id>')
@admin_required
def delete_maintenance(m_id):
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "DELETE FROM maintenance WHERE id=?", (m_id,))
    conn.commit()
    conn.close()
    flash('تم حذف الصيانة', 'success')
    return redirect(url_for('maintenance'))


@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)


@app.route('/reports')
@login_required
def reports():
    month = int(request.args.get('month', datetime.now().month))
    year = int(request.args.get('year', datetime.now().year))
    car_id = request.args.get('car_id') or ''
    driver_id = request.args.get('driver_id') or ''
    report_type = request.args.get('report_type') or 'all'
    payment_filter = request.args.get('payment_filter') or ''
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars WHERE active=1 ORDER BY name")
    cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers WHERE active=1 ORDER BY name")
    drivers_list = cursor.fetchall()
    years = list(range(2024, datetime.now().year + 2))
    report_data = None
    if request.args.get('month'):
        month_str = f"{year:04d}-{month:02d}"
        query_trips = """
            SELECT t.*, c.name as car_name, d.name as driver_name
            FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id
            WHERE t.trip_date LIKE ?
        """
        params = [month_str + '%']
        if car_id:
            query_trips += " AND t.car_id = ?"
            params.append(car_id)
        if driver_id:
            query_trips += " AND t.driver_id = ?"
            params.append(driver_id)
        query_trips += """ ORDER BY COALESCE(t.end_km, 0) DESC, t.trip_date DESC, 
                          COALESCE(t.exit_time, '00:00') DESC, t.id DESC"""
        execute_query(cursor, query_trips, params)
        all_trips = cursor.fetchall()
        query_fuel = """
            SELECT f.*, c.name as car_name, d.name as driver_name
            FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id
            WHERE f.fuel_date LIKE ?
        """
        fuel_params = [month_str + '%']
        if car_id:
            query_fuel += " AND f.car_id = ?"
            fuel_params.append(car_id)
        if driver_id:
            query_fuel += " AND f.driver_id = ?"
            fuel_params.append(driver_id)
        if payment_filter:
            if payment_filter == 'خزان الشركة':
                query_fuel += " AND f.source = ?"
                fuel_params.append('خزان الشركة')
            elif payment_filter == 'كازية':
                query_fuel += " AND f.source = ?"
                fuel_params.append('كازية')
            else:
                query_fuel += " AND f.payment_status = ?"
                fuel_params.append(payment_filter)
        execute_query(cursor, query_fuel, fuel_params)
        all_fuel = cursor.fetchall()
        query_maint = """
            SELECT m.*, c.name as car_name
            FROM maintenance m JOIN cars c ON m.car_id = c.id
            WHERE m.date LIKE ?
        """
        maint_params = [month_str + '%']
        if car_id:
            query_maint += " AND m.car_id = ?"
            maint_params.append(car_id)
        execute_query(cursor, query_maint, maint_params)
        all_maint = cursor.fetchall()
        query_travel = """
            SELECT tm.*, d.name as driver_name
            FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id
            WHERE tm.date LIKE ?
        """
        travel_params = [month_str + '%']
        if driver_id:
            query_travel += " AND tm.driver_id = ?"
            travel_params.append(driver_id)
        execute_query(cursor, query_travel, travel_params)
        all_travel = cursor.fetchall()
        query_overtime = """
            SELECT o.*, d.name as driver_name
            FROM overtime o JOIN drivers d ON o.driver_id = d.id
            WHERE o.date LIKE ?
        """
        overtime_params = [month_str + '%']
        if driver_id:
            query_overtime += " AND o.driver_id = ?"
            overtime_params.append(driver_id)
        execute_query(cursor, query_overtime, overtime_params)
        all_overtime = cursor.fetchall()
        cars_report = []
        total_trips = len(all_trips)
        total_km = sum(t['distance'] or 0 for t in all_trips)
        total_liters = sum(f['liters'] or 0 for f in all_fuel)
        gasoline_liters = sum(f['liters'] or 0 for f in all_fuel if f['fuel_type'] == 'بنزين')
        diesel_liters = sum(f['liters'] or 0 for f in all_fuel if f['fuel_type'] == 'ديزل')
        gasoline_cost = sum(get_effective_cost(f) for f in all_fuel if f['fuel_type'] == 'بنزين')
        diesel_cost = sum(get_effective_cost(f) for f in all_fuel if f['fuel_type'] == 'ديزل')
        tank_liters = sum(f['liters'] or 0 for f in all_fuel if f['source'] == 'خزان الشركة')
        station_liters = sum(f['liters'] or 0 for f in all_fuel if f['source'] == 'كازية')
        gateway_liters = sum(f['liters'] or 0 for f in all_fuel if f['payment_status'] == 'اشتراك شركة البوابة الذهبية')
        total_cost = 0
        for f in all_fuel:
            total_cost += get_effective_cost(f)
        total_maint_cost = sum(m['cost'] or 0 for m in all_maint)
        if car_id:
            car_list = [c for c in cars_list if str(c['id']) == str(car_id)]
        else:
            car_ids_with_data = set(t['car_id'] for t in all_trips) | set(f['car_id'] for f in all_fuel) | set(m['car_id'] for m in all_maint)
            car_list = [c for c in cars_list if c['id'] in car_ids_with_data]
        for c in car_list:
            c_trips = [t for t in all_trips if t['car_id'] == c['id']]
            c_fuel = [f for f in all_fuel if f['car_id'] == c['id']]
            c_maint = [m for m in all_maint if m['car_id'] == c['id']]
            car_fuel_cost = 0
            for f in c_fuel:
                car_fuel_cost += get_effective_cost(f)
            c_gasoline = sum(f['liters'] or 0 for f in c_fuel if f['fuel_type'] == 'بنزين')
            c_diesel = sum(f['liters'] or 0 for f in c_fuel if f['fuel_type'] == 'ديزل')
            c_gasoline_cost = sum(get_effective_cost(f) for f in c_fuel if f['fuel_type'] == 'بنزين')
            c_diesel_cost = sum(get_effective_cost(f) for f in c_fuel if f['fuel_type'] == 'ديزل')
            c_tank = sum(f['liters'] or 0 for f in c_fuel if f['source'] == 'خزان الشركة')
            c_station = sum(f['liters'] or 0 for f in c_fuel if f['source'] == 'كازية')
            c_gateway = sum(f['liters'] or 0 for f in c_fuel if f['payment_status'] == 'اشتراك شركة البوابة الذهبية')
            c_total_km = sum(t['distance'] or 0 for t in c_trips)
            c_cost_per_km = get_car_cost_per_km(c['id'], conn)
            c_liters_per_100 = get_car_liters_per_100km(c['id'], conn)
            c_trips_with_cost = []
            for t in c_trips:
                t_dict = dict(t)
                t_dict['cost_per_km'] = c_cost_per_km
                t_dict['trip_cost'] = (t['distance'] or 0) * c_cost_per_km
                c_trips_with_cost.append(t_dict)
            cars_report.append({
                'car_name': c['name'], 'trips': c_trips_with_cost, 'fuel': c_fuel, 'maintenance': c_maint,
                'total_km': c_total_km,
                'total_liters': sum(f['liters'] or 0 for f in c_fuel),
                'total_fuel_cost': car_fuel_cost,
                'total_maint_cost': sum(m['cost'] or 0 for m in c_maint),
                'gasoline_liters': c_gasoline, 'diesel_liters': c_diesel,
                'gasoline_cost': c_gasoline_cost, 'diesel_cost': c_diesel_cost,
                'tank_liters': c_tank, 'station_liters': c_station,
                'gateway_liters': c_gateway,
                'cost_per_km': c_cost_per_km,
                'liters_per_100': c_liters_per_100,
            })
        drivers_report = []
        if driver_id:
            drivers_to_report = [d for d in drivers_list if str(d['id']) == str(driver_id)]
        else:
            driver_ids_with_data = set(tm['driver_id'] for tm in all_travel) | set(o['driver_id'] for o in all_overtime) | set(t['driver_id'] for t in all_trips)
            drivers_to_report = [d for d in drivers_list if d['id'] in driver_ids_with_data]
        for d in drivers_to_report:
            d_travel = [tm for tm in all_travel if tm['driver_id'] == d['id']]
            d_overtime = [o for o in all_overtime if o['driver_id'] == d['id']]
            d_trips = [t for t in all_trips if t['driver_id'] == d['id']]
            d_travel_count = len([tm for tm in d_travel if tm['is_travel'] == 1])
            d_travel_km = sum(tm['distance_km'] or 0 for tm in d_travel if tm['is_travel'] == 1)
            d_morning = sum(o['morning_hours'] or 0 for o in d_overtime)
            d_evening = sum(o['evening_hours'] or 0 for o in d_overtime)
            drivers_report.append({
                'driver_name': d['name'],
                'travel_missions': d_travel,
                'overtimes': d_overtime,
                'trips': d_trips,
                'travel_count': d_travel_count,
                'travel_km': d_travel_km,
                'morning_hours': d_morning,
                'evening_hours': d_evening,
                'total_overtime': d_morning + d_evening,
                'trips_count': len(d_trips),
            })
        report_data = {
            'summary': {
                'total_trips': total_trips, 'total_km': total_km,
                'total_liters': total_liters, 'total_fuel_cost': total_cost,
                'total_maint_cost': total_maint_cost,
                'gasoline_liters': gasoline_liters, 'diesel_liters': diesel_liters,
                'gasoline_cost': gasoline_cost, 'diesel_cost': diesel_cost,
                'tank_liters': tank_liters, 'station_liters': station_liters,
                'gateway_liters': gateway_liters,
            },
            'cars': cars_report,
            'drivers': drivers_report,
        }
    conn.close()
    return render_template('reports.html',
                         cars=cars_list, drivers=drivers_list,
                         years=years, selected_month=month, selected_year=year,
                         selected_car=car_id, selected_driver=driver_id,
                         report_data=report_data, report_type=report_type,
                         payment_filter=payment_filter)


@app.route('/reports/drivers')
@login_required
def reports_drivers():
    driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY name")
    drivers_list = cursor.fetchall()
    query_travel = """
        SELECT tm.*, d.name as driver_name
        FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id
        WHERE 1=1
    """
    travel_params = []
    if driver_id:
        query_travel += " AND tm.driver_id = ?"
        travel_params.append(driver_id)
    if date_from:
        query_travel += " AND tm.date >= ?"
        travel_params.append(date_from)
    if date_to:
        query_travel += " AND tm.date <= ?"
        travel_params.append(date_to)
    query_travel += " ORDER BY tm.date DESC, tm.id DESC"
    execute_query(cursor, query_travel, travel_params)
    all_travel = cursor.fetchall()
    query_overtime = """
        SELECT o.*, d.name as driver_name
        FROM overtime o JOIN drivers d ON o.driver_id = d.id
        WHERE 1=1
    """
    overtime_params = []
    if driver_id:
        query_overtime += " AND o.driver_id = ?"
        overtime_params.append(driver_id)
    if date_from:
        query_overtime += " AND o.date >= ?"
        overtime_params.append(date_from)
    if date_to:
        query_overtime += " AND o.date <= ?"
        overtime_params.append(date_to)
    query_overtime += " ORDER BY o.date DESC, o.id DESC"
    execute_query(cursor, query_overtime, overtime_params)
    all_overtime = cursor.fetchall()
    conn.close()
    if driver_id:
        drivers_to_report = [d for d in drivers_list if str(d['id']) == str(driver_id)]
    else:
        driver_ids_with_data = set(tm['driver_id'] for tm in all_travel) | set(o['driver_id'] for o in all_overtime)
        drivers_to_report = [d for d in drivers_list if d['id'] in driver_ids_with_data]
    drivers_report = []
    for d in drivers_to_report:
        d_travel = [tm for tm in all_travel if tm['driver_id'] == d['id']]
        d_overtime = [o for o in all_overtime if o['driver_id'] == d['id']]
        d_travel_count = len([tm for tm in d_travel if tm['is_travel'] == 1])
        d_travel_km = sum(tm['distance_km'] or 0 for tm in d_travel if tm['is_travel'] == 1)
        d_morning = sum(o['morning_hours'] or 0 for o in d_overtime)
        d_evening = sum(o['evening_hours'] or 0 for o in d_overtime)
        drivers_report.append({
            'driver_name': d['name'],
            'travel_missions': d_travel,
            'overtimes': d_overtime,
            'travel_count': d_travel_count,
            'travel_km': d_travel_km,
            'morning_hours': d_morning,
            'evening_hours': d_evening,
            'total_overtime': d_morning + d_evening,
        })
    drivers_by_name = {dr['driver_name']: dr for dr in drivers_report}
    drivers_by_name = dict(sorted(drivers_by_name.items()))
    return render_template('reports_drivers.html',
                         drivers=drivers_list,
                         drivers_report=drivers_report,
                         drivers_by_name=drivers_by_name,
                         selected_driver=driver_id,
                         date_from=date_from, date_to=date_to,
                         has_filter=bool(driver_id or date_from or date_to))


@app.route('/reports/trips')
@login_required
def reports_trips():
    car_id = request.args.get('car_id') or ''
    driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY name")
    cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY name")
    drivers_list = cursor.fetchall()
    query = """
        SELECT t.*, c.name as car_name, d.name as driver_name
        FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id
        WHERE 1=1
    """
    params = []
    if car_id:
        query += " AND t.car_id = ?"
        params.append(car_id)
    if driver_id:
        query += " AND t.driver_id = ?"
        params.append(driver_id)
    if date_from:
        query += " AND t.trip_date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND t.trip_date <= ?"
        params.append(date_to)
    query += """ ORDER BY COALESCE(t.end_km, 0) DESC, t.trip_date DESC, 
                 COALESCE(t.exit_time, '00:00') DESC, t.id DESC"""
    execute_query(cursor, query, params)
    trips = cursor.fetchall()
    total_trips = len(trips)
    total_km = sum(t['distance'] or 0 for t in trips)
    by_car = {}
    for t in trips:
        if t['car_name'] not in by_car:
            by_car[t['car_name']] = {'count': 0, 'km': 0}
        by_car[t['car_name']]['count'] += 1
        by_car[t['car_name']]['km'] += t['distance'] or 0
    cost_per_km_map = {}
    execute_query(cursor, "SELECT DISTINCT car_id FROM trips")
    for row in cursor.fetchall():
        cost_per_km_map[row['car_id']] = get_car_cost_per_km(row['car_id'], conn)
    trips_with_cost = []
    for t in trips:
        t_dict = dict(t)
        t_dict['cost_per_km'] = cost_per_km_map.get(t['car_id'], 0)
        t_dict['trip_cost'] = (t['distance'] or 0) * t_dict['cost_per_km']
        trips_with_cost.append(t_dict)
    trips_by_car = {}
    for t in trips_with_cost:
        car_name = t['car_name']
        if car_name not in trips_by_car:
            trips_by_car[car_name] = []
        trips_by_car[car_name].append(t)
    trips_by_car = dict(sorted(trips_by_car.items()))
    conn.close()
    summary = {'total_trips': total_trips, 'total_km': total_km, 'by_car': by_car}
    return render_template('reports_trips.html',
                         cars=cars_list, drivers=drivers_list,
                         trips=trips_with_cost, trips_by_car=trips_by_car,
                         summary=summary,
                         selected_car=car_id, selected_driver=driver_id,
                         date_from=date_from, date_to=date_to,
                         has_filter=bool(car_id or driver_id or date_from or date_to))


@app.route('/reports/fuel')
@login_required
def reports_fuel():
    car_id = request.args.get('car_id') or ''
    driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    payment_filter = request.args.get('payment_filter') or ''
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY name")
    cars_list = cursor.fetchall()
    execute_query(cursor, "SELECT * FROM drivers ORDER BY name")
    drivers_list = cursor.fetchall()
    query = """
        SELECT f.*, c.name as car_name, d.name as driver_name
        FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id
        WHERE 1=1
    """
    params = []
    if car_id:
        query += " AND f.car_id = ?"
        params.append(car_id)
    if driver_id:
        query += " AND f.driver_id = ?"
        params.append(driver_id)
    if date_from:
        query += " AND f.fuel_date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND f.fuel_date <= ?"
        params.append(date_to)
    if payment_filter:
        if payment_filter == 'خزان الشركة':
            query += " AND f.source = ?"
            params.append('خزان الشركة')
        elif payment_filter == 'كازية':
            query += " AND f.source = ?"
            params.append('كازية')
        else:
            query += " AND f.payment_status = ?"
            params.append(payment_filter)
    query += " ORDER BY f.fuel_date DESC, f.km_at_fill DESC, f.id DESC"
    execute_query(cursor, query, params)
    fuel_records = cursor.fetchall()
    total_records = len(fuel_records)
    total_liters = sum(f['liters'] or 0 for f in fuel_records)
    gasoline_liters = sum(f['liters'] or 0 for f in fuel_records if f['fuel_type'] == 'بنزين')
    diesel_liters = sum(f['liters'] or 0 for f in fuel_records if f['fuel_type'] == 'ديزل')
    gasoline_cost = sum(get_effective_cost(f) for f in fuel_records if f['fuel_type'] == 'بنزين')
    diesel_cost = sum(get_effective_cost(f) for f in fuel_records if f['fuel_type'] == 'ديزل')
    tank_liters = sum(f['liters'] or 0 for f in fuel_records if f['source'] == 'خزان الشركة')
    station_liters = sum(f['liters'] or 0 for f in fuel_records if f['source'] == 'كازية')
    gateway_liters = sum(f['liters'] or 0 for f in fuel_records if f['payment_status'] == 'اشتراك شركة البوابة الذهبية')
    total_cost = 0
    for f in fuel_records:
        total_cost += get_effective_cost(f)
    avg_price = total_cost / total_liters if total_liters > 0 else 0
    by_car = {}
    for f in fuel_records:
        car_name = f['car_name']
        if car_name not in by_car:
            by_car[car_name] = {'count': 0, 'liters': 0, 'cost': 0, 'cost_estimated': 0, 'car_id': f['car_id'], 'km': 0, 'cost_per_km': 0, 'liters_per_100': 0}
        by_car[car_name]['count'] += 1
        by_car[car_name]['liters'] += f['liters'] or 0
        by_car[car_name]['cost'] += get_effective_cost(f)
        by_car[car_name]['cost_estimated'] += get_estimated_cost(f)
    for car_name in by_car:
        cid = by_car[car_name]['car_id']
        execute_query(cursor, "SELECT COALESCE(SUM(distance), 0) as total_km FROM trips WHERE car_id=?", (cid,))
        row = cursor.fetchone()
        km = row['total_km'] if row else 0
        by_car[car_name]['km'] = km
        by_car[car_name]['cost_per_km'] = get_car_cost_per_km(cid, conn)
        by_car[car_name]['liters_per_100'] = get_car_liters_per_100km(cid, conn)
    conn.close()
    summary = {
        'total_records': total_records, 'total_liters': total_liters,
        'total_cost': total_cost, 'avg_price': avg_price, 'by_car': by_car,
        'gasoline_liters': gasoline_liters, 'diesel_liters': diesel_liters,
        'gasoline_cost': gasoline_cost, 'diesel_cost': diesel_cost,
        'tank_liters': tank_liters, 'station_liters': station_liters,
        'gateway_liters': gateway_liters,
    }
    return render_template('reports_fuel.html',
                         cars=cars_list, drivers=drivers_list,
                         fuel_records=fuel_records, summary=summary,
                         selected_car=car_id, selected_driver=driver_id,
                         date_from=date_from, date_to=date_to,
                         payment_filter=payment_filter,
                         has_filter=bool(car_id or driver_id or date_from or date_to or payment_filter))


@app.route('/reports/maintenance')
@login_required
def reports_maintenance():
    car_id = request.args.get('car_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    mtype = request.args.get('type') or ''
    category = request.args.get('category') or ''
    conn = get_connection()
    cursor = conn.cursor()
    execute_query(cursor, "SELECT * FROM cars ORDER BY name")
    cars_list = cursor.fetchall()
    query = """
        SELECT m.*, c.name as car_name
        FROM maintenance m JOIN cars c ON m.car_id = c.id
        WHERE 1=1
    """
    params = []
    if car_id:
        query += " AND m.car_id = ?"
        params.append(car_id)
    if date_from:
        query += " AND m.date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND m.date <= ?"
        params.append(date_to)
    if mtype:
        query += " AND m.type = ?"
        params.append(mtype)
    if category:
        query += " AND m.category = ?"
        params.append(category)
    query += " ORDER BY m.date DESC, m.km_at_service DESC, m.id DESC"
    execute_query(cursor, query, params)
    records = cursor.fetchall()
    total_records = len(records)
    total_cost = sum(m['cost'] or 0 for m in records)
    by_category = {}
    for m in records:
        cat = m['category'] or 'غير محدد'
        if cat not in by_category:
            by_category[cat] = {'count': 0, 'cost': 0}
        by_category[cat]['count'] += 1
        by_category[cat]['cost'] += m['cost'] or 0
    by_car = {}
    for m in records:
        if m['car_name'] not in by_car:
            by_car[m['car_name']] = {'count': 0, 'cost': 0}
        by_car[m['car_name']]['count'] += 1
        by_car[m['car_name']]['cost'] += m['cost'] or 0
    conn.close()
    summary = {'total_records': total_records, 'total_cost': total_cost, 'by_category': by_category, 'by_car': by_car}
    return render_template('reports_maintenance.html',
                         cars=cars_list, records=records, summary=summary,
                         selected_car=car_id, date_from=date_from, date_to=date_to,
                         selected_type=mtype, selected_category=category,
                         has_filter=bool(car_id or date_from or date_to or mtype or category))


@app.route('/reports/export')
@login_required
def export_excel():
    month = int(request.args.get('month', datetime.now().month))
    year = int(request.args.get('year', datetime.now().year))
    car_id = request.args.get('car_id') or ''
    driver_id = request.args.get('driver_id') or ''
    report_type = request.args.get('report_type') or 'all'
    payment_filter = request.args.get('payment_filter') or ''
    month_str = f"{year:04d}-{month:02d}"
    conn = get_connection()
    cursor = conn.cursor()
    query = """
        SELECT t.*, c.name as car_name, d.name as driver_name
        FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id
        WHERE t.trip_date LIKE ?
    """
    params = [month_str + '%']
    if car_id:
        query += " AND t.car_id = ?"
        params.append(car_id)
    if driver_id:
        query += " AND t.driver_id = ?"
        params.append(driver_id)
    query += """ ORDER BY COALESCE(t.end_km, 0) DESC, t.trip_date DESC, 
                 COALESCE(t.exit_time, '00:00') DESC, t.id DESC"""
    execute_query(cursor, query, params)
    trips = cursor.fetchall()
    query_f = """
        SELECT f.*, c.name as car_name, d.name as driver_name
        FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id
        WHERE f.fuel_date LIKE ?
    """
    params_f = [month_str + '%']
    if car_id:
        query_f += " AND f.car_id = ?"
        params_f.append(car_id)
    if driver_id:
        query_f += " AND f.driver_id = ?"
        params_f.append(driver_id)
    if payment_filter:
        if payment_filter == 'خزان الشركة':
            query_f += " AND f.source = ?"
            params_f.append('خزان الشركة')
        elif payment_filter == 'كازية':
            query_f += " AND f.source = ?"
            params_f.append('كازية')
        else:
            query_f += " AND f.payment_status = ?"
            params_f.append(payment_filter)
    execute_query(cursor, query_f, params_f)
    fuels = cursor.fetchall()
    query_m = """
        SELECT m.*, c.name as car_name
        FROM maintenance m JOIN cars c ON m.car_id = c.id
        WHERE m.date LIKE ?
    """
    params_m = [month_str + '%']
    if car_id:
        query_m += " AND m.car_id = ?"
        params_m.append(car_id)
    execute_query(cursor, query_m, params_m)
    maints = cursor.fetchall()
    query_travel = """
        SELECT tm.*, d.name as driver_name
        FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id
        WHERE tm.date LIKE ?
    """
    travel_params = [month_str + '%']
    if driver_id:
        query_travel += " AND tm.driver_id = ?"
        travel_params.append(driver_id)
    execute_query(cursor, query_travel, travel_params)
    travels = cursor.fetchall()
    query_overtime = """
        SELECT o.*, d.name as driver_name
        FROM overtime o JOIN drivers d ON o.driver_id = d.id
        WHERE o.date LIKE ?
    """
    overtime_params = [month_str + '%']
    if driver_id:
        query_overtime += " AND o.driver_id = ?"
        overtime_params.append(driver_id)
    execute_query(cursor, query_overtime, overtime_params)
    overtimes = cursor.fetchall()
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = f"Trips {month_str}"
    ws.sheet_view.rightToLeft = True
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    headers = ['التاريخ', 'السيارة', 'السائق', 'الوجهة', 'بداية الانطلاق', 'نوع الرحلة', 'طالب الرحلة', 'وقت الخروج', 'وقت العودة', 'عداد البداية', 'عداد النهاية', 'المسافة (كم)', 'ملاحظات']
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for t in trips:
        ws.append([t['trip_date'], t['car_name'], t['driver_name'], t['destination'] or '',
                   t['start_location'] or '', t['trip_type'] or '', t['requester'] or '',
                   t['exit_time'] or '', t['return_time'] or '', t['start_km'] or '',
                   t['end_km'] or '', t['distance'] or '', t['notes'] or ''])
    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3
    ws2 = wb.create_sheet(f"Fuel {month_str}")
    ws2.sheet_view.rightToLeft = True
    headers2 = ['التاريخ', 'السيارة', 'السائق', 'نوع الوقود', 'المصدر', 'حالة الدفع', 'اللترات', 'سعر اللتر', 'الكلفة النهائية', 'العداد', 'ملاحظات']
    ws2.append(headers2)
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for f in fuels:
        if should_count_cost(f):
            cost_val = f['final_cost'] if f['final_cost'] else (f['total_cost'] or '')
        else:
            cost_val = ''
        ws2.append([f['fuel_date'], f['car_name'], f['driver_name'] or '',
                    f['fuel_type'] or '', f['source'] or '', f['payment_status'] or '',
                    f['liters'] or '', f['price_per_liter'] or '', cost_val,
                    f['km_at_fill'] or '', f['notes'] or ''])
    for col in ws2.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws2.column_dimensions[col[0].column_letter].width = max_length + 3
    ws3 = wb.create_sheet(f"Maintenance {month_str}")
    ws3.sheet_view.rightToLeft = True
    headers3 = ['التاريخ', 'السيارة', 'النوع', 'الفئة', 'العداد', 'التكلفة', 'رقم الفاتورة', 'الورشة', 'ملاحظات', 'الصيانة القادمة (كم)']
    ws3.append(headers3)
    for cell in ws3[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for m in maints:
        ws3.append([m['date'], m['car_name'], m['type'], m['category'] or '',
                    m['km_at_service'] or '', m['cost'] or '', m['invoice_number'] or '',
                    m['workshop'] or '', m['notes'] or '', m['next_service_km'] or ''])
    for col in ws3.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws3.column_dimensions[col[0].column_letter].width = max_length + 3
    ws4 = wb.create_sheet(f"Travel {month_str}")
    ws4.sheet_view.rightToLeft = True
    headers4 = ['التاريخ', 'السائق', 'الوصف', 'المسافة (كم)', 'مهمة سفر؟', 'وقت الذهاب', 'وقت العودة', 'ملاحظات']
    ws4.append(headers4)
    for cell in ws4[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for tm in travels:
        travel_flag = 'نعم' if tm['is_travel'] == 1 else 'لا'
        ws4.append([tm['date'], tm['driver_name'], tm['description'] or '',
                    tm['distance_km'] or 0, travel_flag,
                    tm['departure_time'] or '', tm['return_time'] or '', tm['notes'] or ''])
    for col in ws4.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws4.column_dimensions[col[0].column_letter].width = max_length + 3
    ws5 = wb.create_sheet(f"Overtime {month_str}")
    ws5.sheet_view.rightToLeft = True
    headers5 = ['التاريخ', 'السائق', 'ساعات صباحية', 'ساعات مسائية', 'الإجمالي', 'ملاحظات']
    ws5.append(headers5)
    for cell in ws5[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for o in overtimes:
        mh = o['morning_hours'] or 0
        eh = o['evening_hours'] or 0
        ws5.append([o['date'], o['driver_name'], mh, eh, mh + eh, o['notes'] or ''])
    for col in ws5.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws5.column_dimensions[col[0].column_letter].width = max_length + 3
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name=f'qamh_report_{month_str}.xlsx')


@app.route('/reports/trips/export')
@login_required
def export_trips_excel():
    car_id = request.args.get('car_id') or ''
    driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    conn = get_connection()
    cursor = conn.cursor()
    query = """
        SELECT t.*, c.name as car_name, d.name as driver_name
        FROM trips t JOIN cars c ON t.car_id = c.id JOIN drivers d ON t.driver_id = d.id
        WHERE 1=1
    """
    params = []
    if car_id:
        query += " AND t.car_id = ?"
        params.append(car_id)
    if driver_id:
        query += " AND t.driver_id = ?"
        params.append(driver_id)
    if date_from:
        query += " AND t.trip_date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND t.trip_date <= ?"
        params.append(date_to)
    query += """ ORDER BY COALESCE(t.end_km, 0) DESC, t.trip_date DESC, 
                 COALESCE(t.exit_time, '00:00') DESC, t.id DESC"""
    execute_query(cursor, query, params)
    trips = cursor.fetchall()
    cost_per_km_map = {}
    execute_query(cursor, "SELECT DISTINCT car_id FROM trips")
    for row in cursor.fetchall():
        cost_per_km_map[row['car_id']] = get_car_cost_per_km(row['car_id'], conn)
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "Trips Report"
    ws.sheet_view.rightToLeft = True
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    headers = ['التاريخ', 'السيارة', 'السائق', 'الوجهة', 'بداية الانطلاق', 'نوع الرحلة', 'طالب الرحلة', 'وقت الخروج', 'وقت العودة', 'عداد البداية', 'عداد النهاية', 'المسافة (كم)', 'تكلفة/كم', 'كلفة الرحلة (تقديري)', 'ملاحظات']
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for t in trips:
        cpk = round(cost_per_km_map.get(t['car_id'], 0), 0)
        trip_cost = round((t['distance'] or 0) * cpk, 0)
        ws.append([t['trip_date'], t['car_name'], t['driver_name'], t['destination'] or '',
                   t['start_location'] or '', t['trip_type'] or '', t['requester'] or '',
                   t['exit_time'] or '', t['return_time'] or '', t['start_km'] or '',
                   t['end_km'] or '', t['distance'] or '', cpk, trip_cost, t['notes'] or ''])
    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='trips_report.xlsx')


@app.route('/reports/fuel/export')
@login_required
def export_fuel_excel():
    car_id = request.args.get('car_id') or ''
    driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    payment_filter = request.args.get('payment_filter') or ''
    conn = get_connection()
    cursor = conn.cursor()
    query = """
        SELECT f.*, c.name as car_name, d.name as driver_name
        FROM fuel f JOIN cars c ON f.car_id = c.id LEFT JOIN drivers d ON f.driver_id = d.id
        WHERE 1=1
    """
    params = []
    if car_id:
        query += " AND f.car_id = ?"
        params.append(car_id)
    if driver_id:
        query += " AND f.driver_id = ?"
        params.append(driver_id)
    if date_from:
        query += " AND f.fuel_date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND f.fuel_date <= ?"
        params.append(date_to)
    if payment_filter:
        if payment_filter == 'خزان الشركة':
            query += " AND f.source = ?"
            params.append('خزان الشركة')
        elif payment_filter == 'كازية':
            query += " AND f.source = ?"
            params.append('كازية')
        else:
            query += " AND f.payment_status = ?"
            params.append(payment_filter)
    query += " ORDER BY f.fuel_date DESC, f.km_at_fill DESC, f.id DESC"
    execute_query(cursor, query, params)
    fuel_records = cursor.fetchall()
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "Fuel Report"
    ws.sheet_view.rightToLeft = True
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    headers = ['التاريخ', 'السيارة', 'السائق', 'نوع الوقود', 'المصدر', 'حالة الدفع', 'اللترات', 'سعر اللتر', 'الكلفة النهائية', 'العداد', 'ملاحظات']
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for f in fuel_records:
        if should_count_cost(f):
            cost_val = f['final_cost'] if f['final_cost'] else (f['total_cost'] or '')
        else:
            cost_val = ''
        ws.append([f['fuel_date'], f['car_name'], f['driver_name'] or '',
                   f['fuel_type'] or '', f['source'] or '', f['payment_status'] or '',
                   f['liters'] or '', f['price_per_liter'] or '', cost_val,
                   f['km_at_fill'] or '', f['notes'] or ''])
    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='fuel_report.xlsx')


@app.route('/reports/maintenance/export')
@login_required
def export_maintenance_excel():
    car_id = request.args.get('car_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    mtype = request.args.get('type') or ''
    category = request.args.get('category') or ''
    conn = get_connection()
    cursor = conn.cursor()
    query = """
        SELECT m.*, c.name as car_name
        FROM maintenance m JOIN cars c ON m.car_id = c.id
        WHERE 1=1
    """
    params = []
    if car_id:
        query += " AND m.car_id = ?"
        params.append(car_id)
    if date_from:
        query += " AND m.date >= ?"
        params.append(date_from)
    if date_to:
        query += " AND m.date <= ?"
        params.append(date_to)
    if mtype:
        query += " AND m.type = ?"
        params.append(mtype)
    if category:
        query += " AND m.category = ?"
        params.append(category)
    query += " ORDER BY m.date DESC, m.km_at_service DESC, m.id DESC"
    execute_query(cursor, query, params)
    records = cursor.fetchall()
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "Maintenance Report"
    ws.sheet_view.rightToLeft = True
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    headers = ['التاريخ', 'السيارة', 'النوع', 'الفئة', 'العداد', 'التكلفة', 'رقم الفاتورة', 'الورشة', 'ملاحظات', 'الصيانة القادمة (كم)', 'تاريخ الصيانة القادمة']
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for m in records:
        ws.append([m['date'], m['car_name'], m['type'], m['category'] or '',
                   m['km_at_service'] or '', m['cost'] or '', m['invoice_number'] or '',
                   m['workshop'] or '', m['notes'] or '', m['next_service_km'] or '',
                   m['next_service_date'] or ''])
    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='maintenance_report.xlsx')


@app.route('/reports/drivers/export')
@login_required
def export_drivers_excel():
    driver_id = request.args.get('driver_id') or ''
    date_from = request.args.get('date_from') or ''
    date_to = request.args.get('date_to') or ''
    conn = get_connection()
    cursor = conn.cursor()
    query_travel = """
        SELECT tm.*, d.name as driver_name
        FROM travel_missions tm JOIN drivers d ON tm.driver_id = d.id
        WHERE 1=1
    """
    travel_params = []
    if driver_id:
        query_travel += " AND tm.driver_id = ?"
        travel_params.append(driver_id)
    if date_from:
        query_travel += " AND tm.date >= ?"
        travel_params.append(date_from)
    if date_to:
        query_travel += " AND tm.date <= ?"
        travel_params.append(date_to)
    query_travel += " ORDER BY tm.date DESC, tm.id DESC"
    execute_query(cursor, query_travel, travel_params)
    travels = cursor.fetchall()
    query_overtime = """
        SELECT o.*, d.name as driver_name
        FROM overtime o JOIN drivers d ON o.driver_id = d.id
        WHERE 1=1
    """
    overtime_params = []
    if driver_id:
        query_overtime += " AND o.driver_id = ?"
        overtime_params.append(driver_id)
    if date_from:
        query_overtime += " AND o.date >= ?"
        overtime_params.append(date_from)
    if date_to:
        query_overtime += " AND o.date <= ?"
        overtime_params.append(date_to)
    query_overtime += " ORDER BY o.date DESC, o.id DESC"
    execute_query(cursor, query_overtime, overtime_params)
    overtimes = cursor.fetchall()
    conn.close()
    wb = Workbook()
    ws = wb.active
    ws.title = "Travel Missions"
    ws.sheet_view.rightToLeft = True
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=12)
    headers = ['التاريخ', 'السائق', 'الوصف', 'المسافة (كم)', 'مهمة سفر؟', 'وقت الذهاب', 'وقت العودة', 'ملاحظات']
    ws.append(headers)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for tm in travels:
        travel_flag = 'نعم' if tm['is_travel'] == 1 else 'لا'
        ws.append([tm['date'], tm['driver_name'], tm['description'] or '',
                   tm['distance_km'] or 0, travel_flag,
                   tm['departure_time'] or '', tm['return_time'] or '', tm['notes'] or ''])
    for col in ws.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = max_length + 3
    ws2 = wb.create_sheet("Overtime")
    ws2.sheet_view.rightToLeft = True
    headers2 = ['التاريخ', 'السائق', 'ساعات صباحية', 'ساعات مسائية', 'الإجمالي', 'ملاحظات']
    ws2.append(headers2)
    for cell in ws2[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
    for o in overtimes:
        mh = o['morning_hours'] or 0
        eh = o['evening_hours'] or 0
        ws2.append([o['date'], o['driver_name'], mh, eh, mh + eh, o['notes'] or ''])
    for col in ws2.columns:
        max_length = max(len(str(cell.value or '')) for cell in col)
        ws2.column_dimensions[col[0].column_letter].width = max_length + 3
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                     as_attachment=True, download_name='drivers_report.xlsx')


@app.route('/driver')
def driver_app():
    return render_template('driver_app.html')


@app.route('/api/driver/init', methods=['GET'])
def api_driver_init():
    try:
        conn = get_connection()
        cursor = conn.cursor()
        execute_query(cursor, "SELECT id, name, phone FROM drivers WHERE active=1 ORDER BY name")
        drivers_list = [dict(row) for row in cursor.fetchall()]
        execute_query(cursor, "SELECT id, name, plate_number FROM cars WHERE active=1 ORDER BY name")
        cars_list = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return jsonify({'success': True, 'drivers': drivers_list, 'cars': cars_list})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/driver/start_trip', methods=['POST'])
def api_driver_start_trip():
    try:
        data = request.get_json()
        driver_id = data.get('driver_id')
        car_id = data.get('car_id')
        destination = data.get('destination', '').strip()
        start_km = data.get('start_km')
        start_image_b64 = data.get('start_image', '')
        
        if not driver_id or not car_id:
            return jsonify({'success': False, 'error': 'الرجاء اختيار السائق والسيارة'}), 400
        if not destination:
            return jsonify({'success': False, 'error': 'الرجاء إدخال الوجهة'}), 400
        if start_km is None or str(start_km).strip() == '':
            return jsonify({'success': False, 'error': 'الرجاء إدخال قراءة العداد'}), 400
        try:
            start_km = float(start_km)
            if start_km <= 0:
                return jsonify({'success': False, 'error': 'قراءة العداد يجب أن تكون أكبر من صفر'}), 400
        except (ValueError, TypeError):
            return jsonify({'success': False, 'error': 'قراءة العداد يجب أن تكون رقماً صحيحاً'}), 400
        if not start_image_b64 or len(start_image_b64) < 100:
            return jsonify({'success': False, 'error': 'صورة العداد إلزامية'}), 400
        
        start_image_filename = None
        try:
            if ',' in start_image_b64:
                start_image_b64 = start_image_b64.split(',', 1)[1]
            img_data = base64.b64decode(start_image_b64)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            start_image_filename = f"trip_start_{driver_id}_{car_id}_{timestamp}.jpg"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], start_image_filename)
            with open(filepath, 'wb') as f:
                f.write(img_data)
        except Exception as e:
            print(f"خطأ في حفظ صورة البداية: {e}")
            return jsonify({'success': False, 'error': 'فشل حفظ صورة العداد'}), 500
        
        now = datetime.now()
        trip_date = now.strftime('%Y-%m-%d')
        start_time = now.strftime('%H:%M:%S')
        
        conn = get_connection()
        cursor = conn.cursor()
        execute_query(cursor, """
            INSERT INTO trips 
            (trip_date, car_id, driver_id, destination, start_km, start_time, 
             start_image, trip_source, driver_confirmed, admin_seen, notes)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'driver_app', 1, 0, ?)
        """, (trip_date, car_id, driver_id, destination, start_km, 
              start_time, start_image_filename, 'رحلة مسجلة من تطبيق السائق'))
        conn.commit()
        
        execute_query(cursor, "SELECT MAX(id) as max_id FROM trips")
        row = cursor.fetchone()
        trip_id = row['max_id'] if row else None
        conn.close()
        
        return jsonify({'success': True, 'trip_id': trip_id, 'message': 'تم بدء الرحلة بنجاح', 'start_time': start_time})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/driver/end_trip', methods=['POST'])
def api_driver_end_trip():
    try:
        data = request.get_json()
        trip_id = data.get('trip_id')
        end_km = data.get('end_km')
        end_image_b64 = data.get('end_image', '')
        notes = data.get('notes', '')
        
        if not trip_id:
            return jsonify({'success': False, 'error': 'رقم الرحلة مفقود'}), 400
        if end_km is None or str(end_km).strip() == '':
            return jsonify({'success': False, 'error': 'الرجاء إدخال قراءة العداد النهائية'}), 400
        try:
            end_km = float(end_km)
            if end_km <= 0:
                return jsonify({'success': False, 'error': 'قراءة العداد يجب أن تكون أكبر من صفر'}), 400
        except (ValueError, TypeError):
            return jsonify({'success': False, 'error': 'قراءة العداد يجب أن تكون رقماً صحيحاً'}), 400
        if not end_image_b64 or len(end_image_b64) < 100:
            return jsonify({'success': False, 'error': 'صورة العداد النهائية إلزامية'}), 400
        
        conn = get_connection()
        cursor = conn.cursor()
        execute_query(cursor, "SELECT start_km FROM trips WHERE id=?", (trip_id,))
        trip = cursor.fetchone()
        if not trip:
            conn.close()
            return jsonify({'success': False, 'error': 'الرحلة غير موجودة'}), 404
        
        start_km = trip['start_km']
        if start_km is not None:
            try:
                if end_km <= float(start_km):
                    conn.close()
                    return jsonify({'success': False, 'error': f'قراءة العداد النهائية ({end_km}) يجب أن تكون أكبر من قراءة البداية ({start_km})'}), 400
            except (ValueError, TypeError):
                pass
        
        distance = None
        if start_km is not None and end_km is not None:
            try:
                distance = float(end_km) - float(start_km)
                if distance < 0:
                    distance = None
            except (ValueError, TypeError):
                distance = None
        
        end_image_filename = None
        try:
            if ',' in end_image_b64:
                end_image_b64 = end_image_b64.split(',', 1)[1]
            img_data = base64.b64decode(end_image_b64)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            end_image_filename = f"trip_end_{trip_id}_{timestamp}.jpg"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], end_image_filename)
            with open(filepath, 'wb') as f:
                f.write(img_data)
        except Exception as e:
            print(f"خطأ في حفظ صورة النهاية: {e}")
            conn.close()
            return jsonify({'success': False, 'error': 'فشل حفظ صورة العداد'}), 500
        
        now = datetime.now()
        end_time = now.strftime('%H:%M:%S')
        
        execute_query(cursor, """
            UPDATE trips SET end_km=?, end_time=?, end_image=?, distance=?, notes=? WHERE id=?
        """, (end_km, end_time, end_image_filename, distance, notes, trip_id))
        
        conn.commit()
        conn.close()
        
        return jsonify({'success': True, 'message': 'تم إنهاء الرحلة بنجاح', 'end_time': end_time, 'distance': distance})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/driver/active_trip/<int:driver_id>', methods=['GET'])
def api_driver_active_trip(driver_id):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        execute_query(cursor, """
            SELECT t.*, c.name as car_name
            FROM trips t JOIN cars c ON t.car_id = c.id
            WHERE t.driver_id = ? AND t.trip_source = 'driver_app' AND t.end_km IS NULL
            ORDER BY t.id DESC LIMIT 1
        """, (driver_id,))
        trip = cursor.fetchone()
        conn.close()
        if trip:
            return jsonify({'success': True, 'has_active': True, 'trip': dict(trip)})
        return jsonify({'success': True, 'has_active': False})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/driver/recent_trips/<int:driver_id>', methods=['GET'])
def api_driver_recent_trips(driver_id):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        execute_query(cursor, """
            SELECT t.id, t.trip_date, t.destination, t.start_km, t.end_km, 
                   t.distance, t.start_time, t.end_time, c.name as car_name
            FROM trips t JOIN cars c ON t.car_id = c.id
            WHERE t.driver_id = ?
            ORDER BY t.id DESC LIMIT 5
        """, (driver_id,))
        trips = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return jsonify({'success': True, 'trips': trips})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


@app.route('/api/driver/pending_count/<int:driver_id>', methods=['GET'])
def api_driver_pending_count(driver_id):
    try:
        conn = get_connection()
        cursor = conn.cursor()
        execute_query(cursor, """
            SELECT COUNT(*) as cnt FROM trips
            WHERE driver_id = ? AND trip_source = 'driver_app' AND end_km IS NULL
        """, (driver_id,))
        count = cursor.fetchone()['cnt']
        conn.close()
        return jsonify({'success': True, 'count': count})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print("=" * 60)
    print(f"  قاعدة البيانات: {'PostgreSQL (Supabase)' if is_postgres() else 'SQLite (محلي)'}")
    print(f"  Driver App:  http://localhost:{port}/driver")
    print(f"  Dashboard:   http://localhost:{port}/")
    print("=" * 60)
    app.run(host='0.0.0.0', port=port, debug=False)
