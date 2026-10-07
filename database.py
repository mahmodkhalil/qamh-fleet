import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(__file__), 'qamh_fleet.db')

def init_db():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    
    # جدول السيارات
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cars (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            plate_number TEXT,
            car_type TEXT,
            notes TEXT,
            active INTEGER DEFAULT 1
        )
    ''')
    
    # جدول السائقين
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS drivers (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            phone TEXT,
            notes TEXT,
            active INTEGER DEFAULT 1
        )
    ''')
    
    # جدول الرحلات
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS trips (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            trip_date TEXT NOT NULL,
            car_id INTEGER NOT NULL,
            driver_id INTEGER NOT NULL,
            destination TEXT,
            start_location TEXT,
            trip_type TEXT,
            requester TEXT,
            exit_time TEXT,
            return_time TEXT,
            start_km REAL,
            end_km REAL,
            distance REAL,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (car_id) REFERENCES cars(id),
            FOREIGN KEY (driver_id) REFERENCES drivers(id)
        )
    ''')
    
    # جدول تعبئات الوقود
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS fuel (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fuel_date TEXT NOT NULL,
            car_id INTEGER NOT NULL,
            driver_id INTEGER,
            fuel_type TEXT,
            source TEXT,
            payment_status TEXT,
            liters REAL,
            price_per_liter REAL,
            total_cost REAL,
            final_cost REAL,
            km_at_fill REAL,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (car_id) REFERENCES cars(id),
            FOREIGN KEY (driver_id) REFERENCES drivers(id)
        )
    ''')
    
    # جدول الصيانات
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS maintenance (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            car_id INTEGER NOT NULL,
            type TEXT NOT NULL,
            category TEXT DEFAULT 'دورية',
            km_at_service REAL,
            cost REAL,
            invoice_number TEXT,
            invoice_image TEXT,
            workshop TEXT,
            notes TEXT,
            next_service_km REAL,
            next_service_date TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (car_id) REFERENCES cars(id)
        )
    ''')
    
    # جدول مهمات السفر
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS travel_missions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            driver_id INTEGER NOT NULL,
            description TEXT,
            distance_km REAL,
            is_travel INTEGER DEFAULT 0,
            departure_time TEXT,
            return_time TEXT,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (driver_id) REFERENCES drivers(id)
        )
    ''')
    
    # جدول الإضافي
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS overtime (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date TEXT NOT NULL,
            driver_id INTEGER NOT NULL,
            morning_hours REAL DEFAULT 0,
            evening_hours REAL DEFAULT 0,
            notes TEXT,
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (driver_id) REFERENCES drivers(id)
        )
    ''')
    
    conn.commit()
    conn.close()
    print("OK - Database created successfully")

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

if __name__ == '__main__':
    init_db()