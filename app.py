import os
import csv
import sqlite3
from io import StringIO
from functools import wraps
from flask import Flask, render_template, request, jsonify, session, Response
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = 'stock-secret-key-change-this'
DB_FILE = 'inventory.db'

def auto_import_excel():
    excel_file = 'NEW APRIL STOCK SHEET 2026.xlsx'
    if not os.path.exists(excel_file):
        files = [f for f in os.listdir('.') if f.endswith('.xlsx')]
        if files:
            excel_file = files[0]
        else:
            return

    try:
        import openpyxl
        wb = openpyxl.load_workbook(excel_file, data_only=True)
        sheet_name = 'Sheet2' if 'Sheet2' in wb.sheetnames else wb.sheetnames[0]
        sheet = wb[sheet_name]

        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        
        # Reset products table to reload fresh excel data
        cursor.execute("DELETE FROM products")
        
        last_brand = ""
        row_count = 0
        
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 2:
                continue
            
            raw_brand = str(row[0]).strip() if row[0] is not None and str(row[0]).strip() != 'None' else ""
            description = str(row[1]).strip() if row[1] is not None and str(row[1]).strip() != 'None' else ""
            
            if raw_brand and raw_brand != 'BRAND NAME':
                last_brand = raw_brand
            brand = last_brand
            
            if not description or description in ['DISCRIPTION', 'DESCRIPTION']:
                continue
                
            def clean_int(val):
                try:
                    return int(float(val))
                except:
                    return 0

            rtt = clean_int(row[2]) if len(row) > 2 else 0
            rin = clean_int(row[3]) if len(row) > 3 else 0
            rit = clean_int(row[4]) if len(row) > 4 else 0
            ge = clean_int(row[5]) if len(row) > 5 else 0
            total = clean_int(row[6]) if len(row) > 6 else (rtt + rin + rit + ge)
            actual_stock = clean_int(row[7]) if len(row) > 7 else 0

            cursor.execute("""
                INSERT INTO products (brand, description, rtt, rin, rit, ge, total, actual_stock)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (brand, description, rtt, rin, rit, ge, total, actual_stock))
            row_count += 1

        conn.commit()
        conn.close()
        print(f"Successfully loaded {row_count} items from Excel!")
    except Exception as e:
        print(f"Error during excel import: {e}")

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            brand TEXT NOT NULL,
            description TEXT NOT NULL,
            rtt INTEGER DEFAULT 0,
            rin INTEGER DEFAULT 0,
            rit INTEGER DEFAULT 0,
            ge INTEGER DEFAULT 0,
            total INTEGER DEFAULT 0,
            actual_stock INTEGER DEFAULT 0
        )
    ''')
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS audit_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
            username TEXT NOT NULL,
            action TEXT NOT NULL,
            details TEXT NOT NULL
        )
    ''')
    
    cursor.execute("SELECT COUNT(*) FROM users")
    if cursor.fetchone()[0] == 0:
        default_users = [
            ('master', generate_password_hash('master123'), 'master'),
            ('worker1', generate_password_hash('worker123'), 'worker'),
            ('worker2', generate_password_hash('worker123'), 'worker'),
            ('worker3', generate_password_hash('worker123'), 'worker')
        ]
        cursor.executemany("INSERT INTO users (username, password, role) VALUES (?, ?, ?)", default_users)
        
    conn.commit()

    cursor.execute("SELECT COUNT(*) FROM products")
    count = cursor.fetchone()[0]
    conn.close()
    
    if count == 0:
        auto_import_excel()

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if 'username' not in session:
            return jsonify({'error': 'Unauthorized'}), 401
        return f(*args, **kwargs)
    return decorated

def master_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if session.get('role') != 'master':
            return jsonify({'error': 'Master access required'}), 403
        return f(*args, **kwargs)
    return decorated

def log_action(username, action, details):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO audit_logs (username, action, details) VALUES (?, ?, ?)", (username, action, details))
    conn.commit()
    conn.close()

with app.app_context():
    init_db()

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    username = data.get('username')
    password = data.get('password')
    
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT password, role FROM users WHERE username = ?", (username,))
    user = cursor.fetchone()
    conn.close()
    
    if user and check_password_hash(user[0], password):
        session['username'] = username
        session['role'] = user[1]
        log_action(username, 'LOGIN', 'Logged into dashboard')
        return jsonify({'status': 'success', 'username': username, 'role': user[1]})
    
    return jsonify({'error': 'Invalid credentials'}), 400

@app.route('/api/logout', methods=['POST'])
def logout():
    if 'username' in session:
        log_action(session['username'], 'LOGOUT', 'Logged out of app')
    session.clear()
    return jsonify({'status': 'success'})

@app.route('/api/products', methods=['GET'])
@login_required
def get_products():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM products ORDER BY brand, description")
    products = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({'products': products, 'user_role': session.get('role')})

@app.route('/api/products/add', methods=['POST'])
@login_required
def add_product():
    data = request.json
    brand = data.get('brand', '').upper().strip()
    description = data.get('description', '').strip()
    actual_stock = int(data.get('actual_stock', 0))
    rtt = int(data.get('rtt', 0))
    rin = int(data.get('rin', 0))
    rit = int(data.get('rit', 0))
    ge = int(data.get('ge', 0))
    total = rtt + rin + rit + ge

    if not brand or not description:
        return jsonify({'error': 'Brand and Description are required'}), 400

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO products (brand, description, rtt, rin, rit, ge, total, actual_stock)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (brand, description, rtt, rin, rit, ge, total, actual_stock))
    conn.commit()
    conn.close()
    
    log_action(session['username'], 'ADD_PRODUCT', f"Added product '{description}' under brand '{brand}' with stock {actual_stock}")
    return jsonify({'status': 'success'})

@app.route('/api/products/adjust', methods=['POST'])
@login_required
def adjust_stock():
    data = request.json
    prod_id = data.get('product_id')
    amount = int(data.get('amount', 0))
    action_type = data.get('type')
    
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT brand, description, actual_stock FROM products WHERE id = ?", (prod_id,))
    prod = cursor.fetchone()
    
    if not prod:
        conn.close()
        return jsonify({'error': 'Product not found'}), 404
        
    current_stock = prod[2]
    new_stock = current_stock - amount if action_type == 'issue' else current_stock + amount
    if new_stock < 0:
        conn.close()
        return jsonify({'error': 'Stock cannot be negative'}), 400
        
    cursor.execute("UPDATE products SET actual_stock = ? WHERE id = ?", (new_stock, prod_id))
    conn.commit()
    conn.close()
    
    log_action(session['username'], action_type.upper(), f"{action_type.capitalize()}d {amount} units of {prod[0]} - {prod[1]} (New Stock: {new_stock})")
    return jsonify({'status': 'success', 'new_stock': new_stock})

@app.route('/api/products/delete/<int:prod_id>', methods=['DELETE'])
@master_required
def delete_product(prod_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT brand, description FROM products WHERE id = ?", (prod_id,))
    prod = cursor.fetchone()
    if prod:
        cursor.execute("DELETE FROM products WHERE id = ?", (prod_id,))
        conn.commit()
        log_action(session['username'], 'DELETE', f"Deleted product '{prod[1]}' under brand '{prod[0]}'")
    conn.close()
    return jsonify({'status': 'success'})

@app.route('/api/activity', methods=['GET'])
@master_required
def get_activity():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM audit_logs ORDER BY id DESC LIMIT 40")
    logs = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({'logs': logs})

@app.route('/api/export/csv', methods=['GET'])
@master_required
def export_csv():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT brand, description, rtt, rin, rit, ge, total, actual_stock FROM products ORDER BY brand, description")
    rows = cursor.fetchall()
    conn.close()

    si = StringIO()
    cw = csv.writer(si)
    cw.writerow(['BRAND', 'DESCRIPTION', 'RTT', 'RIN', 'RIT', 'GE', 'TOTAL', 'ACTUAL STOCK'])
    cw.writerows(rows)

    output = Response(si.getvalue(), mimetype='text/csv')
    output.headers["Content-Disposition"] = "attachment; filename=Stock_Report.csv"
    return output

@app.route('/api/admin/reload', methods=['GET'])
@master_required
def force_reload():
    auto_import_excel()
    return jsonify({'status': 'Excel data reloaded successfully!'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
