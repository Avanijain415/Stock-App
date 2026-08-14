import os
import sqlite3
from datetime import datetime
from flask import Flask, render_template, request, jsonify, session, Response
import openpyxl

app = Flask(__name__)
app.secret_key = 'super_secret_enterprise_key_2026'

DB_NAME = 'stock.db'
EXCEL_FILE = 'NEW APRIL STOCK SHEET 2026.xlsx'

def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
    cursor = conn.cursor()
    
    # 1. Users Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            role TEXT NOT NULL
        )
    ''')
    
    # 2. Comprehensive Products Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            hsn_code TEXT,
            country TEXT,
            brand TEXT,
            description TEXT NOT NULL,
            price REAL DEFAULT 0,
            tax REAL DEFAULT 0,
            price_after_tax REAL DEFAULT 0,
            mrp REAL DEFAULT 0,
            per_petti INTEGER DEFAULT 0,
            expiry_date TEXT,
            opening_stock INTEGER DEFAULT 0,
            actual_stock INTEGER DEFAULT 0,
            damage INTEGER DEFAULT 0,
            status TEXT DEFAULT 'Completed'
        )
    ''')
    
    # 3. Item Ledger & Activity Log Table
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS activity_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            product_id INTEGER,
            product_desc TEXT,
            username TEXT,
            action TEXT,
            quantity INTEGER,
            prev_stock INTEGER,
            new_stock INTEGER,
            details TEXT,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Default Users
    cursor.execute("INSERT OR IGNORE INTO users (username, password, role) VALUES ('master', 'admin123', 'master')")
    cursor.execute("INSERT OR IGNORE INTO users (username, password, role) VALUES ('worker1', 'worker123', 'worker')")
    
    conn.commit()
    conn.close()

def migrate_from_excel():
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute("SELECT COUNT(*) FROM products")
    if cursor.fetchone()[0] > 0:
        conn.close()
        return

    if not os.path.exists(EXCEL_FILE):
        conn.close()
        return

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    # Read Sheet1 (Master detailed sheet)
    sheet = wb['Sheet1'] if 'Sheet1' in wb.sheetnames else wb.active
    
    # Sheet1 Headers are in row 1
    # Row 2 onwards are products
    for row in range(2, sheet.max_row + 1):
        desc = sheet.cell(row=row, column=5).value  # Column E: DISCRIPTION
        if not desc:
            continue
            
        hsn = str(sheet.cell(row=row, column=2).value or '')
        country = str(sheet.cell(row=row, column=3).value or 'INDIA')
        brand = str(sheet.cell(row=row, column=4).value or 'GENERAL')
        
        try: price = float(sheet.cell(row=row, column=6).value or 0)
        except: price = 0.0
        try: tax = float(sheet.cell(row=row, column=7).value or 0)
        except: tax = 0.0
        try: price_after_tax = float(sheet.cell(row=row, column=8).value or 0)
        except: price_after_tax = 0.0
        try: mrp = float(sheet.cell(row=row, column=9).value or 0)
        except: mrp = 0.0
        try: per_petti = int(sheet.cell(row=row, column=10).value or 0)
        except: per_petti = 0
        
        expiry = str(sheet.cell(row=row, column=11).value or '')
        
        try: opening = int(sheet.cell(row=row, column=14).value or 0)
        except: opening = 0
        try: actual = int(sheet.cell(row=row, column=15).value or opening)
        except: actual = opening
        try: damage = int(sheet.cell(row=row, column=17).value or 0)
        except: damage = 0

        cursor.execute('''
            INSERT INTO products (hsn_code, country, brand, description, price, tax, price_after_tax, mrp, per_petti, expiry_date, opening_stock, actual_stock, damage, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (hsn, country, brand, str(desc).strip(), price, tax, price_after_tax, mrp, per_petti, expiry, opening, actual, damage, 'Completed'))
        prod_id = cursor.lastrowid
        
        # Parse Date-wise history columns from Sheet1 (Col 18 onwards)
        for col in range(18, min(sheet.max_column + 1, 80)):
            val = sheet.cell(row=row, column=col).value
            col_hdr = sheet.cell(row=1, column=col).value
            if val is not None and str(val).strip() != '':
                try:
                    qty = int(val)
                    action = "RECEIVE" if qty > 0 else "ISSUE"
                    date_str = str(col_hdr)[:10] if col_hdr else "Legacy Entry"
                    cursor.execute('''
                        INSERT INTO activity_log (product_id, product_desc, username, action, quantity, prev_stock, new_stock, details, timestamp)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ''', (prod_id, str(desc).strip(), 'system_excel', action, abs(qty), opening, actual, f"Excel entry on {date_str}", datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
                except:
                    pass

    conn.commit()
    conn.close()

init_db()
migrate_from_excel()

# ----------------- ROUTES & APIS -----------------

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json
    username = data.get('username')
    password = data.get('password')
    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password)).fetchone()
    conn.close()
    if user:
        session['username'] = user['username']
        session['role'] = user['role']
        return jsonify({"status": "success", "username": user['username'], "role": user['role']})
    return jsonify({"status": "error", "message": "Invalid username or password"}), 401

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({"status": "success"})

@app.route('/api/products', methods=['GET'])
def get_products():
    conn = get_db()
    products = conn.execute("SELECT * FROM products ORDER BY id ASC").fetchall()
    conn.close()
    return jsonify([dict(p) for p in products])

@app.route('/api/products/adjust', methods=['POST'])
def adjust_stock():
    if 'username' not in session:
        return jsonify({"status": "error", "message": "Unauthorized"}), 401
    
    data = request.json
    prod_id = data.get('id')
    change = int(data.get('change', 0))
    action_type = "RECEIVE" if change > 0 else "ISSUE"
    
    conn = get_db()
    cursor = conn.cursor()
    prod = cursor.execute("SELECT * FROM products WHERE id = ?", (prod_id,)).fetchone()
    if not prod:
        conn.close()
        return jsonify({"status": "error", "message": "Product not found"}), 404
    
    prev_stock = prod['actual_stock']
    new_stock = prev_stock + change
    if new_stock < 0:
        new_stock = 0
    
    # Task progress status update
    status = "Pending" if new_stock <= 5 else "Completed"
    
    cursor.execute("UPDATE products SET actual_stock = ?, status = ? WHERE id = ?", (new_stock, status, prod_id))
    cursor.execute('''
        INSERT INTO activity_log (product_id, product_desc, username, action, quantity, prev_stock, new_stock, details)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    ''', (prod_id, prod['description'], session['username'], action_type, abs(change), prev_stock, new_stock, f"{session['username']} {action_type}D {abs(change)} units"))
    
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "new_stock": new_stock, "product_status": status})

@app.route('/api/products/<int:prod_id>/ledger', methods=['GET'])
def get_product_ledger(prod_id):
    conn = get_db()
    product = conn.execute("SELECT * FROM products WHERE id = ?", (prod_id,)).fetchone()
    if not product:
        conn.close()
        return jsonify({"status": "error", "message": "Not found"}), 404
        
    logs = conn.execute("SELECT * FROM activity_log WHERE product_id = ? ORDER BY id DESC", (prod_id,)).fetchall()
    conn.close()
    return jsonify({
        "product": dict(product),
        "ledger": [dict(log) for log in logs]
    })

@app.route('/api/activity_logs', methods=['GET'])
def get_activity_logs():
    conn = get_db()
    logs = conn.execute("SELECT * FROM activity_log ORDER BY id DESC LIMIT 50").fetchall()
    conn.close()
    return jsonify([dict(l) for l in logs])

@app.route('/api/export/csv')
def export_csv():
    if session.get('role') != 'master':
        return jsonify({"status": "error", "message": "Forbidden"}), 403
    conn = get_db()
    products = conn.execute("SELECT * FROM products").fetchall()
    conn.close()
    
    def generate():
        yield "ID,HSN Code,Country,Brand,Description,Price,Tax,Price After Tax,MRP,Per Petti,Actual Stock,Status\n"
        for p in products:
            yield f'"{p["id"]}","{p["hsn_code"]}","{p["country"]}","{p["brand"]}","{p["description"]}","{p["price"]}","{p["tax"]}","{p["price_after_tax"]}","{p["mrp"]}","{p["per_petti"]}","{p["actual_stock"]}","{p["status"]}"\n'
            
    return Response(generate(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=master_inventory_report.csv"})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
