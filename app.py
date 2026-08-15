import os
import sqlite3
from datetime import datetime, timezone, timedelta
from flask import Flask, render_template, request, jsonify, session, Response
import openpyxl

app = Flask(__name__)
app.secret_key = 'enterprise_super_secret_key_2026'

DB_NAME = 'stock.db'
EXCEL_FILE = 'NEW APRIL STOCK SHEET 2026.xlsx'
MASTER_RESET_PIN = '2026'

IST = timezone(timedelta(hours=5, minutes=30))

def get_current_ist_time():
    return datetime.now(IST).strftime('%Y-%m-%d %I:%M:%S %p')

def get_current_ist_date():
    return datetime.now(IST).strftime('%Y-%m-%d')

def get_db():
    conn = sqlite3.connect(DB_NAME)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_db()
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
    
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS marketplace_products (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            marketplace TEXT NOT NULL,
            marketplace_label TEXT NOT NULL,
            sku TEXT,
            asin TEXT,
            selling_price REAL DEFAULT 0,
            fba_fee REAL DEFAULT 0,
            purchase_price REAL DEFAULT 0,
            conv_rate REAL DEFAULT 0,
            purchase_rate REAL DEFAULT 0,
            freight REAL DEFAULT 0,
            min_selling_rate REAL DEFAULT 0,
            margin REAL DEFAULT 0,
            product_link TEXT
        )
    ''')
    
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
            timestamp TEXT
        )
    ''')
    
    cursor.execute("INSERT OR IGNORE INTO users (username, password, role) VALUES ('master', 'admin123', 'master')")
    cursor.execute("INSERT OR IGNORE INTO users (username, password, role) VALUES ('worker1', 'worker123', 'worker')")
    cursor.execute("INSERT OR IGNORE INTO users (username, password, role) VALUES ('worker2', 'worker123', 'worker')")
    
    conn.commit()
    conn.close()

def migrate_from_excel():
    conn = get_db()
    cursor = conn.cursor()
    
    if not os.path.exists(EXCEL_FILE):
        conn.close()
        return

    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    
    # 1. Migrate Warehouse Products
    cursor.execute("SELECT COUNT(*) FROM products")
    if cursor.fetchone()[0] == 0 and 'Sheet1' in wb.sheetnames:
        sheet = wb['Sheet1']
        for row in range(3, sheet.max_row + 1):
            desc = sheet.cell(row=row, column=5).value
            if not desc or str(desc).strip().lower() in ['discription', 'description', 'none', '']:
                continue
                
            hsn = str(sheet.cell(row=row, column=2).value or '').replace('.0', '')
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
            try: opening = int(float(sheet.cell(row=row, column=14).value or 0))
            except: opening = 0
            try: actual = int(float(sheet.cell(row=row, column=15).value or opening))
            except: actual = opening
            try: damage = int(float(sheet.cell(row=row, column=17).value or 0))
            except: damage = 0

            status = "Pending" if actual <= 5 else "Completed"

            cursor.execute('''
                INSERT INTO products (hsn_code, country, brand, description, price, tax, price_after_tax, mrp, per_petti, expiry_date, opening_stock, actual_stock, damage, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (hsn, country, brand, str(desc).strip(), price, tax, price_after_tax, mrp, per_petti, expiry, opening, actual, damage, status))
            prod_id = cursor.lastrowid
            
            for col in range(18, min(sheet.max_column + 1, 50)):
                val = sheet.cell(row=row, column=col).value
                col_hdr = sheet.cell(row=2, column=col).value
                if val is not None and str(val).strip() not in ['', 'None', 'nan']:
                    try:
                        qty = int(float(val))
                        action = "RECEIVE" if qty > 0 else "ISSUE"
                        date_str = str(col_hdr)[:10] if col_hdr else "Excel Import"
                        cursor.execute('''
                            INSERT INTO activity_log (product_id, product_desc, username, action, quantity, prev_stock, new_stock, details, timestamp)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                        ''', (prod_id, str(desc).strip(), 'System Excel', action, abs(qty), opening, actual, f"Initial Data from {date_str}", get_current_ist_time()))
                    except:
                        pass

    # 2. Migrate All Global Country Marketplace Sheets
    cursor.execute("SELECT COUNT(*) FROM marketplace_products")
    if cursor.fetchone()[0] == 0:
        marketplaces = [
            ('GTUS', 'USA 🇺🇸 (GTUS)'),
            ('CANADA', 'Canada 🇨🇦'),
            ('GTAU', 'Australia 🇦🇺 (GTAU)'),
            ('UAE', 'UAE 🇦🇪'),
            ('INDIA', 'India 🇮🇳'),
            ('VAIS NEW', 'USA 🇺🇸 (VAIS)'),
            ('RTAU', 'Australia 🇦🇺 (RTAU)')
        ]
        
        for m_sheet, m_label in marketplaces:
            if m_sheet not in wb.sheetnames:
                continue
            ws = wb[m_sheet]
            header_row = None
            for r in range(1, 10):
                row_vals = [str(ws.cell(row=r, column=c).value or '').strip().upper() for c in range(1, 10)]
                if any('SKU' in v for v in row_vals):
                    header_row = r
                    break
            
            if not header_row:
                continue
                
            for r in range(header_row + 1, ws.max_row + 1):
                sku = ws.cell(row=r, column=2).value
                if not sku or str(sku).strip().lower() in ['none', '', 'sku', 'seller-sku']:
                    continue
                
                asin = str(ws.cell(row=r, column=3).value or '')
                try: price = float(ws.cell(row=r, column=4).value or 0)
                except: price = 0.0
                try: fba = float(ws.cell(row=r, column=5).value or 0)
                except: fba = 0.0
                try: purchase_price = float(ws.cell(row=r, column=6).value or 0)
                except: purchase_price = 0.0
                try: conv_rate = float(ws.cell(row=r, column=7).value or 0)
                except: conv_rate = 0.0
                try: purchase_rate = float(ws.cell(row=r, column=8).value or 0)
                except: purchase_rate = 0.0
                try: freight = float(ws.cell(row=r, column=9).value or 0)
                except: freight = 0.0
                try: min_selling = float(ws.cell(row=r, column=10).value or 0)
                except: min_selling = 0.0
                try: margin = float(ws.cell(row=r, column=11).value or 0)
                except: margin = 0.0
                
                link = ""
                for l_col in [12, 13, 14]:
                    l_val = str(ws.cell(row=r, column=l_col).value or '').strip()
                    if l_val.startswith('http'):
                        link = l_val
                        break
                
                cursor.execute('''
                    INSERT INTO marketplace_products 
                    (marketplace, marketplace_label, sku, asin, selling_price, fba_fee, purchase_price, conv_rate, purchase_rate, freight, min_selling_rate, margin, product_link)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (m_sheet, m_label, str(sku).strip(), asin, price, fba, purchase_price, conv_rate, purchase_rate, freight, min_selling, margin, link))

    conn.commit()
    conn.close()

if os.path.exists(DB_NAME):
    try: os.remove(DB_NAME)
    except: pass

init_db()
migrate_from_excel()

# ----------------- APIS -----------------

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/login', methods=['POST'])
def login():
    data = request.json or {}
    u = data.get('username', '').strip()
    p = data.get('password', '').strip()
    conn = get_db()
    user = conn.execute("SELECT * FROM users WHERE username = ? AND password = ?", (u, p)).fetchone()
    conn.close()
    if user:
        session['username'] = user['username']
        session['role'] = user['role']
        return jsonify({"status": "success", "username": user['username'], "role": user['role']})
    return jsonify({"status": "error", "message": "Invalid username or password"}), 401

@app.route('/api/reset-password', methods=['POST'])
def reset_password():
    data = request.json or {}
    username = data.get('username', '').strip()
    new_password = data.get('new_password', '').strip()
    pin = data.get('pin', '').strip()
    
    if pin != MASTER_RESET_PIN:
        return jsonify({"status": "error", "message": "Invalid Master PIN"}), 403
    
    if not username or not new_password:
        return jsonify({"status": "error", "message": "All fields are required"}), 400

    conn = get_db()
    cursor = conn.cursor()
    user = cursor.execute("SELECT * FROM users WHERE username = ?", (username,)).fetchone()
    if not user:
        conn.close()
        return jsonify({"status": "error", "message": "User not found"}), 404
        
    cursor.execute("UPDATE users SET password = ? WHERE username = ?", (new_password, username))
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "message": "Password reset successfully!"})

@app.route('/api/logout', methods=['POST'])
def logout():
    session.clear()
    return jsonify({"status": "success"})

@app.route('/api/worker-stats', methods=['GET'])
def worker_stats():
    username = session.get('username', '')
    today_ist = get_current_ist_date()
    conn = get_db()
    today_received = conn.execute("SELECT SUM(quantity) FROM activity_log WHERE username = ? AND action = 'RECEIVE' AND timestamp LIKE ?", (username, f"{today_ist}%")).fetchone()[0] or 0
    today_issued = conn.execute("SELECT SUM(quantity) FROM activity_log WHERE username = ? AND action = 'ISSUE' AND timestamp LIKE ?", (username, f"{today_ist}%")).fetchone()[0] or 0
    today_ops = conn.execute("SELECT COUNT(*) FROM activity_log WHERE username = ? AND timestamp LIKE ?", (username, f"{today_ist}%")).fetchone()[0] or 0
    conn.close()
    return jsonify({
        "today_received": today_received,
        "today_issued": today_issued,
        "today_ops": today_ops
    })

@app.route('/api/stats', methods=['GET'])
def get_stats():
    conn = get_db()
    total_items = conn.execute("SELECT COUNT(*) FROM products").fetchone()[0]
    total_qty = conn.execute("SELECT SUM(actual_stock) FROM products").fetchone()[0] or 0
    low_stock = conn.execute("SELECT COUNT(*) FROM products WHERE actual_stock <= 10").fetchone()[0]
    pending_tasks = conn.execute("SELECT COUNT(*) FROM products WHERE status = 'Pending'").fetchone()[0]
    total_global_skus = conn.execute("SELECT COUNT(*) FROM marketplace_products").fetchone()[0]
    top_stocks = conn.execute("SELECT description, actual_stock FROM products ORDER BY actual_stock DESC LIMIT 6").fetchall()
    conn.close()
    return jsonify({
        "total_items": total_items,
        "total_qty": total_qty,
        "low_stock": low_stock,
        "pending_tasks": pending_tasks,
        "total_global_skus": total_global_skus,
        "top_stocks": [dict(r) for r in top_stocks]
    })

@app.route('/api/products', methods=['GET'])
def get_products():
    conn = get_db()
    products = conn.execute("SELECT * FROM products ORDER BY id ASC").fetchall()
    conn.close()
    return jsonify([dict(p) for p in products])

# Add Product - Available to BOTH Master and Worker
@app.route('/api/products/add', methods=['POST'])
def add_product():
    if 'username' not in session:
        return jsonify({"status": "error", "message": "Unauthorized"}), 401
    
    data = request.json or {}
    desc = data.get('description', '').strip()
    if not desc:
        return jsonify({"status": "error", "message": "Product description is required"}), 400
    
    hsn = data.get('hsn_code', '').strip()
    brand = data.get('brand', 'GENERAL').strip()
    country = data.get('country', 'INDIA').strip()
    price = float(data.get('price', 0) or 0)
    tax = float(data.get('tax', 0.18) or 0.18)
    price_after_tax = round(price * (1 + tax), 2)
    mrp = float(data.get('mrp', 0) or 0)
    per_petti = int(data.get('per_petti', 0) or 0)
    expiry = data.get('expiry_date', '').strip()
    opening = int(data.get('opening_stock', 0) or 0)
    
    status = "Pending" if opening <= 5 else "Completed"
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO products (hsn_code, country, brand, description, price, tax, price_after_tax, mrp, per_petti, expiry_date, opening_stock, actual_stock, damage, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
    ''', (hsn, country, brand, desc, price, tax, price_after_tax, mrp, per_petti, expiry, opening, opening, status))
    prod_id = cursor.lastrowid
    
    cursor.execute('''
        INSERT INTO activity_log (product_id, product_desc, username, action, quantity, prev_stock, new_stock, details, timestamp)
        VALUES (?, ?, ?, 'NEW_ITEM', ?, 0, ?, 'Product Added to Catalog', ?)
    ''', (prod_id, desc, session['username'], opening, opening, get_current_ist_time()))
    
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "message": "Product added successfully!"})

# Delete Product - Exclusive to MASTER
@app.route('/api/products/delete/<int:prod_id>', methods=['DELETE'])
def delete_product(prod_id):
    if session.get('role') != 'master':
        return jsonify({"status": "error", "message": "Only Master can delete products from catalog"}), 403
    
    conn = get_db()
    cursor = conn.cursor()
    prod = cursor.execute("SELECT * FROM products WHERE id = ?", (prod_id,)).fetchone()
    if not prod:
        conn.close()
        return jsonify({"status": "error", "message": "Product not found"}), 404
        
    cursor.execute("DELETE FROM products WHERE id = ?", (prod_id,))
    cursor.execute('''
        INSERT INTO activity_log (product_id, product_desc, username, action, quantity, prev_stock, new_stock, details, timestamp)
        VALUES (?, ?, ?, 'DELETE', 0, ?, 0, 'Product Removed by Master', ?)
    ''', (prod_id, prod['description'], session['username'], prod['actual_stock'], get_current_ist_time()))
    
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "message": f"Product #{prod_id} successfully deleted"})

# Edit Product - Exclusive to MASTER
@app.route('/api/products/edit', methods=['POST'])
def edit_product():
    if session.get('role') != 'master':
        return jsonify({"status": "error", "message": "Only Master can edit product details"}), 403
        
    data = request.json or {}
    prod_id = data.get('id')
    desc = data.get('description', '').strip()
    if not desc:
        return jsonify({"status": "error", "message": "Description cannot be empty"}), 400
        
    hsn = data.get('hsn_code', '').strip()
    brand = data.get('brand', 'GENERAL').strip()
    country = data.get('country', 'INDIA').strip()
    price = float(data.get('price', 0) or 0)
    mrp = float(data.get('mrp', 0) or 0)
    tax = float(data.get('tax', 0.18) or 0.18)
    price_after_tax = round(price * (1 + tax), 2)
    per_petti = int(data.get('per_petti', 0) or 0)
    expiry = data.get('expiry_date', '').strip()
    
    conn = get_db()
    cursor = conn.cursor()
    cursor.execute('''
        UPDATE products SET 
            description = ?, hsn_code = ?, brand = ?, country = ?, 
            price = ?, tax = ?, price_after_tax = ?, mrp = ?, 
            per_petti = ?, expiry_date = ?
        WHERE id = ?
    ''', (desc, hsn, brand, country, price, tax, price_after_tax, mrp, per_petti, expiry, prod_id))
    
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "message": "Product updated successfully!"})

@app.route('/api/products/adjust', methods=['POST'])
def adjust_stock():
    if 'username' not in session:
        return jsonify({"status": "error", "message": "Unauthorized"}), 401
    
    data = request.json or {}
    prod_id = data.get('id')
    mode = data.get('mode', 'STOCK')
    change = int(data.get('change', 0))
    remarks = data.get('remarks', '').strip()
    
    conn = get_db()
    cursor = conn.cursor()
    prod = cursor.execute("SELECT * FROM products WHERE id = ?", (prod_id,)).fetchone()
    if not prod:
        conn.close()
        return jsonify({"status": "error", "message": "Product not found"}), 404
    
    prev_stock = prod['actual_stock']
    prev_damage = prod['damage']
    
    if mode == 'DAMAGE':
        if change > prev_stock:
            conn.close()
            return jsonify({"status": "error", "message": f"Cannot mark {change} as damaged. Only {prev_stock} available in stock!"}), 400
        new_stock = prev_stock - change
        new_damage = prev_damage + change
        status = "Pending" if new_stock <= 5 else "Completed"
        cursor.execute("UPDATE products SET actual_stock = ?, damage = ?, status = ? WHERE id = ?", (new_stock, new_damage, status, prod_id))
        detail_str = remarks if remarks else f"Moved {change} units to Damaged/Defective"
        cursor.execute('''
            INSERT INTO activity_log (product_id, product_desc, username, action, quantity, prev_stock, new_stock, details, timestamp)
            VALUES (?, ?, ?, 'DAMAGE', ?, ?, ?, ?, ?)
        ''', (prod_id, prod['description'], session['username'], change, prev_stock, new_stock, detail_str, get_current_ist_time()))
    else:
        action_type = "RECEIVE" if change > 0 else "ISSUE"
        if change < 0 and abs(change) > prev_stock:
            conn.close()
            return jsonify({"status": "error", "message": f"Insufficient Stock! Available: {prev_stock}, Requested: {abs(change)}"}), 400
            
        new_stock = max(0, prev_stock + change)
        status = "Pending" if new_stock <= 5 else "Completed"
        cursor.execute("UPDATE products SET actual_stock = ?, status = ? WHERE id = ?", (new_stock, status, prod_id))
        detail_str = remarks if remarks else f"{session['username']} {action_type}D {abs(change)} units"
        cursor.execute('''
            INSERT INTO activity_log (product_id, product_desc, username, action, quantity, prev_stock, new_stock, details, timestamp)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (prod_id, prod['description'], session['username'], action_type, abs(change), prev_stock, new_stock, detail_str, get_current_ist_time()))
    
    conn.commit()
    conn.close()
    return jsonify({"status": "success", "new_stock": new_stock, "product_status": status})

@app.route('/api/marketplaces/products', methods=['GET'])
def get_marketplace_products():
    mp = request.args.get('marketplace', 'GTUS')
    conn = get_db()
    items = conn.execute("SELECT * FROM marketplace_products WHERE marketplace = ?", (mp,)).fetchall()
    conn.close()
    return jsonify([dict(i) for i in items])

@app.route('/api/ledger', methods=['GET'])
def get_ledger():
    prod_id = request.args.get('product_id')
    conn = get_db()
    if prod_id:
        logs = conn.execute("SELECT * FROM activity_log WHERE product_id = ? ORDER BY id DESC", (prod_id,)).fetchall()
    else:
        logs = conn.execute("SELECT * FROM activity_log ORDER BY id DESC LIMIT 150").fetchall()
    conn.close()
    return jsonify([dict(l) for l in logs])

# ----------------- CSV EXPORT SUITE -----------------

@app.route('/api/export/warehouse-csv')
def export_warehouse_csv():
    conn = get_db()
    products = conn.execute("SELECT * FROM products ORDER BY id ASC").fetchall()
    conn.close()
    
    def generate():
        yield "ID,HSN Code,Country,Brand,Description,Price,Tax,Price After Tax,MRP,Units Per Petti,Expiry Date,Opening Stock,Actual Stock,Damage Units,Status\n"
        for p in products:
            desc = p["description"].replace('"', '""')
            yield f'"{p["id"]}","{p["hsn_code"]}","{p["country"]}","{p["brand"]}","{desc}","{p["price"]}","{p["tax"]}","{p["price_after_tax"]}","{p["mrp"]}","{p["per_petti"]}","{p["expiry_date"]}","{p["opening_stock"]}","{p["actual_stock"]}","{p["damage"]}","{p["status"]}"\n'
            
    return Response(generate(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=Warehouse_Stock_Report.csv"})

@app.route('/api/export/marketplace-csv')
def export_marketplace_csv():
    mp = request.args.get('marketplace', 'GTUS')
    conn = get_db()
    items = conn.execute("SELECT * FROM marketplace_products WHERE marketplace = ?", (mp,)).fetchall()
    conn.close()
    
    def generate():
        yield "ID,Marketplace,SKU,ASIN,Selling Price,FBA Fee,Purchase Price,Conversion Rate,Purchase Rate,Freight,Min Selling Price,Profit Margin %,Product Link\n"
        for item in items:
            sku = item["sku"].replace('"', '""')
            yield f'"{item["id"]}","{item["marketplace_label"]}","{sku}","{item["asin"]}","{item["selling_price"]}","{item["fba_fee"]}","{item["purchase_price"]}","{item["conv_rate"]}","{item["purchase_rate"]}","{item["freight"]}","{item["min_selling_rate"]}","{item["margin"]}","{item["product_link"]}"\n'
            
    return Response(generate(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=Marketplace_{mp}_Pricing_Report.csv"})

@app.route('/api/export/ledger-csv')
def export_ledger_csv():
    prod_id = request.args.get('product_id', '')
    conn = get_db()
    if prod_id:
        logs = conn.execute("SELECT * FROM activity_log WHERE product_id = ? ORDER BY id DESC", (prod_id,)).fetchall()
    else:
        logs = conn.execute("SELECT * FROM activity_log ORDER BY id DESC").fetchall()
    conn.close()
    
    def generate():
        yield "ID,Product ID,Product Name,Operator,Action,Quantity,Previous Stock,New Balance,Audit Details,Timestamp (IST)\n"
        for l in logs:
            pname = l["product_desc"].replace('"', '""')
            det = (l["details"] or "").replace('"', '""')
            yield f'"{l["id"]}","{l["product_id"]}","{pname}","{l["username"]}","{l["action"]}","{l["quantity"]}","{l["prev_stock"]}","{l["new_stock"]}","{det}","{l["timestamp"]}"\n'
            
    return Response(generate(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=Complete_Audit_Ledger.csv"})

if __name__ == '__main__':
    app.run(debug=True, port=5000)
