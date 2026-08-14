import os
import sqlite3
from functools import wraps
from flask import Flask, render_template, request, jsonify, session
from werkzeug.security import generate_password_hash, check_password_hash

app = Flask(__name__)
app.secret_key = 'stock-secret-key-change-this'
DB_FILE = 'inventory.db'

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
    conn.close()

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
        log_action(username, 'LOGIN', 'Logged into app')
        return jsonify({'status': 'success', 'username': username, 'role': user[1]})
    
    return jsonify({'error': 'Invalid credentials'}), 400

@app.route('/api/logout', methods=['POST'])
def logout():
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
    brand = data.get('brand', '').upper()
    description = data.get('description', '')
    actual_stock = int(data.get('actual_stock', 0))
    rtt = int(data.get('rtt', 0))
    rin = int(data.get('rin', 0))
    rit = int(data.get('rit', 0))
    ge = int(data.get('ge', 0))
    total = rtt + rin + rit + ge

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("""
        INSERT INTO products (brand, description, rtt, rin, rit, ge, total, actual_stock)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, (brand, description, rtt, rin, rit, ge, total, actual_stock))
    conn.commit()
    conn.close()
    
    log_action(session['username'], 'ADD_PRODUCT', f"Added {brand} - {description}")
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
    
    log_action(session['username'], action_type.upper(), f"{action_type.capitalize()}d {amount} units of {prod[0]} - {prod[1]}")
    return jsonify({'status': 'success', 'new_stock': new_stock})

@app.route('/api/products/delete/<int:prod_id>', methods=['DELETE'])
@master_required
def delete_product(prod_id):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM products WHERE id = ?", (prod_id,))
    conn.commit()
    conn.close()
    log_action(session['username'], 'DELETE', f"Deleted product ID {prod_id}")
    return jsonify({'status': 'success'})

@app.route('/api/activity', methods=['GET'])
@master_required
def get_activity():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM audit_logs ORDER BY timestamp DESC LIMIT 30")
    logs = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return jsonify({'logs': logs})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
