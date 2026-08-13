import pandas as pd
import sqlite3
import os

DB_FILE = 'inventory.db'

def migrate():
    excel_file = 'NEW APRIL STOCK SHEET 2026.xlsx'
    if not os.path.exists(excel_file):
        print(f"File {excel_file} not found!")
        return

    df = pd.read_excel(excel_file, sheet_name='Sheet2', skiprows=1)
    df = df.iloc[:, :8]
    df.columns = ['brand', 'description', 'rtt', 'rin', 'rit', 'ge', 'total', 'actual_stock']
    
    df['brand'] = df['brand'].ffill()
    df = df.dropna(subset=['description'])
    
    num_cols = ['rtt', 'rin', 'rit', 'ge', 'total', 'actual_stock']
    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0).astype(int)

    conn = sqlite3.connect(DB_FILE)
    df.to_sql('products', conn, if_exists='replace', index_label='id')
    conn.close()
    print("Excel Data Migrated Successfully!")

if __name__ == '__main__':
    migrate()