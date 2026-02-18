# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Module   : process_data.py — Orders Processing
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
# ============================================================
import pandas as pd
import sqlite3
import os
from datetime import datetime

# Configuration
DB_NAME = "finance_system.db"
SAMPLES_DIR = "samples"

def get_db_connection():
    return sqlite3.connect(DB_NAME)

def parse_date(date_str):
    if pd.isna(date_str): return None
    clean_str = str(date_str).strip().replace('"', '').replace('=', '')
    
    formats = [
        '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
        '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y %H:%M UTC', '%d.%m.%Y'
    ]
    for fmt in formats:
        try: return datetime.strptime(clean_str, fmt).date()
        except ValueError: continue
    
    try: return pd.to_datetime(clean_str).date()
    except: return None

def get_week_number(date_obj):
    if date_obj: return date_obj.isocalendar()[1]
    return None

def normalize_columns(df):
    """Clean column names removing quotes and spaces"""
    df.columns = [str(c).replace('"', '').replace('=', '').strip() for c in df.columns]
    return df

def read_file_safe(file_path):
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ['.xls', '.xlsx']:
        try: return pd.read_excel(file_path)
        except: return None
    elif ext == '.csv':
        # Update encoding priority
        encodings = ['utf-8-sig', 'utf-8', 'cp1256', 'latin1']
        separators = [',', ';', '\t', None]
        for enc in encodings:
            for sep in separators:
                try:
                    df = pd.read_csv(file_path, encoding=enc, sep=sep, engine='python')
                    if len(df.columns) > 1: return normalize_columns(df)
                except: continue
    return None

def identify_platform(df, filename):
    columns = [str(c).lower().strip() for c in df.columns]
    col_str = " ".join(columns)
    fname_lower = filename.lower()
    
    # 1. Noon
    if 'order_nr' in columns and 'order_status' in columns: return 'Noon'
    if 'order_nr' in columns and 'noon' in col_str: return 'Noon'
    if 'id_partner' in columns and 'statement_nr' in columns: return 'Noon' # Statement

    # 2. Ilasouq
    if 'ilasouq' in fname_lower:
        if 'تاريخ الطلب' in columns: return 'Ilasouq Orders'
        return 'Ilasouq Collection'
    if 'رقم الطلب' in columns and 'طريقة الدفع' in columns: return 'Ilasouq Orders' # Generic

    # 3. Trendyol
    if 'transaction no' in columns and 'storefront' in columns: return 'Trendyol' # Statement
    if 'الباركود' in columns and 'اسم المنتج' in columns: return 'Trendyol Sales' # Sales Report

    # 4. Amazon
    # Amazon often has quoted headers like "رقم الطلب"
    if any('amazon' in c for c in columns): return 'Amazon'
    if any('نوع المعاملة' in c for c in columns) and any('رقم الطلب' in c for c in columns): return 'Amazon'

    # 5. Tabby & SMSA (Collections)
    if 'تابي' in fname_lower or 'tabby' in fname_lower: return 'Tabby'
    if 'سمسا' in fname_lower or 'smsa' in fname_lower: return 'SMSA'

    return 'Unknown'

def process_file_content(df, platform, file_path):
    print(f"  -> Identified as: {platform}")
    conn = get_db_connection()
    cursor = conn.cursor()
    count = 0
    orders_data = [] # List of tuples to insert

    try:
        # === ILASOUQ ===
        if platform == 'Ilasouq Orders':
             for _, row in df.iterrows():
                oid = str(row['رقم الطلب'])
                date_val = row['تاريخ الطلب']
                if isinstance(date_val, (pd.Timestamp, datetime)):
                    date = date_val.date()
                else:
                    date = parse_date(date_val)
                
                price = float(row['إجمالي الطلب']) if pd.notna(row['إجمالي الطلب']) else 0.0
                shipping = float(row['تكلفة الشحن']) if pd.notna(row['تكلفة الشحن']) else 0.0
                tax = float(row['الضريبة']) if pd.notna(row['الضريبة']) else 0.0
                orders_data.append((oid, 'Ilasouq', date, price, 0.0, shipping, 0.0, tax))

        # === NOON ===
        elif platform == 'Noon':
            # Check if Orders or Statement
            if 'statement_nr' in df.columns: # Statement
                pass 
            elif 'order_nr' in df.columns: # Orders List
                for _, row in df.iterrows():
                    if pd.isna(row['order_nr']): continue
                    raw_oid = str(row['order_nr']).strip()
                    if raw_oid == '': continue
                    
                    # Clean: take only the first token (in case of concatenated text)
                    # Noon order IDs are purely numeric (e.g. 'N123456789')
                    # Strip any non-alphanumeric prefix/suffix noise
                    oid = raw_oid.split()[0]  # Take first word only
                    # Further: keep only alphanumeric characters
                    oid = ''.join(c for c in oid if c.isalnum() or c == '-')
                    if not oid: continue
                    
                    date_val = row.get('order_received_at') or row.get('ordered_date')
                    date = parse_date(date_val)
                    
                    orders_data.append((oid, 'Noon', date, 0.0, 0.0, 0.0, 0.0, 0.0))

        # === TRENDYOL ===
        elif platform == 'Trendyol': # Statement File
            for _, row in df.iterrows():
                # Filter for Sales
                if row.get('Transaction Type') == 'Sale':
                    if pd.notna(row.get('Order Number')):
                        oid = str(int(row['Order Number']))
                        date_str = str(row['Order Date']).replace(' UTC', '')
                        date = parse_date(date_str)
                        price = float(row['Credit']) # Using Credit as price proxy
                        orders_data.append((oid, 'Trendyol', date, price, 0.0, 0.0, 0.0, 0.0))

        # === AMAZON ===
        elif platform == 'Amazon':
            # 1. Try Pandas extraction first
            oid_col = next((c for c in df.columns if 'رقم الطلب' in c or 'Order ID' in c), None)
            date_col = next((c for c in df.columns if 'التاريخ' in c or 'Date' in c), None)
            
            extracted_via_pandas = False
            if oid_col:
                rows_checked = 0
                for _, row in df.iterrows():
                    raw_val = str(row[oid_col])
                    rows_checked += 1

                    raw_oid = raw_val.replace('"', '').replace("'", "").replace('=', '').strip()
                    if len(raw_oid) < 5 or '-' not in raw_oid: continue
                    if not raw_oid[0].isdigit(): continue

                    date_val = row[date_col] if date_col else None
                    try: date = parse_date(date_val)
                    except: date = None
                    
                    orders_data.append((raw_oid, 'Amazon', date, 0.0, 0.0, 0.0, 0.0, 0.0))
                    extracted_via_pandas = True

            # 2. Fallback: Manual CSV Parsing if Pandas failed
            if not extracted_via_pandas or len(orders_data) == 0:
                import csv
                try:
                    with open(file_path, 'r', encoding='utf-8-sig') as f:
                        lines = f.readlines()
                        
                    header_idx = -1
                    headers = []
                    
                    for i, line in enumerate(lines):
                        if 'رقم الطلب' in line or 'Order ID' in line:
                            header_idx = i
                            headers = list(csv.reader([line], delimiter=','))[0]
                            break
                    
                    if header_idx != -1:
                        # Map indices
                        norm_headers = [h.replace('"', '').strip() for h in headers]
                        
                        # New Logic: forceful delimiter
                        if len(headers) == 1 and ',' in headers[0]:
                             headers = headers[0].split(',')
                             norm_headers = [h.replace('"', '').strip() for h in headers]

                        try:
                            oid_idx = next(i for i, h in enumerate(norm_headers) if 'رقم الطلب' in h or 'Order ID' in h)
                            date_idx = next((i for i, h in enumerate(norm_headers) if 'التاريخ' in h or 'Date' in h), -1)
                            
                            # Reading data rows with explicit delimiter
                            reader = csv.reader(lines[header_idx+1:], delimiter=',')
                            for row in reader:
                                if not row: continue
                                
                                # Safety net for split failure
                                if len(row) == 1 and ',' in row[0]:
                                     row = row[0].split(',')

                                if len(row) <= oid_idx: continue
                                
                                raw_val = row[oid_idx]
                                raw_oid = raw_val.replace('"', '').replace("'", "").replace('=', '').strip()
                                
                                if len(raw_oid) < 5 or '-' not in raw_oid: continue
                                if not raw_oid[0].isdigit(): continue
                                
                                date_val = row[date_idx] if date_idx != -1 and len(row) > date_idx else None
                                date = parse_date(date_val)
                                
                                # Avoid duplicates if pandas got some (unlikely if we are here)
                                if not any(d[0] == raw_oid for d in orders_data):
                                    orders_data.append((raw_oid, 'Amazon', date, 0.0, 0.0, 0.0, 0.0, 0.0))
                            
                            print(f"     Manual parsing found {len(orders_data)} orders.")
                        except StopIteration: pass
                except Exception: pass


    except Exception as e:
        print(f"    Error processing rows: {e}")

    # Batch Insert
    new_inserts = 0
    for data in orders_data:
        try:
            order_id = data[0]
            order_date = data[2].isoformat() if data[2] else None
            wk = get_week_number(data[2])
            yr = data[2].year if data[2] else None
            
            cursor.execute('''
                INSERT OR IGNORE INTO orders 
                (order_id, platform, order_date, price, cost, shipping, commission, tax, week_number, year)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (order_id, data[1], order_date, data[3], data[4], data[5], data[6], data[7], wk, yr))
            if cursor.rowcount > 0:
                new_inserts += 1
            count += 1
        except: continue
    
    conn.commit()
    conn.close()
    
    if count > 0:
        print(f"  -> Processed {count} orders (New inserts: {new_inserts}).")
    else:
        print(f"  -> No valid orders extracted.")

def main():
    print(f"Scanning directory: {SAMPLES_DIR}")
    if not os.path.exists(SAMPLES_DIR):
        print("Directory not found!")
        return

    for f in os.listdir(SAMPLES_DIR):
        if f.startswith('~$') or not os.path.isfile(os.path.join(SAMPLES_DIR, f)): continue
        
        path = os.path.join(SAMPLES_DIR, f)
        print(f"\nEvaluating: {f}")
        
        df = read_file_safe(path)
        if df is None:
            print("  -> Failed to read file.")
            continue
            
        platform = identify_platform(df, f)
        if platform == 'Unknown':
            print("  -> Unknown Platform.")
        elif platform in ['Tabby', 'SMSA', 'Trendyol Sales', 'Ilasouq Collection']:
            print(f"  -> Identified as {platform} (Collections/Report). Skipping Orders Insert.")
        else:
            process_file_content(df, platform, path)

if __name__ == "__main__":
    main()
