# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Module   : process_collections.py — Collections Processing
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
# ============================================================
import pandas as pd
import sqlite3
import os
import csv
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
    df.columns = [str(c).replace('"', '').replace('=', '').strip().lower() for c in df.columns]
    return df

def read_file_with_header(file_path, header_row=0):
    """Reads Excel/CSV with specific header row index"""
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xls', '.xlsx']:
            df = pd.read_excel(file_path, header=header_row)
        elif ext == '.csv':
            try:
                df = pd.read_csv(file_path, header=header_row, encoding='utf-8-sig', sep=None, engine='python')
            except:
                df = pd.read_csv(file_path, header=header_row, encoding='latin1')
        
        return normalize_columns(df)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return None

def process_collections():
    conn = get_db_connection()
    cursor = conn.cursor()
    
    files = [f for f in os.listdir(SAMPLES_DIR) if os.path.isfile(os.path.join(SAMPLES_DIR, f)) and not f.startswith('~$')]
    
    total_processed = 0

    for filename in files:
        path = os.path.join(SAMPLES_DIR, filename)
        fname_lower = filename.lower()
        
        # print(f"DEBUG Check: {filename}") # Uncomment if needed
        
        collections_data = [] # List of (order_id, amount, date)
        platform_name = "Unknown"

        try:
            # ... (Existing Tabby/SMSA/Ilasouq blocks remain same - assuming they work) ...
            if 'tabby' in fname_lower or 'تابي' in fname_lower:
                platform_name = "Tabby (Ilasouq)"
                # ... (rest of Tabby logic)
                df = read_file_with_header(path, header_row=10)
                for _, row in df.iterrows():
                    try:
                        oid = str(row['order number']).strip()
                        if oid.lower() == 'nan': continue
                        amount = float(row['transferred amount'])
                        date = parse_date(row['transfer date'])
                        collections_data.append((oid, amount, date))
                    except: continue

            elif 'smsa' in fname_lower or 'سمسا' in fname_lower:
                platform_name = "SMSA (Ilasouq COD)"
                df = read_file_with_header(path, header_row=2)
                for _, row in df.iterrows():
                    try:
                        oid = str(row['ref no']).strip()
                        if oid.lower() == 'nan': continue
                        amount = float(row['cod amount'])
                        date = parse_date(row['payment date'])
                        collections_data.append((oid, amount, date))
                    except: continue

            elif 'تحصيل-ilasouq' in fname_lower:
                platform_name = "Ilasouq Electronic"
                df = read_file_with_header(path, header_row=0)
                col_amount = next((c for c in df.columns if 'بعد الضريبة' in c or 'إجمالي الطلب' in c), None)
                col_oid = next((c for c in df.columns if 'رقم الطلب' in c), None)
                if col_amount and col_oid:
                    for _, row in df.iterrows():
                        try:
                            oid = str(row[col_oid]).strip()
                            amount = float(row[col_amount])
                            date = datetime.now().date() 
                            collections_data.append((oid, amount, date))
                        except: continue

            # -----------------------------------------------------
            # 4. NOON (Statement)
            # -----------------------------------------------------
            elif 'noon' in fname_lower or 'كشف-حساب-نون' in fname_lower:
                platform_name = "Noon Statement"
                print(f"Processing {filename} as {platform_name}...")
                
                df = read_file_with_header(path, header_row=0)
                if df is not None:
                     if 'statement_nr' in df.columns or 'order_nr' in df.columns:
                         col_oid = 'order_nr' if 'order_nr' in df.columns else None
                         if col_oid:
                             for _, row in df.iterrows():
                                try:
                                    oid = str(row[col_oid]).strip()
                                    amount = float(row['total_payment']) 
                                    date = parse_date(row.get('statement_date', datetime.now().date())) # fallback date
                                    if amount != 0:
                                        collections_data.append((oid, amount, date))
                                except: continue

            # -----------------------------------------------------
            # 5. TRENDYOL (Statement)
            # -----------------------------------------------------
            elif ('trendyol' in fname_lower or 'ترنديول' in fname_lower) and ('statement' in fname_lower or 'كشف' in fname_lower):
                if 'sales' in fname_lower or 'مبيعات' in fname_lower: continue 

                platform_name = "Trendyol Statement"
                print(f"Processing {filename} as {platform_name}...")
                
                df = read_file_with_header(path, header_row=0)
                # Columns: 'transaction type', 'credit', 'order number', 'payment date'
                for _, row in df.iterrows():
                    try:
                        if pd.notna(row.get('payment date')):
                            oid = str(int(row['order number']))
                            amount = float(row['credit'])
                            date = parse_date(str(row['payment date']).replace(' UTC',''))
                            if amount != 0:
                                collections_data.append((oid, amount, date))
                    except: continue

            # -----------------------------------------------------
            # 6. AMAZON (Statement/Transactions)
            # -----------------------------------------------------
            elif 'المعاملات' in fname_lower or 'amazon' in fname_lower:
                platform_name = "Amazon Statement"
                # print(f"Processing {filename} as {platform_name}...")
                
                # Manual Parsing Fallback immediately
                import csv
                try:
                    lines = []
                    try:
                        with open(path, 'r', encoding='utf-8-sig') as f: lines = f.readlines()
                    except:
                        with open(path, 'r', encoding='cp1256') as f: lines = f.readlines()

                    header_idx = -1
                    for i, line in enumerate(lines[:20]):
                        if 'رقم الطلب' in line or 'Order ID' in line:
                            header_idx = i
                            break
                    
                    if header_idx != -1:
                        # Try standard csv reader first
                        headers = list(csv.reader([lines[header_idx]]))[0]
                        
                        # Fallback: if failed to split, split manually
                        if len(headers) < 2 and ',' in lines[header_idx]:
                            headers = lines[header_idx].strip().replace('"', '').split(',')

                        norm_headers = [h.replace('"', '').strip() for h in headers]
                        
                        oid_idx = next((i for i, h in enumerate(norm_headers) if 'رقم الطلب' in h or 'Order ID' in h), -1)
                        total_idx = next((i for i, h in enumerate(norm_headers) if 'الإجمالي' in h or 'Total' in h), -1)
                        date_idx = next((i for i, h in enumerate(norm_headers) if 'التاريخ' in h or 'Date' in h), -1)

                        if oid_idx != -1 and total_idx != -1:
                            reader = csv.reader(lines[header_idx+1:])
                            for row in reader:
                                if not row: continue
                                
                                # Fallback for row splitting failure
                                if len(row) < 2 and len(lines[header_idx+1]) > 10:
                                     # This is tricky for rows with quoted commas, but let's try a naive split if reader failed
                                     # Re-reading the raw line from lines[] strictly is hard due to index alignment.
                                     # But usually if header splits, row splits.
                                     # If header fallback was used, row might need it too?
                                     # Actually, let's trust csv reader first, but if row is length 1, try naive split.
                                     if len(row) == 1 and ',' in row[0]:
                                         row = row[0].split(',')

                                if len(row) <= max(oid_idx, total_idx): continue
                                
                                raw_oid = row[oid_idx].replace('"', '').replace('=', '').strip()
                                if len(raw_oid) < 5 or '-' not in raw_oid: continue
                                
                                raw_amt = row[total_idx].replace('"', '').replace('=', '').strip().replace(',', '')
                                try: amount = float(raw_amt)
                                except: continue
                                
                                date_val = row[date_idx] if date_idx != -1 and len(row) > date_idx else None
                                date = parse_date(date_val)
                                
                                if amount != 0:
                                    collections_data.append((raw_oid, amount, date))
                except Exception as e:
                    print(f"  Amazon Parse Error: {e}")

            # -----------------------------------------------------
            # INSERT DATA
            # -----------------------------------------------------
            if collections_data:
                inserted_count = 0
                for item in collections_data:
                    oid, amt, dt = item
                    if dt is None: dt_str = None
                    else: dt_str = dt.isoformat()
                    
                    wk = get_week_number(dt)
                    yr = dt.year if dt else None
                    
                    # Check duplication
                    cursor.execute('''
                        SELECT 1 FROM collections 
                        WHERE order_id = ? AND ABS(collected_amount - ?) < 0.001 AND (collection_date = ? OR collection_date IS NULL)
                    ''', (oid, amt, dt_str))
                    
                    if cursor.fetchone() is None:
                        cursor.execute('''
                            INSERT INTO collections (order_id, collected_amount, collection_date, week_number, year)
                            VALUES (?, ?, ?, ?, ?)
                        ''', (oid, amt, dt_str, wk, yr))
                        inserted_count += 1
                
                conn.commit()
                if inserted_count > 0:
                     print(f"  -> Inserted {inserted_count} collection records.")
                elif len(collections_data) > 0:
                     print(f"  -> {len(collections_data)} records found but already exist.")
                total_processed += inserted_count
            else:
                if platform_name != "Unknown":
                    print(f"  -> No valid records extracted from {platform_name}.")

        except Exception as e:
            print(f"Error processing {filename}: {e}")

    conn.close()
    print(f"\nTotal Collections Inserted: {total_processed}")

if __name__ == "__main__":
    process_collections()
