# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Module   : process_collections.py — Collections Processing
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
#  Updated   : February 2026 — Full platform support with fee tracking & returns
# ============================================================
#
#  منطق التحصيل لكل منصة:
#  ─────────────────────────────────────────────────────────
#  Tabby   : Order Number = صلة الوصل
#            original_amount = Order Amount (المبلغ الأصلي)
#            collection_fee  = Total Deduction (عمولة تابي)
#            collected_amount= Transferred amount (صافي المحصل)
#
#  SMSA    : Ref No = رقم البوليصية = صلة الوصل مع الطلبات
#            COD Amount = المبلغ المحصل
#            collection_fee = COD Charges (عمولة التحصيل)
#
#  Noon    : order_nr = صلة الوصل (كشف حساب CSV)
#
#  Trendyol: يحتوي على Sales و Returns في نفس الملف
#            is_return=1 للمرتجعات (قيمة سالبة)
#
#  Amazon  : رقم الطلب = صلة الوصل
#
#  Ilasouq : رقم الطلب = صلة الوصل
# ============================================================

import pandas as pd
import sqlite3
import os
import csv
from datetime import datetime

# Configuration
DB_NAME     = "finance_system.db"
SAMPLES_DIR = "samples"

def get_db_connection():
    return sqlite3.connect(DB_NAME)

def parse_date(date_str):
    if pd.isna(date_str): return None
    clean_str = str(date_str).strip().replace('"', '').replace('=', '').replace(' UTC', '')
    formats = [
        '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M', '%Y-%m-%d',
        '%d/%m/%Y', '%d-%m-%Y', '%d.%m.%Y %H:%M', '%d.%m.%Y',
        '%m/%d/%Y %I:%M:%S %p', '%m/%d/%Y'
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
    df.columns = [str(c).replace('"', '').replace('=', '').strip() for c in df.columns]
    return df

def read_file_with_header(file_path, header_row=0):
    """Reads Excel/CSV with specific header row index"""
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xls', '.xlsx']:
            df = pd.read_excel(file_path, header=header_row)
        elif ext in ['.csv', '.txt']:
            try:
                df = pd.read_csv(file_path, header=header_row, encoding='utf-8-sig', sep=None, engine='python')
            except:
                df = pd.read_csv(file_path, header=header_row, encoding='latin1')
        return normalize_columns(df)
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return None

def detect_account_name(filename):
    """
    يكتشف اسم الحساب من اسم الملف
    مثال: 'تابي-تحصيل-ilasouq.xlsx' -> 'ILASOQ'
           'كشف-حساب-العمليات-ترنديول-امواج.xlsx' -> 'أمواج'
    """
    fname = filename.lower()
    if 'امواج' in fname or 'amwaj' in fname:
        return 'أمواج'
    if 'ilasouq' in fname or 'ilasoq' in fname:
        return 'ILASOQ'
    return ''

def insert_collection(cursor, snapshot_id, order_id, original_amount,
                      collection_fee, collected_amount, collection_date,
                      is_return=0, account_name=''):
    """
    إدراج سجل تحصيل واحد مع التحقق من التكرار
    Returns True if inserted, False if duplicate
    """
    dt_str = collection_date.isoformat() if collection_date else None
    wk     = get_week_number(collection_date)
    yr     = collection_date.year if collection_date else None

    # التحقق من التكرار
    cursor.execute('''
        SELECT 1 FROM collections
        WHERE snapshot_id = ? AND order_id = ?
          AND ABS(collected_amount - ?) < 0.001
          AND (collection_date = ? OR collection_date IS NULL)
          AND is_return = ?
    ''', (snapshot_id, order_id, collected_amount, dt_str, is_return))

    if cursor.fetchone() is None:
        cursor.execute('''
            INSERT INTO collections
            (snapshot_id, order_id, original_amount, collection_fee,
             collected_amount, collection_date, is_return, account_name,
             week_number, year)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (snapshot_id, order_id, original_amount, collection_fee,
              collected_amount, dt_str, is_return, account_name, wk, yr))
        return True
    return False


def process_collections(snapshot_id=0):
    conn   = get_db_connection()
    cursor = conn.cursor()

    files = [
        f for f in os.listdir(SAMPLES_DIR)
        if os.path.isfile(os.path.join(SAMPLES_DIR, f)) and not f.startswith('~$')
    ]

    total_processed = 0

    for filename in files:
        path        = os.path.join(SAMPLES_DIR, filename)
        fname_lower = filename.lower()
        account     = detect_account_name(filename)

        # قائمة السجلات: (order_id, original_amount, collection_fee, collected_amount, date, is_return)
        collections_data = []
        platform_name    = "Unknown"

        try:
            # ── 1. TABBY ─────────────────────────────────────────────────────
            # صلة الوصل: Order Number
            # المبلغ الأصلي: Order Amount
            # عمولة التحصيل: Total Deduction
            # صافي المحصل: Transferred amount
            # ─────────────────────────────────────────────────────────────────
            if 'tabby' in fname_lower or 'تابي' in fname_lower:
                platform_name = "Tabby"
                df = read_file_with_header(path, header_row=10)
                if df is not None:
                    # تطبيع أسماء الأعمدة للمقارنة
                    col_map = {c.strip().lower(): c for c in df.columns}

                    col_order_num   = next((col_map[k] for k in col_map if 'order number' in k), None)
                    col_order_amt   = next((col_map[k] for k in col_map if 'order amount' in k), None)
                    col_total_ded   = next((col_map[k] for k in col_map if 'total deduction' in k or 'total fee' in k), None)
                    col_transferred = next((col_map[k] for k in col_map if 'transferred amount' in k or 'transfer amount' in k), None)
                    col_date        = next((col_map[k] for k in col_map if 'transfer date' in k), None)
                    col_type        = next((col_map[k] for k in col_map if col_map[k].strip().lower() == 'type'), None)

                    if col_order_num and col_transferred:
                        for _, row in df.iterrows():
                            try:
                                raw_oid = str(row[col_order_num]).strip()
                                if raw_oid.lower() in ('nan', '', 'none'): continue
                                # تحويل رقم الطلب لعدد صحيح إذا أمكن
                                try: oid = str(int(float(raw_oid)))
                                except: oid = raw_oid

                                original_amt  = float(row[col_order_amt])   if col_order_amt  and pd.notna(row.get(col_order_amt))  else 0.0
                                total_ded     = float(row[col_total_ded])   if col_total_ded  and pd.notna(row.get(col_total_ded))  else 0.0
                                transferred   = float(row[col_transferred]) if pd.notna(row.get(col_transferred)) else 0.0
                                date          = parse_date(row[col_date])   if col_date else None

                                # تحديد إذا كان مرتجع
                                row_type = str(row.get(col_type, '')).strip().lower() if col_type else ''
                                is_ret   = 1 if 'refund' in row_type else 0

                                if transferred != 0:
                                    collections_data.append((oid, original_amt, total_ded, transferred, date, is_ret))
                            except Exception as ex:
                                continue
                    else:
                        print(f"  -> Tabby: لم يتم العثور على أعمدة أساسية. الأعمدة: {list(df.columns)}")

            # ── 2. SMSA ──────────────────────────────────────────────────────
            # صلة الوصل: Ref No (رقم البوليصية)
            # المبلغ المحصل: COD Amount
            # عمولة التحصيل: COD Charges
            # ─────────────────────────────────────────────────────────────────
            elif 'smsa' in fname_lower or 'سمسا' in fname_lower:
                platform_name = "SMSA"
                df = read_file_with_header(path, header_row=2)
                if df is not None:
                    col_map = {c.strip().lower(): c for c in df.columns}

                    col_ref      = next((col_map[k] for k in col_map if 'ref no' in k or 'ref_no' in k), None)
                    col_cod      = next((col_map[k] for k in col_map if 'cod amount' in k), None)
                    col_cod_fee  = next((col_map[k] for k in col_map if 'cod charges' in k or 'cod charge' in k), None)
                    col_date     = next((col_map[k] for k in col_map if 'payment date' in k), None)

                    if col_ref and col_cod:
                        for _, row in df.iterrows():
                            try:
                                # Ref No = رقم البوليصية = صلة الوصل
                                raw_ref = str(row[col_ref]).strip()
                                if raw_ref.lower() in ('nan', '', 'none'): continue
                                try: oid = str(int(float(raw_ref)))
                                except: oid = raw_ref

                                cod_amount = float(row[col_cod]) if pd.notna(row.get(col_cod)) else 0.0
                                cod_fee    = float(row[col_cod_fee]) if col_cod_fee and pd.notna(row.get(col_cod_fee)) else 0.0
                                date       = parse_date(row[col_date]) if col_date and pd.notna(row.get(col_date)) else None

                                # فقط الطلبات التي تم تحصيلها (COD Amount > 0)
                                if cod_amount > 0:
                                    # صافي المحصل = COD Amount - COD Charges
                                    net_collected = cod_amount - cod_fee
                                    collections_data.append((oid, cod_amount, cod_fee, net_collected, date, 0))
                            except Exception as ex:
                                continue
                    else:
                        print(f"  -> SMSA: لم يتم العثور على أعمدة أساسية. الأعمدة: {list(df.columns)}")

            # ── 3. ILASOUQ ELECTRONIC ────────────────────────────────────────
            elif 'تحصيل-ilasouq' in fname_lower:
                platform_name = "Ilasouq Electronic"
                df = read_file_with_header(path, header_row=0)
                if df is not None:
                    col_amount = next((c for c in df.columns if 'بعد الضريبة' in c or 'إجمالي الطلب' in c), None)
                    col_oid    = next((c for c in df.columns if 'رقم الطلب' in c), None)
                    if col_amount and col_oid:
                        for _, row in df.iterrows():
                            try:
                                oid    = str(row[col_oid]).strip()
                                if oid.lower() in ('nan', ''): continue
                                amount = float(row[col_amount])
                                date   = datetime.now().date()
                                collections_data.append((oid, amount, 0.0, amount, date, 0))
                            except: continue

            # ── 4. NOON STATEMENT (CSV) ───────────────────────────────────────
            # صلة الوصل: order_nr
            # ─────────────────────────────────────────────────────────────────
            elif 'noon' in fname_lower or 'كشف-حساب-نون' in fname_lower or 'نون' in fname_lower:
                platform_name = "Noon Statement"
                print(f"Processing {filename} as {platform_name}...")
                df = read_file_with_header(path, header_row=0)
                if df is not None:
                    col_map = {c.strip().lower(): c for c in df.columns}

                    # order_nr هو صلة الوصل الأساسية
                    col_oid = next((col_map[k] for k in col_map if k == 'order_nr'), None)
                    if not col_oid:
                        col_oid = next((col_map[k] for k in col_map if 'order_nr' in k or 'order nr' in k), None)

                    col_amount = next((col_map[k] for k in col_map if 'total_payment' in k or 'payment' in k or 'amount' in k), None)
                    col_date   = next((col_map[k] for k in col_map if 'statement_date' in k or 'date' in k), None)

                    if col_oid:
                        for _, row in df.iterrows():
                            try:
                                raw_oid = str(row[col_oid]).strip()
                                if raw_oid.lower() in ('nan', '', 'none'): continue
                                # تنظيف رقم الطلب
                                oid = raw_oid.split()[0]
                                oid = ''.join(c for c in oid if c.isalnum() or c == '-')
                                if not oid: continue

                                amount = float(row[col_amount]) if col_amount and pd.notna(row.get(col_amount)) else 0.0
                                date   = parse_date(row[col_date]) if col_date and pd.notna(row.get(col_date)) else None

                                if amount != 0:
                                    collections_data.append((oid, amount, 0.0, amount, date, 0))
                            except: continue
                    else:
                        print(f"  -> Noon: لم يتم العثور على عمود order_nr. الأعمدة: {list(df.columns)}")

            # ── 5. TRENDYOL STATEMENT ─────────────────────────────────────────
            # يحتوي على Sales و Returns في نفس الملف
            # المرتجعات: قيمة سالبة في Credit أو Transaction Type = Refund
            # ─────────────────────────────────────────────────────────────────
            elif ('trendyol' in fname_lower or 'ترنديول' in fname_lower) and \
                 ('statement' in fname_lower or 'كشف' in fname_lower or 'عمليات' in fname_lower):
                if 'sales' in fname_lower or 'مبيعات' in fname_lower:
                    continue  # ملف المبيعات يُعالج في process_data.py
                platform_name = "Trendyol Statement"
                print(f"Processing {filename} as {platform_name}...")
                df = read_file_with_header(path, header_row=0)
                if df is not None:
                    col_map = {c.strip().lower(): c for c in df.columns}

                    col_oid    = next((col_map[k] for k in col_map if 'order number' in k), None)
                    col_credit = next((col_map[k] for k in col_map if k == 'credit'), None)
                    col_date   = next((col_map[k] for k in col_map if 'payment date' in k or 'settlement date' in k), None)
                    col_type   = next((col_map[k] for k in col_map if 'transaction type' in k or 'type' in k), None)

                    for _, row in df.iterrows():
                        try:
                            if not pd.notna(row.get(col_oid if col_oid else 'Order Number')): continue

                            raw_oid = row[col_oid] if col_oid else None
                            if raw_oid is None or pd.isna(raw_oid): continue
                            try: oid = str(int(float(str(raw_oid))))
                            except: oid = str(raw_oid).strip()

                            amount = float(row[col_credit]) if col_credit and pd.notna(row.get(col_credit)) else 0.0
                            date   = parse_date(str(row[col_date]).replace(' UTC', '')) if col_date and pd.notna(row.get(col_date)) else None

                            if amount == 0: continue

                            # تحديد نوع العملية: مبيعات أم مرتجع
                            tx_type = str(row.get(col_type, '')).strip().lower() if col_type else ''
                            # المرتجع: إما نوع Refund أو قيمة سالبة
                            is_ret = 1 if ('refund' in tx_type or 'return' in tx_type or amount < 0) else 0

                            collections_data.append((oid, abs(amount), 0.0, amount, date, is_ret))
                        except: continue

            # ── 6. AMAZON STATEMENT ───────────────────────────────────────────
            elif 'المعاملات' in fname_lower or 'amazon' in fname_lower:
                platform_name = "Amazon Statement"
                try:
                    lines = []
                    try:
                        with open(path, 'r', encoding='utf-8-sig') as f: lines = f.readlines()
                    except:
                        with open(path, 'r', encoding='cp1256') as f: lines = f.readlines()

                    header_idx = -1
                    for i, line in enumerate(lines[:20]):
                        if 'رقم الطلب' in line or 'Order ID' in line:
                            header_idx = i; break

                    if header_idx != -1:
                        headers = list(csv.reader([lines[header_idx]]))[0]
                        if len(headers) < 2 and ',' in lines[header_idx]:
                            headers = lines[header_idx].strip().replace('"', '').split(',')
                        norm_headers = [h.replace('"', '').strip() for h in headers]

                        oid_idx   = next((i for i, h in enumerate(norm_headers) if 'رقم الطلب' in h or 'Order ID' in h), -1)
                        total_idx = next((i for i, h in enumerate(norm_headers) if 'الإجمالي' in h or 'Total' in h), -1)
                        date_idx  = next((i for i, h in enumerate(norm_headers) if 'التاريخ' in h or 'Date' in h), -1)

                        if oid_idx != -1 and total_idx != -1:
                            reader = csv.reader(lines[header_idx+1:])
                            for row in reader:
                                if not row: continue
                                if len(row) == 1 and ',' in row[0]: row = row[0].split(',')
                                if len(row) <= max(oid_idx, total_idx): continue
                                raw_oid = row[oid_idx].replace('"', '').replace('=', '').strip()
                                if len(raw_oid) < 5 or '-' not in raw_oid: continue
                                raw_amt = row[total_idx].replace('"', '').replace('=', '').strip().replace(',', '')
                                try: amount = float(raw_amt)
                                except: continue
                                date_val = row[date_idx] if date_idx != -1 and len(row) > date_idx else None
                                date = parse_date(date_val)
                                if amount != 0:
                                    collections_data.append((raw_oid, abs(amount), 0.0, amount, date, 0))
                except Exception as e:
                    print(f"  Amazon Parse Error: {e}")

            # ── 7. WEBSITE (موقع خاص) ─────────────────────────────────────────
            elif any(k in fname_lower for k in ['website', 'موقع', 'site', 'web', 'store', 'متجر']):
                platform_name = "Website"
                df = read_file_with_header(path, header_row=0)
                if df is not None:
                    col_map = {c.strip().lower(): c for c in df.columns}
                    oid_col  = next((col_map[k] for k in col_map if any(x in k for x in ['order id', 'رقم الطلب', 'order_id'])), None)
                    amt_col  = next((col_map[k] for k in col_map if any(x in k for x in ['total', 'إجمالي', 'المبلغ', 'amount'])), None)
                    date_col = next((col_map[k] for k in col_map if any(x in k for x in ['date', 'تاريخ'])), None)
                    if oid_col and amt_col:
                        for _, row in df.iterrows():
                            try:
                                oid    = str(row[oid_col]).strip()
                                if not oid or oid.lower() == 'nan': continue
                                amount = float(row[amt_col])
                                date   = parse_date(row[date_col]) if date_col else datetime.now().date()
                                if amount != 0:
                                    collections_data.append((oid, abs(amount), 0.0, amount, date, 0))
                            except: continue

            # ── INSERT DATA ───────────────────────────────────────────────────
            if collections_data:
                inserted_count = 0
                for item in collections_data:
                    oid, orig_amt, fee, collected, dt, is_ret = item
                    if insert_collection(cursor, snapshot_id, oid, orig_amt, fee, collected, dt, is_ret, account):
                        inserted_count += 1

                conn.commit()
                if inserted_count > 0:
                    print(f"  -> [{platform_name}] Inserted {inserted_count} records (account: {account or 'N/A'}).")
                elif len(collections_data) > 0:
                    print(f"  -> [{platform_name}] {len(collections_data)} records found but already exist.")
                total_processed += inserted_count
            else:
                if platform_name != "Unknown":
                    print(f"  -> No valid records extracted from {platform_name}.")

        except Exception as e:
            print(f"Error processing {filename}: {e}")
            import traceback; traceback.print_exc()

    conn.close()
    print(f"\nTotal Collections Inserted: {total_processed}")


if __name__ == "__main__":
    process_collections()
