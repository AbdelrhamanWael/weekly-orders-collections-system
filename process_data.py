# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Module   : process_data.py — Orders Processing
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
#  Updated   : February 2026 — Snapshot support + Website platform
# ============================================================
import pandas as pd
import sqlite3
import os
from datetime import datetime

# Configuration
DB_NAME    = "finance_system.db"
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
    elif ext in ['.csv', '.txt']:
        encodings  = ['utf-8-sig', 'utf-8', 'cp1256', 'latin1']
        separators = [',', ';', '\t', None]
        for enc in encodings:
            for sep in separators:
                try:
                    df = pd.read_csv(file_path, encoding=enc, sep=sep, engine='python')
                    if len(df.columns) > 1: return normalize_columns(df)
                except: continue
    return None

def detect_account_name(filename):
    """
    يكتشف اسم الحساب من اسم الملف
    مثال: 'مبيعات-ترنديول-امواج.xlsx' -> 'أمواج'
           'طلبات-ilasouq.xlsx' -> 'ILASOQ'
    """
    fname = filename.lower()
    if 'امواج' in fname or 'amwaj' in fname:
        return 'أمواج'
    if 'ilasouq' in fname or 'ilasoq' in fname:
        return 'ILASOQ'
    return ''


def identify_platform(df, filename):
    columns    = [str(c).lower().strip() for c in df.columns]
    col_str    = " ".join(columns)
    fname_lower = filename.lower()

    # 1. Noon — order_nr هو صلة الوصل الأساسية
    if 'order_nr' in columns and 'order_status' in columns: return 'Noon'
    if 'order_nr' in columns and 'noon' in col_str:         return 'Noon'
    if 'order_nr' in columns:                               return 'Noon'
    if 'id_partner' in columns and 'statement_nr' in columns: return 'Noon'

    # 2. Ilasouq
    if 'ilasouq' in fname_lower:
        if 'تاريخ الطلب' in columns: return 'Ilasouq Orders'
        return 'Ilasouq Collection'
    if 'رقم الطلب' in columns and 'طريقة الدفع' in columns: return 'Ilasouq Orders'

    # 3. Trendyol
    if 'transaction no' in columns and 'storefront' in columns: return 'Trendyol'
    if 'الباركود' in columns and 'اسم المنتج' in columns:      return 'Trendyol Sales'

    # 4. Amazon
    if any('amazon' in c for c in columns): return 'Amazon'
    if any('نوع المعاملة' in c for c in columns) and any('رقم الطلب' in c for c in columns):
        return 'Amazon'

    # 5. Website (موقع خاص) — flexible detection
    if any(k in fname_lower for k in ['website', 'موقع', 'site', 'web', 'store', 'متجر']):
        return 'Website'
    if 'order_source' in columns and any('website' in str(v).lower() for v in df.get('order_source', [])):
        return 'Website'
    # Typical WooCommerce / custom store columns
    if all(c in columns for c in ['order id', 'order status', 'order total']):  return 'Website'
    if all(c in columns for c in ['رقم الطلب', 'حالة الطلب', 'إجمالي الطلب']): return 'Website'

    # 6. Tabby & SMSA (Collections only)
    if 'تابي' in fname_lower or 'tabby' in fname_lower: return 'Tabby'
    if 'سمسا' in fname_lower or 'smsa' in fname_lower:  return 'SMSA'

    return 'Unknown'


def process_file_content(df, platform, file_path, snapshot_id=0, account_name=''):
    print(f"  -> Identified as: {platform}")
    conn   = get_db_connection()
    cursor = conn.cursor()
    orders_data = []   # list of tuples

    try:
        # ── ILASOUQ ──────────────────────────────────────────
        if platform == 'Ilasouq Orders':
            for _, row in df.iterrows():
                oid      = str(row['رقم الطلب'])
                date_val = row['تاريخ الطلب']
                date     = date_val.date() if isinstance(date_val, (pd.Timestamp, datetime)) else parse_date(date_val)
                price    = float(row['إجمالي الطلب']) if pd.notna(row.get('إجمالي الطلب')) else 0.0
                shipping = float(row['تكلفة الشحن'])  if pd.notna(row.get('تكلفة الشحن'))  else 0.0
                tax      = float(row['الضريبة'])       if pd.notna(row.get('الضريبة'))       else 0.0

                # استخراج الأصناف من skus_json أولاً (أدق بكثير)
                items = ''
                skus_json_val = row.get('skus_json')
                if pd.notna(skus_json_val) and str(skus_json_val).strip() not in ('', 'nan'):
                    try:
                        import json, re
                        # skus_json: [["اسم المنتج", qty, sku, price, total], ...]
                        parsed = json.loads(str(skus_json_val))
                        parts = []
                        for entry in parsed:
                            if isinstance(entry, list) and len(entry) >= 2:
                                name = str(entry[0]).strip()
                                qty  = int(entry[1]) if entry[1] else 1
                                if name:
                                    parts.append(f"{name} x{qty}")
                        items = ' | '.join(parts)
                    except Exception:
                        items = ''

                # البديل: عمود 'اسماء المنتجات مع SKU'
                if not items:
                    col_items = 'اسماء المنتجات مع SKU'
                    raw = row.get(col_items)
                    if pd.notna(raw) and str(raw).strip() not in ('', 'nan'):
                        import re
                        # صيغة: '(SKU: )اسم المنتج(Qty: 2)'
                        matches = re.findall(r'\(SKU:[^)]*\)([^(]+)\(Qty:\s*(\d+)\)', str(raw))
                        if matches:
                            items = ' | '.join(f"{name.strip()} x{qty}" for name, qty in matches)
                        else:
                            # نص خام بدون تنسيق
                            items = str(raw).strip("'")

                orders_data.append((oid, 'Ilasouq', date, price, 0.0, shipping, 0.0, tax, items))

        # ── NOON ─────────────────────────────────────────────
        # order_nr هو صلة الوصل بين كشف البنك والطلبات
        elif platform == 'Noon':
            col_map = {c.strip().lower(): c for c in df.columns}
            col_oid = next((col_map[k] for k in col_map if k == 'order_nr'), None)
            if col_oid:
                # Noon CSV: order_nr, title (or title_ar), quantity, order_received_at
                col_date  = next((col_map[k] for k in col_map if k == 'order_received_at'), None) or \
                            next((col_map[k] for k in col_map if 'date' in k), None)
                col_item  = next((col_map[k] for k in col_map if k == 'title'), None) or \
                            next((col_map[k] for k in col_map if k == 'title_ar'), None)
                col_qty   = next((col_map[k] for k in col_map if k == 'quantity'), None)
                col_price = next((col_map[k] for k in col_map if 'total_price' in k or k == 'price'), None)
                for _, row in df.iterrows():
                    if pd.isna(row[col_oid]): continue
                    raw_oid = str(row[col_oid]).strip()
                    if not raw_oid or raw_oid.lower() == 'nan': continue
                    # نظافة order_nr: أخذ أول جزء فقط (NSAH60039264702)
                    oid = raw_oid.split(',')[0].strip()
                    oid = ''.join(c for c in oid if c.isalnum() or c == '-')
                    if not oid: continue
                    date_val = row[col_date] if col_date else None
                    # معالجة التاريخ من الصف المدمج (row 2)
                    if pd.isna(date_val) or str(date_val).strip() == 'nan':
                        # حاول استخراج التاريخ من raw_oid (Noon CSV المدمج)
                        import re
                        date_match = re.search(r'(\d{4}-\d{2}-\d{2})', raw_oid)
                        date_val = date_match.group(1) if date_match else None
                    date  = parse_date(date_val)
                    price = float(row[col_price]) if col_price and pd.notna(row.get(col_price)) else 0.0
                    # اسم المنتج + الكمية
                    items = ''
                    if col_item:
                        item_val = str(row.get(col_item, '')).strip()
                        # معالجة الصف المدمج: استخراج العنوان من raw_oid
                        if item_val in ('', 'nan', 'None'):
                            parts = raw_oid.split(',')
                            item_val = parts[14].strip() if len(parts) > 14 else ''
                        if item_val and item_val not in ('nan', 'None'):
                            qty = 1
                            if col_qty and pd.notna(row.get(col_qty)):
                                try: qty = int(float(row[col_qty]))
                                except: qty = 1
                            items = f"{item_val} x{qty}"
                    orders_data.append((oid, 'Noon', date, price, 0.0, 0.0, 0.0, 0.0, items))

        # ── TRENDYOL ─────────────────────────────────────────
        # يحتوي على Sales و Returns — نعالج كلاهما
        elif platform == 'Trendyol':
            col_map   = {c.strip().lower(): c for c in df.columns}
            col_oid   = next((col_map[k] for k in col_map if 'order number' in k), None)
            col_date  = next((col_map[k] for k in col_map if 'order date' in k), None)
            col_credit= next((col_map[k] for k in col_map if k == 'credit'), None)
            col_type  = next((col_map[k] for k in col_map if 'transaction type' in k), None)
            # اسم المنتج: عمود 'product name' أو 'storefront' أو أي عمود يحتوي على 'name'
            col_item  = next((col_map[k] for k in col_map if k == 'product name'), None) or \
                        next((col_map[k] for k in col_map if 'product' in k and 'name' in k), None) or \
                        next((col_map[k] for k in col_map if k == 'storefront'), None)
            col_qty   = next((col_map[k] for k in col_map if 'quantity' in k or 'qty' in k), None)
            for _, row in df.iterrows():
                tx_type = str(row.get(col_type, '')).strip() if col_type else ''
                if tx_type not in ('Sale', 'Refund', 'Return', ''): continue
                if col_oid and pd.notna(row.get(col_oid)):
                    try:
                        oid      = str(int(row[col_oid]))
                        date_str = str(row[col_date]).replace(' UTC', '') if col_date else ''
                        date     = parse_date(date_str)
                        price    = float(row[col_credit]) if col_credit and pd.notna(row.get(col_credit)) else 0.0
                        items    = ''
                        if col_item and pd.notna(row.get(col_item)):
                            qty   = 1
                            if col_qty and pd.notna(row.get(col_qty)):
                                try: qty = int(float(row[col_qty]))
                                except: qty = 1
                            items = f"{str(row[col_item]).strip()} x{qty}"
                        orders_data.append((oid, 'Trendyol', date, price, 0.0, 0.0, 0.0, 0.0, items))
                    except: continue

        # ── TRENDYOL SALES REPORT ──────────────────────────────
        # ملف مبيعات ترنديول (بالعربي) — لا يحتوي على order_id لكن يحتوي على أسماء المنتجات
        elif platform == 'Trendyol Sales':
            # هذا الملف للمرجع فقط (لا يتم إدخال طلبات منه)
            # لكن نستخرج منه ملخص الأصناف ونخزنه في الذاكرة لاستخدامه في التقرير
            print(f"  -> Trendyol Sales report: {len(df)} products found (for reference only).")
            return  # لا ندخل طلبات من هذا الملف

        # ── AMAZON ───────────────────────────────────────────
        elif platform == 'Amazon':
            oid_col  = next((c for c in df.columns if 'رقم الطلب' in c or 'Order ID' in c), None)
            date_col = next((c for c in df.columns if 'التاريخ' in c or 'Date' in c), None)

            extracted = False
            if oid_col:
                for _, row in df.iterrows():
                    raw_oid = str(row[oid_col]).replace('"', '').replace("'", "").replace('=', '').strip()
                    if len(raw_oid) < 5 or '-' not in raw_oid: continue
                    if not raw_oid[0].isdigit(): continue
                    date_val = row[date_col] if date_col else None
                    try: date = parse_date(date_val)
                    except: date = None
                    orders_data.append((raw_oid, 'Amazon', date, 0.0, 0.0, 0.0, 0.0, 0.0))
                    extracted = True

            if not extracted or len(orders_data) == 0:
                import csv
                try:
                    with open(file_path, 'r', encoding='utf-8-sig') as f:
                        lines = f.readlines()
                    header_idx = -1
                    for i, line in enumerate(lines):
                        if 'رقم الطلب' in line or 'Order ID' in line:
                            header_idx = i; break
                    if header_idx != -1:
                        headers      = list(csv.reader([lines[header_idx]], delimiter=','))[0]
                        norm_headers = [h.replace('"', '').strip() for h in headers]
                        if len(headers) == 1 and ',' in headers[0]:
                            headers      = headers[0].split(',')
                            norm_headers = [h.replace('"', '').strip() for h in headers]
                        try:
                            oid_idx  = next(i for i, h in enumerate(norm_headers) if 'رقم الطلب' in h or 'Order ID' in h)
                            date_idx = next((i for i, h in enumerate(norm_headers) if 'التاريخ' in h or 'Date' in h), -1)
                            reader   = csv.reader(lines[header_idx+1:], delimiter=',')
                            for row in reader:
                                if not row: continue
                                if len(row) == 1 and ',' in row[0]: row = row[0].split(',')
                                if len(row) <= oid_idx: continue
                                raw_oid = row[oid_idx].replace('"', '').replace("'", "").replace('=', '').strip()
                                if len(raw_oid) < 5 or '-' not in raw_oid: continue
                                if not raw_oid[0].isdigit(): continue
                                date_val = row[date_idx] if date_idx != -1 and len(row) > date_idx else None
                                date     = parse_date(date_val)
                                if not any(d[0] == raw_oid for d in orders_data):
                                    orders_data.append((raw_oid, 'Amazon', date, 0.0, 0.0, 0.0, 0.0, 0.0))
                            print(f"     Manual parsing found {len(orders_data)} orders.")
                        except StopIteration: pass
                except Exception: pass

        # ── WEBSITE (موقع خاص) ───────────────────────────────
        elif platform == 'Website':
            col_map   = {c.strip().lower(): c for c in df.columns}
            oid_col   = next((col_map[k] for k in col_map if any(x in k for x in ['order id', 'رقم الطلب', 'order_id'])), None)
            date_col  = next((col_map[k] for k in col_map if any(x in k for x in ['date', 'تاريخ'])), None)
            price_col = next((col_map[k] for k in col_map if any(x in k for x in ['total', 'إجمالي', 'المبلغ', 'price'])), None)
            ship_col  = next((col_map[k] for k in col_map if any(x in k for x in ['shipping', 'شحن'])), None)
            tax_col   = next((col_map[k] for k in col_map if any(x in k for x in ['tax', 'ضريبة'])), None)
            item_col  = next((col_map[k] for k in col_map if any(x in k for x in ['product', 'item', 'منتج', 'صنف'])), None)
            qty_col   = next((col_map[k] for k in col_map if any(x in k for x in ['quantity', 'qty', 'كمية'])), None)

            if oid_col:
                for _, row in df.iterrows():
                    raw_oid = str(row[oid_col]).replace('"', '').replace('=', '').strip()
                    if not raw_oid or raw_oid.lower() == 'nan': continue
                    date_val = row[date_col]  if date_col  else None
                    date     = parse_date(date_val)
                    price    = float(row[price_col]) if price_col and pd.notna(row.get(price_col)) else 0.0
                    shipping = float(row[ship_col])  if ship_col  and pd.notna(row.get(ship_col))  else 0.0
                    tax      = float(row[tax_col])   if tax_col   and pd.notna(row.get(tax_col))   else 0.0
                    items    = ''
                    if item_col and pd.notna(row.get(item_col)):
                        qty   = int(row[qty_col]) if qty_col and pd.notna(row.get(qty_col)) else 1
                        items = f"{row[item_col]} x{qty}"
                    orders_data.append((raw_oid, 'Website', date, price, 0.0, shipping, 0.0, tax, items))
            else:
                print("  -> Website file: could not detect Order ID column.")

    except Exception as e:
        print(f"    Error processing rows: {e}")

    # ── Batch Insert ──────────────────────────────────────────
    new_inserts = 0
    for data in orders_data:
        try:
            order_id   = data[0]
            order_date = data[2].isoformat() if data[2] else None
            wk         = get_week_number(data[2])
            yr         = data[2].year if data[2] else None
            items_sum  = data[8] if len(data) > 8 else ''
            cursor.execute('''
                INSERT OR IGNORE INTO orders
                (order_id, snapshot_id, platform, account_name, order_date,
                 price, cost, shipping, commission, tax, items_summary, week_number, year)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (order_id, snapshot_id, data[1], account_name, order_date,
                  data[3], data[4], data[5], data[6], data[7], items_sum, wk, yr))
            if cursor.rowcount > 0:
                new_inserts += 1
        except: continue

    conn.commit()
    conn.close()

    total = len(orders_data)
    if total > 0:
        print(f"  -> Processed {total} orders (New inserts: {new_inserts}).")
    else:
        print(f"  -> No valid orders extracted.")


def main(snapshot_id=0):
    print(f"Scanning directory: {SAMPLES_DIR}")
    if not os.path.exists(SAMPLES_DIR):
        print("Directory not found!")
        return

    for f in os.listdir(SAMPLES_DIR):
        if f.startswith('~$') or not os.path.isfile(os.path.join(SAMPLES_DIR, f)): continue
        path    = os.path.join(SAMPLES_DIR, f)
        account = detect_account_name(f)
        print(f"\nEvaluating: {f}" + (f" [حساب: {account}]" if account else ""))
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
            process_file_content(df, platform, path, snapshot_id=snapshot_id, account_name=account)


if __name__ == "__main__":
    main()
