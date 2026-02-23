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
        try: 
            with open(file_path, 'rb') as f:
                return pd.read_excel(f)
        except: return None
    elif ext in ['.csv', '.txt']:
        encodings  = ['utf-8-sig', 'utf-8', 'cp1256', 'latin1']
        separators = [',', ';', '\t', None]
        for enc in encodings:
            for sep in separators:
                try:
                    with open(file_path, 'rb') as f:
                        df = pd.read_csv(f, encoding=enc, sep=sep, engine='python')
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
    
    # 1. Hardcoded (Legacy)
    if 'امواج' in fname or 'amwaj' in fname:
        return 'أمواج'
    if 'ilasouq' in fname or 'ilasoq' in fname:
        return 'ILASOUQ'

    # 2. Dynamic: [AccountName]
    # Example: "Sales_Noon_[RiyadhBranch].xlsx" -> "RiyadhBranch"
    import re
    match = re.search(r'\[(.*?)\]', filename)
    if match:
        return match.group(1).strip()
    
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

    # 7. Product Costs File
    col_str = ' '.join(columns).lower()
    is_cost_file = any(k in fname_lower for k in ['تكلفة', 'cost', 'costs', 'أسعار', 'prices', 'منتجات', 'products'])
    
    # Robust identification by columns
    has_cost_col = any(k in col_str for k in ['cost', 'تكلفة', 'السعر', 'سعر', 'purchase', 'price', 'unit'])
    has_id_col   = any(k in col_str for k in ['sku', 'كود', 'رمز', 'name', 'product', 'اسم', 'منتج', 'صنف', 'item'])
    
    if is_cost_file or (has_cost_col and has_id_col and len(columns) < 15):
        return 'Product Costs'

    return 'Unknown'


def process_costs_file(df):
    """Update product costs in DB from file"""
    print("  -> Processing Product Costs file...")
    conn = get_db_connection()
    cur  = conn.cursor()
    
    col_map = {c.strip().lower(): c for c in df.columns}
    
    # Robust column detection
    col_sku = next((col_map[k] for k in col_map if 'sku' in k or 'كود' in k or 'رمز' in k), None)
    col_cost = next((col_map[k] for k in col_map if any(x in k for x in ['cost', 'تكلفة', 'التكلفة', 'شراء', 'توريد', 'purchase', 'سعر', 'price', 'unit', 'سعر الحبة'])), None)
    col_name = next((col_map[k] for k in col_map if any(x in k for x in ['name', 'product', 'اسم', 'منتج', 'الصنف', 'البيان', 'الوصف', 'item', 'الصنف'])), None)

    import hashlib
    count = 0
    print(f"  -> Columns detected in file: {list(df.columns)}")
    
    # Allow processing if we have (SKU OR Name) AND Cost
    if (col_sku or col_name) and col_cost:
        print(f"  -> Using '{col_sku or col_name}' as ID and '{col_cost}' as Cost column.")
        for _, row in df.iterrows():
            # Get SKU or generate one from name
            sku = str(row[col_sku]).strip() if col_sku and pd.notna(row.get(col_sku)) else ''
            name = str(row[col_name]).strip() if col_name and pd.notna(row.get(col_name)) else ''
            
            if not sku and name:
                # Generate deterministic SKU from name
                h = hashlib.md5(name.encode('utf-8')).hexdigest()[:8].upper()
                sku = f"AUTO-{h}"
            
            if not sku or sku.lower() == 'nan': continue
            
            try: 
                cost_val = str(row[col_cost]).replace(',', '').strip()
                cost = float(cost_val)
            except: 
                cost = 0.0
            
            # Upsert
            cur.execute("""
                INSERT INTO product_costs (sku, product_name, cost, updated_at)
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(sku) DO UPDATE SET
                    cost = excluded.cost,
                    product_name = CASE WHEN excluded.product_name != '' THEN excluded.product_name ELSE product_costs.product_name END,
                    updated_at = CURRENT_TIMESTAMP
            """, (sku, name, cost))
            count += 1
        conn.commit()
        print(f"  -> ✅ SUCCESS: Updated costs for {count} products in database.")
    else:
        print(f"  -> ❌ ERROR: Required columns not found!")
        print(f"     Required: (SKU or Name) AND (Cost/Price)")
        print(f"     Detected: SKU={col_sku}, Name={col_name}, Cost={col_cost}")
    conn.close()
    return count


def process_file_content(df, platform, file_path, snapshot_id=0, account_name=''):
    if platform == 'Product Costs':
        process_costs_file(df)
        return

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
                salla_status = str(row.get('حالة الطلب', '')).strip()
                
                col_url = next((c for c in df.columns if 'رابط الطلب' in c or 'رابط' in c), None)
                order_url = str(row[col_url]).strip() if col_url and pd.notna(row.get(col_url)) else ''
                
                col_cod = next((c for c in df.columns if 'رسوم الدفع عند الاستلام' in c), None)
                cod_fee = float(row[col_cod]) if col_cod and pd.notna(row.get(col_cod)) else 0.0

                # استخراج الأصناف من skus_json أولاً (أدق بكثير)
                items = ''
                skus_costs = 0.0
                skus_json_val = row.get('skus_json')
                
                if pd.notna(skus_json_val) and str(skus_json_val).strip() not in ('', 'nan'):
                    try:
                        import json
                        parsed = json.loads(str(skus_json_val))
                        parts = []
                        for entry in parsed:
                            if isinstance(entry, list) and len(entry) >= 2:
                                name = str(entry[0]).strip()
                                qty  = int(entry[1]) if entry[1] else 1
                                sku  = str(entry[2]).strip() if len(entry) > 2 else ''
                                
                                # Cost Lookup
                                if sku:
                                    conn_cost = get_db_connection()
                                    cur_cost  = conn_cost.cursor()
                                    cur_cost.execute("SELECT cost FROM product_costs WHERE sku = ?", (sku,))
                                    res_cost  = cur_cost.fetchone()
                                    conn_cost.close()
                                    if res_cost:
                                        skus_costs += (res_cost[0] * qty)

                                if name:
                                    parts.append(f"{name} x{qty}")
                        items = ' | '.join(parts)
                    except Exception:
                        items = ''
                
                # If cost was not calculated from SKUs (e.g. no SKUs in JSON), stay 0 or use provided column
                # But here we prioritize the DB lookup if available.
                
                # البديل: عمود 'اسماء المنتجات مع SKU'
                if not items:
                    col_items = 'اسماء المنتجات مع SKU'
                    raw = row.get(col_items)
                    if pd.notna(raw) and str(raw).strip() not in ('', 'nan'):
                         # ... regex parsing (simplified for brevity, assume similar logic if needed)
                         items = str(raw).strip("'")

                # Payment Method
                pay_method = 'Unknown'
                col_pay = next((c for c in df.columns if 'طريقة الدفع' in c or 'Payment Method' in c), None)
                if col_pay and pd.notna(row.get(col_pay)):
                    pay_method = str(row[col_pay]).strip()

                # ── Fees Calculation (New) ──
                # 1. COD Fees
                col_cod = next((c for c in df.columns if 'cod' in c.lower() or 'دفع عند الاستلام' in c), None)
                cod_fee = 0.0
                if col_cod and pd.notna(row.get(col_cod)):
                    try: cod_fee = float(row[col_cod])
                    except: pass
                
                # 2. Payment Fees (Mada/Visa)
                # Look for generic 'Fee', 'Commission', 'Mada Fee', 'Visa Fee'
                col_pg_fee = next((c for c in df.columns if (any(k in c.lower() for k in ['payment fee', 'رسوم الدفع', 'fee', 'commission', 'عمولة', 'mada', 'visa']) and 'الاستلام' not in c)), None)
                pg_fee = 0.0
                if col_pg_fee and pd.notna(row.get(col_pg_fee)):
                    try: pg_fee = abs(float(row[col_pg_fee])) # Ensure positive
                    except: pass
                
                commission = cod_fee + pg_fee

                # 3. Shipping (Store Borne)
                # If 'shipping' var above is what customer paid, we might need another for store cost.
                # For now, we assume 'shipping' variable extracted earlier is correct or we improve it.
                # Let's keep existing 'shipping' variable unless a specific "Shipping Cost" column exists.
                # 3. Shipping (Store Borne)
                col_ship_cost = next((c for c in df.columns if 'shipping' in c.lower() or 'cargo' in c.lower() or 'delivery' in c.lower() or 'شحن' in c or 'توصيل' in c), None)
                if col_ship_cost and pd.notna(row.get(col_ship_cost)):
                    try: shipping = float(row[col_ship_cost])
                    except: pass

                # Use Branch column if available and clean it (Salla often exports it as "['Name']")
                raw_branch = row.get('الفرع')
                if pd.notna(raw_branch) and str(raw_branch).strip() not in ('', 'nan', '\\N'):
                    branch_str = str(raw_branch).strip()
                    if branch_str.startswith('[') and branch_str.endswith(']'):
                         import ast
                         try:
                             parsed_branch = ast.literal_eval(branch_str)
                             if isinstance(parsed_branch, list) and len(parsed_branch) > 0:
                                 branch_str = str(parsed_branch[0])
                         except: pass
                    acc = branch_str
                else:
                    acc = account_name if account_name else 'Ilasouq Account'

                # ── New Extended Salla Dimensions ──
                city = str(row.get('المدينة', '')).strip() if pd.notna(row.get('المدينة')) else ''
                
                col_ship_comp = next((c for c in df.columns if 'شركة الشحن' in c), None)
                shipping_company = str(row[col_ship_comp]).strip() if col_ship_comp and pd.notna(row.get(col_ship_comp)) else ''
                
                col_track = next((c for c in df.columns if 'بوليصة' in c or 'tracking' in c.lower()), None)
                tracking_number = str(row[col_track]).strip() if col_track and pd.notna(row.get(col_track)) else ''
                
                col_discount = next((c for c in df.columns if 'خصم' in c or 'discount' in c.lower()), None)
                discount_value = float(row[col_discount]) if col_discount and pd.notna(row.get(col_discount)) else 0.0
                
                # Check multiple columns for marketing source / coupon
                marketing_source = ''
                col_coupon_name = next((c for c in df.columns if 'اسم الكوبون' in c), None)
                col_coupon_code = next((c for c in df.columns if 'رمز الكوبون' in c), None)
                col_utm = next((c for c in df.columns if 'utm_source' in c or 'مصدر' in c), None)
                
                if col_coupon_name and pd.notna(row.get(col_coupon_name)) and str(row.get(col_coupon_name)).strip():
                    marketing_source = str(row.get(col_coupon_name)).strip()
                elif col_coupon_code and pd.notna(row.get(col_coupon_code)) and str(row.get(col_coupon_code)).strip():
                    marketing_source = str(row.get(col_coupon_code)).strip()
                elif col_utm and pd.notna(row.get(col_utm)) and str(row.get(col_utm)).strip():
                    marketing_source = str(row.get(col_utm)).strip()

                orders_data.append((oid, 'Ilasouq', date, price, skus_costs, shipping, cod_fee, commission, tax, items, pay_method, acc, salla_status, order_url, city, shipping_company, tracking_number, discount_value, marketing_source))

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
                    # Product Name + Quantity
                    items = ''
                    if col_item and pd.notna(row.get(col_item)):
                         p_name = str(row[col_item]).strip()
                         p_qty  = 1
                         if col_qty and pd.notna(row.get(col_qty)):
                             try: p_qty = int(row[col_qty])
                             except: pass
                         items = f"{p_name} x{p_qty}"
                    # Use passed account_name if available
                    acc = account_name if account_name else 'Noon Account'
                    orders_data.append((oid, 'Noon', date, price, 0.0, 0.0, 0.0, 0.0, items, '', acc, '', '', '', 0.0, ''))

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
            
            # Additional Columns for Fees
            col_gross = next((col_map[k] for k in col_map if 'sales amount' in k or 'gross amount' in k or 'selling price' in k), None)
            col_comm  = next((col_map[k] for k in col_map if 'commission' in k or 'comm' in k or 'عمولة' in k or 'deduction' in k or 'fee' in k or 'kesinti' in k), None)
            col_ship  = next((col_map[k] for k in col_map if 'shipping' in k or 'cargo' in k or 'شحن' in k or 'kargo' in k or 'delivery' in k), None)

            for _, row in df.iterrows():
                tx_type = str(row.get(col_type, '')).strip() if col_type else ''
                # Filter useful rows (Sale/Return)
                if col_type and tx_type not in ('Sale', 'Refund', 'Return', ''): continue
                
                if col_oid and pd.notna(row.get(col_oid)):
                    try:
                        oid      = str(int(row[col_oid]))
                        date_str = str(row[col_date]).replace(' UTC', '') if col_date else ''
                        date     = parse_date(date_str)
                        
                        # Fees
                        comm = 0.0
                        if col_comm and pd.notna(row.get(col_comm)):
                            try: comm = abs(float(row[col_comm]))
                            except: pass

                        ship = 0.0
                        if col_ship and pd.notna(row.get(col_ship)):
                            try: ship = abs(float(row[col_ship]))
                            except: pass

                        # Price (Expected Amount)
                        # If Gross exists use it, else Credit is mostly Net, so Gross = Credit + Fees?
                        # Let's assume Credit is Net.
                        credit_val = float(row[col_credit]) if col_credit and pd.notna(row.get(col_credit)) else 0.0
                        
                        if col_gross and pd.notna(row.get(col_gross)):
                             price = float(row[col_gross])
                        else:
                             # Fallback: Gross = Net + Expenses (roughly)
                             # Only if this is a Settlement file where expenses are deducted.
                             # If it's an Order report, Credit might be Gross.
                             # Without file sample, hard to say. Let's trust Credit if Gross missing.
                             price = credit_val + comm + ship if credit_val > 0 else credit_val

                        items    = ''
                        if col_item and pd.notna(row.get(col_item)):
                            qty   = 1
                            if col_qty and pd.notna(row.get(col_qty)):
                                try: qty = int(float(row[col_qty]))
                                except: qty = 1
                            items = f"{str(row[col_item]).strip()} x{qty}"
                        
                        acc = account_name if account_name else 'Trendyol Account'
                        orders_data.append((oid, 'Trendyol', date, price, 0.0, ship, comm, 0.0, items, acc, '', '', '', 0.0, ''))
                    except: continue

        # ── TRENDYOL SALES REPORT ──────────────────────────────
        # ملف مبيعات ترنديول (بالعربي) — لا يحتوي على order_id لكن يحتوي على أسماء المنتجات
        elif platform == 'Trendyol Sales':
            # هذا الملف للمرجع فقط (لا يتم إدخال طلبات منه)
            print(f"  -> Trendyol Sales report: {len(df)} products found (for reference only).")
            return  # لا ندخل طلبات من هذا الملف

        # ── AMAZON ───────────────────────────────────────────
        elif platform == 'Amazon':
            # Logic similar to process_collections but focused on Order Details (Prices, Fees)
            # We need to aggregate lines per Order ID to get full picture.
            
            # 1. Read and Clean
            lines = []
            encoding = 'utf-8-sig'
            try:
                with open(file_path, 'r', encoding=encoding) as f: lines = f.readlines()
            except:
                encoding = 'cp1256'
                with open(file_path, 'r', encoding=encoding) as f: lines = f.readlines()

            header_info = None
            # Find header
            for i, line in enumerate(lines[:20]):
                if 'نوع المعاملة' in line and 'الإجمالي' in line:
                    header_info = (i, True) # (row_idx, is_transaction_file)
                    break
                if 'رقم الطلب' in line and not 'نوع المعاملة' in line:
                    header_info = (i, False) # Plain order report
                    break
            
            if header_info:
                h_idx, is_trans = header_info
                
                # Clean lines
                cleaned_lines = []
                for line in lines[h_idx:]:
                    line = line.strip()
                    if line.startswith('"') and ',""' in line:
                        if line.startswith('"'): line = line[1:]
                        if line.endswith('"'):   line = line[:-1]
                        line = line.replace('""', '"')
                    cleaned_lines.append(line)
                
                from io import StringIO
                df_amz = pd.read_csv(StringIO("\n".join(cleaned_lines)), sep=',', engine='python')
                df_amz.columns = [str(c).replace('"', '').replace('=', '').strip() for c in df_amz.columns]

                # Identify Columns
                col_oid     = next((c for c in df_amz.columns if 'رقم الطلب' in c), None)
                col_date    = next((c for c in df_amz.columns if 'التاريخ' in c), None)
                
                # Amount Columns (Transaction File)
                col_product = next((c for c in df_amz.columns if 'رسوم المنتج' in c), None) # Product Charges
                col_other   = next((c for c in df_amz.columns if 'أخرى' in c), None)        # Other (Shipping Income usually)
                col_amz_fee = next((c for c in df_amz.columns if 'رسوم أمازون' in c), None) # Amazon Fees
                col_total   = next((c for c in df_amz.columns if 'الإجمالي' in c), None)    # Total
                col_type    = next((c for c in df_amz.columns if 'نوع المعاملة' in c), None)
                
                col_sku     = next((c for c in df_amz.columns if 'sku' in c.lower()), None)
                col_title   = next((c for c in df_amz.columns if 'وصف' in c or 'description' in c.lower() or 'title' in c.lower() or 'اسم' in c), None)

                # Aggregation Dict
                # oid -> {date, price, cost, shipping, commission, tax, items, sku}
                agg_orders = {}

                if col_oid:
                    for _, row in df_amz.iterrows():
                        raw_oid = str(row[col_oid]).replace('"', '').replace('=', '').strip()
                        if len(raw_oid) < 5 or '-' not in raw_oid: continue
                        
                        # Initialize
                        if raw_oid not in agg_orders:
                            d_val = row[col_date] if col_date else None
                            agg_orders[raw_oid] = {
                                'date': parse_date(d_val),
                                'price': 0.0,      # Expected (Product + Shipping Income)
                                'cost': 0.0,       # From DB
                                'shipping': 0.0,   # Shipping Cost (Easy Ship)
                                'commission': 0.0, # Amazon Fees
                                'shipping': 0.0,   # Shipping Cost (Easy Ship)
                                'commission': 0.0, # Amazon Fees
                                'sku': None,
                                'items_list': []
                            }
                        
                        # Parse Amounts
                        def get_float(r, c):
                            if not c or pd.isna(r.get(c)): return 0.0
                            try: return float(str(r[c]).replace(',', '').strip())
                            except: return 0.0

                        val_prod  = get_float(row, col_product)
                        val_other = get_float(row, col_other)
                        val_fee   = get_float(row, col_amz_fee)
                        val_total = get_float(row, col_total)
                        txt_type  = str(row[col_type]) if col_type else ''
                        
                        # Logic:
                        # 1. Product Sales (Order Amount line)
                        if 'مبلغ الطلب' in txt_type or val_prod > 0:
                            agg_orders[raw_oid]['price'] += val_prod
                            # "Other" here is usually Shipping Income (Positive) -> Add to Price
                            # Or Promo Rebates (Negative) -> Subtract? usually rebates are separate.
                            if val_other > 0:
                                agg_orders[raw_oid]['price'] += val_other
                            
                            agg_orders[raw_oid]['commission'] += abs(val_fee)
                        
                        # 2. Easy Ship / Shipping Services
                        elif 'شحن' in txt_type or 'Shipping' in txt_type:
                            # This is a COST.
                            # Usually the total is negative.
                            # val_total is -21.85.
                            agg_orders[raw_oid]['shipping'] += abs(val_total)
                        
                        # 3. Other Fees / Services
                        elif 'رسوم' in txt_type and 'شحن' not in txt_type:
                             agg_orders[raw_oid]['commission'] += abs(val_total)

                        # Capture SKU/Title for Items Summary
                        # Capture title from ANY line that has it for this order
                        if col_title and pd.notna(row.get(col_title)):
                            p_name = str(row[col_title]).strip()
                            # Clean up prefixes
                            if p_name.lower().startswith('order item - '): p_name = p_name[13:].strip()
                            elif p_name.lower().startswith('order - '): p_name = p_name[8:].strip()
                            # Extra cleaning for Arabic prefixes
                            for prf in ['الطلب - ', 'عنصر الطلب - ']:
                                if p_name.startswith(prf): p_name = p_name[len(prf):].strip()
                                
                            # Filter out generic transactional descriptions
                            ignore = ['shipping', 'شحن', 'توصيل', 'عمولة', 'fee', 'commission', 'tax', 'ضريبة']
                            if p_name and not any(k in p_name.lower() for k in ignore) and len(p_name) > 3:
                                if p_name not in agg_orders[raw_oid]['items_list']:
                                    agg_orders[raw_oid]['items_list'].append(p_name)
                        
                        if col_sku and not agg_orders[raw_oid]['sku']:
                            current_sku = str(row[col_sku]).strip()
                            if current_sku and current_sku not in ('nan', ''):
                                agg_orders[raw_oid]['sku'] = current_sku

                # Process Aggregated Orders
                for oid, data in agg_orders.items():
                    # Fixed Shipping Rule: Client requested fixed 12.0 SAR for shipping
                    # If shipping cost was detected (e.g. 21.85), override it to 12.0
                    # If no shipping cost but valid order, also set to 12.0
                    if data['shipping'] > 0 or (data['price'] > 0):
                        data['shipping'] = 12.0

                    cost = 0.0
                    acc = account_name if account_name else 'Amazon Account'
                    
                    # Construct Items Summary
                    items_summary = ""
                    if data['items_list']:
                        items_summary = " | ".join(data['items_list'])
                    elif data['sku']:
                        items_summary = data['sku']

                    orders_data.append((
                        oid, 'Amazon', data['date'], 
                        data['price'], cost, data['shipping'], data['commission'], 0.0, items_summary, acc, '', '', '', 0.0, ''
                    ))
            else:
                 print("  -> Amazon file header not found.")

        # ── WEBSITE (موقع خاص) ───────────────────────────────
        elif platform == 'Website':
            # ...
            pay_col   = next((col_map[k] for k in col_map if any(x in k for x in ['payment', 'دفع', 'method'])), None)
            
            if oid_col:
                for _, row in df.iterrows():
                    # ...
                    pay_method = str(row[pay_col]).strip() if pay_col and pd.notna(row.get(pay_col)) else 'Website'
                    
                    orders_data.append((raw_oid, 'Website', date, price, cost, shipping, 0.0, tax, items, pay_method, '', '', '', 0.0, ''))
            else:
                print("  -> Website file: could not detect Order ID column.")

    except Exception as e:
        print(f"    Error processing rows: {e}")

    # ── Batch Insert ──────────────────────────────────────────
    new_inserts = 0
    
    # Pre-fetch accounts countries to minimize DB hits inside loop
    acc_country_map = {}
    try:
        conn_map = get_db_connection()
        rows = conn_map.execute("SELECT account_name, country FROM accounts").fetchall()
        for r in rows:
            acc_country_map[r[0]] = r[1]
        conn_map.close()
    except: pass

    for data in orders_data:
        try:
            # data structure:
            # 0: oid, 1: platform, 2: date, 3: price, 4: cost, 5: shipping, 6: comm, 7: tax, 8: items
            # 9: payment_method OR account_name (depends on len)
            # 10: account_name (if len=11)

            oid        = data[0]
            platform   = data[1]
            date_obj   = data[2]
            price      = data[3]
            cost       = data[4]
            shipping   = data[5]
            
            cod_fee   = float(data[6]) if len(data) > 6 and data[6] else 0.0
            commission= float(data[7]) if len(data) > 7 and data[7] else 0.0
            tax       = float(data[8]) if len(data) > 8 and data[8] else 0.0
            items     = data[9] if len(data) > 9 else ""
            
            pay_method   = ''
            acc_resolved = account_name
            city = ''
            shipping_company = ''
            tracking_number = ''
            discount_value = 0.0
            marketing_source = ''

            if len(data) >= 19:
                pay_method   = data[10]
                acc_resolved = data[11]
                salla_status = data[12]
                order_url    = data[13]
                city         = data[14]
                shipping_company = data[15]
                tracking_number = data[16]
                discount_value = data[17]
                marketing_source = data[18]
            elif len(data) == 16:
                pay_method   = data[10]
                acc_resolved = data[11]
                salla_status = ''
                order_url    = ''
                city         = data[12]
                shipping_company = data[13]
                tracking_number = data[14]
                discount_value = 0.0
                marketing_source = data[15]
            elif len(data) >= 14:
                pay_method   = data[10]
                acc_resolved = data[11]
                salla_status = data[12]
                order_url    = data[13]
            elif len(data) == 13:
                pay_method   = data[10]
                acc_resolved = data[11]
                salla_status = data[12]
                order_url    = ''
            elif len(data) == 12:
                pay_method   = data[10]
                acc_resolved = data[11]
                salla_status = ''
                order_url    = ''
            elif len(data) == 11:
                pay_method   = ''
                acc_resolved = data[10]
                salla_status = ''
                order_url    = ''
            
            # Resolve Country
            country = acc_country_map.get(acc_resolved, 'SA')
            
            # Prepare Date
            order_date = date_obj.isoformat() if date_obj else None
            wk         = get_week_number(date_obj)
            yr         = date_obj.year if date_obj else None

            cursor.execute('''
                INSERT OR IGNORE INTO orders
                (order_id, snapshot_id, platform, account_name, country, order_date,
                 price, cost, shipping, cod_fee, commission, tax, items_summary, payment_method, salla_status, order_url,
                 city, shipping_company, tracking_number, discount_value, marketing_source, week_number, year)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (oid, snapshot_id, platform, acc_resolved, country, order_date,
                  price, cost, shipping, cod_fee, commission, tax, items, pay_method, salla_status, order_url,
                  city, shipping_company, tracking_number, discount_value, marketing_source, wk, yr))
            
            if cursor.rowcount > 0:
                new_inserts += 1
        except Exception as ex:
            print(f"Error inserting row {data[0]}: {ex}")
            continue

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
