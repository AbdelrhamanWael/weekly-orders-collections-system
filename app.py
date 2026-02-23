"""
# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
#  Version   : 2.0 (Flask Edition)
# ============================================================
#  Description:
#    Flask backend for the weekly reconciliation system.
#    Handles file uploads, processing pipeline, report generation,
#    and serves the single-page Arabic RTL dashboard.
# ============================================================
"""

from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
import os, sys, sqlite3, json, shutil, re, threading
from datetime import datetime
from werkzeug.utils import secure_filename

# ── When running as .exe (PyInstaller), use folder of the executable ─────────
if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
    os.chdir(BASE_DIR)  # so process_data / process_collections use correct paths
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import process_data
import process_collections
import generate_report
import init_db

app = Flask(__name__)
app.secret_key = "weekly_reconciliation_2026"

UPLOAD_FOLDER  = os.path.join(BASE_DIR, "samples")
REPORTS_FOLDER = os.path.join(BASE_DIR, "reports")
DB_PATH        = os.path.join(BASE_DIR, "finance_system.db")
ALLOWED_EXT    = {".csv", ".xlsx", ".xls", ".txt"}

os.makedirs(UPLOAD_FOLDER,  exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)

# إنشاء قاعدة البيانات والجداول تلقائياً عند أول تشغيل (مهم للعميل عند تشغيل الـ .exe)
init_db.create_database()


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def get_db():
    # Adding timeout to prevent 'database is locked' errors during concurrent access
    conn = sqlite3.connect(DB_PATH, timeout=20)
    conn.row_factory = sqlite3.Row
    return conn


def get_active_snapshot_id():
    """Return the current active snapshot_id (highest id, or 0 if none)."""
    try:
        conn = get_db()
        row  = conn.execute(
            "SELECT snapshot_id FROM weekly_snapshots ORDER BY snapshot_id DESC LIMIT 1"
        ).fetchone()
        conn.close()
        return row[0] if row else 0
    except:
        return 0


def db_stats(snapshot_id=None):
    """Return quick counts from the database for a given snapshot."""
    try:
        conn = get_db()
        cur  = conn.cursor()
        if snapshot_id is None:
            snapshot_id = get_active_snapshot_id()
        cur.execute("SELECT COUNT(*) FROM orders WHERE snapshot_id=?", (snapshot_id,))
        orders = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM collections WHERE snapshot_id=?", (snapshot_id,))
        collections = cur.fetchone()[0]
        cur.execute("SELECT COALESCE(SUM(price),0) FROM orders WHERE snapshot_id=?", (snapshot_id,))
        total_expected = cur.fetchone()[0]
        cur.execute("SELECT COALESCE(SUM(collected_amount),0) FROM collections WHERE snapshot_id=?", (snapshot_id,))
        total_collected = cur.fetchone()[0]
        conn.close()
        rate = round(total_collected / total_expected * 100, 1) if total_expected else 0
        return {
            "orders": orders,
            "collections": collections,
            "total_expected": round(total_expected, 2),
            "total_collected": round(total_collected, 2),
            "collection_rate": rate,
            "snapshot_id": snapshot_id,
        }
    except Exception as e:
        return {"orders": 0, "collections": 0, "total_expected": 0,
                "total_collected": 0, "collection_rate": 0, "error": str(e)}


def get_all_accounts():
    """Fetch all configured accounts from DB."""
    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT id, platform_name, account_name, country, "
            "COALESCE(fixed_shipping_cost, 0) as fixed_shipping_cost, "
            "COALESCE(cost_includes_tax, 0) as cost_includes_tax, "
            "created_at FROM accounts ORDER BY platform_name, account_name"
        ).fetchall()
        conn.close()
        return [dict(r) for r in rows]
    except:
        return []


def platform_breakdown(snapshot_id=None):
    try:
        conn = get_db()
        cur  = conn.cursor()
        if snapshot_id is None:
            snapshot_id = get_active_snapshot_id()
        cur.execute("""
            SELECT o.platform,
                   COUNT(o.order_id)                        AS orders,
                   COALESCE(SUM(o.price),0)                 AS expected,
                   COALESCE(SUM(o.cost),0)                  AS cost,
                   COALESCE(SUM(c.collected_amount),0)      AS collected
            FROM orders o
            LEFT JOIN collections c
                ON o.order_id = c.order_id AND c.snapshot_id = ?
            WHERE o.snapshot_id = ?
            GROUP BY o.platform
        """, (snapshot_id, snapshot_id))
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        for r in rows:
            r["rate"]      = round(r["collected"] / r["expected"] * 100, 1) if r["expected"] else 0
            r["expected"]  = round(r["expected"],  2)
            r["cost"]      = round(r["cost"], 2)
            r["collected"] = round(r["collected"], 2)
        return rows
    except Exception as e:
        print(f"Error in platform_breakdown: {e}")
        return []


def list_uploaded_files():
    files = []
    for f in os.listdir(UPLOAD_FOLDER):
        if f.startswith("~$"): continue
        ext = os.path.splitext(f)[1].lower()
        if ext in ALLOWED_EXT:
            path = os.path.join(UPLOAD_FOLDER, f)
            files.append({
                "name": f,
                "size": round(os.path.getsize(path) / 1024, 1),
                "modified": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M"),
            })
    return sorted(files, key=lambda x: x["modified"], reverse=True)


def list_reports():
    reports = []
    for f in os.listdir(REPORTS_FOLDER):
        if f.endswith(".xlsx"):
            path = os.path.join(REPORTS_FOLDER, f)
            reports.append({
                "name": f,
                "size": round(os.path.getsize(path) / 1024, 1),
                "modified": datetime.fromtimestamp(os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M"),
            })
    return sorted(reports, key=lambda x: x["modified"], reverse=True)


# ─────────────────────────────────────────────────────────────────────────────
# Routes
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/products")
def products():
    sid = get_active_snapshot_id()
    conn = get_db()
    
    # 1. Fetch all known costs
    try:
        known = conn.execute("SELECT sku, product_name, cost FROM product_costs").fetchall()
    except:
        known = []
        
    # Map encoded Name -> {sku, cost}. We use encoded name to handle potential duplicates or strict matching.
    # We use lowercase keys for matching, but keep original name for display
    cost_map = {}
    known_skus = set()
    for r in known:
        name = r['product_name'] or ''
        sku  = r['sku']
        cost = r['cost']
        if name: 
            cost_map[name.strip().lower()] = {'sku': sku, 'cost': cost, 'name': name.strip()}
        known_skus.add(sku)
        
    # 2. Aggregate from Orders (Active Snapshot)
    items_found = set()
    import re
    if sid:
        try:
            rows = conn.execute("SELECT items_summary FROM orders WHERE snapshot_id=?", (sid,)).fetchall()
            for r in rows:
                summary = r[0]
                if not summary: continue
                parts = summary.split(' | ')
                for part in parts:
                    part = part.strip()
                    if not part: continue
                    m = re.match(r'^(.+?)\s+x(\d+)$', part, re.IGNORECASE)
                    if m: name = m.group(1).strip()
                    else: name = part
                    if name: items_found.add(name)
        except Exception as e:
            print(f"Error fetching items: {e}")

    # 3. Merge Lists
    final_list = []
    
    # Process items_found to separate those that are already known (case-insensitive check)
    unknown_items = []
    
    # Check each found item against cost_map
    found_map = {} # map lower -> original found name
    for name in items_found:
        found_map[name.lower()] = name
        
    # Add Known Items
    for lower_name, data in cost_map.items():
        # If this known item was found in orders, remove it from found_map so we don't add it again
        if lower_name in found_map:
            del found_map[lower_name]
        final_list.append((data['sku'], data['name'], data['cost']))

    # Add remaining (Unknown) Items
    import hashlib
    for lower_name, name in found_map.items():
        # Generate deterministic placeholder SKU
        h = hashlib.md5(name.encode('utf-8')).hexdigest()[:8].upper()
        sku_dummy = f"AUTO-{h}"
        # Avoid collision with real SKUs (rare but possible)
        while sku_dummy in known_skus:
            h = hashlib.md5((name + "1").encode('utf-8')).hexdigest()[:8].upper()
            sku_dummy = f"AUTO-{h}"
        
        final_list.append((sku_dummy, name, 0.0))
        
    conn.close()
    final_list.sort(key=lambda x: x[1]) # Sort by Name
    return render_template("products.html", products=final_list)

def normalize_text(text):
    """Normalize Arabic/English text for better matching.
    Removes diacritics, extra spaces, and converts to lowercase."""
    if not text:
        return ""
    import unicodedata
    # Normalize unicode (remove diacritics)
    text = unicodedata.normalize('NFKD', text)
    # Remove Arabic diacritics (tashkeel)
    arabic_diacritics = re.compile(r'[\u064B-\u065F\u0670]')
    text = arabic_diacritics.sub('', text)
    # Convert to lowercase and strip
    text = text.strip().lower()
    # Remove extra whitespace
    text = ' '.join(text.split())
    return text


def find_best_cost_match(item_name, cost_map, normalized_map, sku_cost_map=None):
    """Find the best matching cost for an item name using multiple strategies."""
    if not item_name:
        return 0
    
    # Clean name from potential Salla prefix like "(SKU: )" or "(SKU: 123)"
    name = str(item_name).strip()
    if ')' in name and (name.startswith('(SKU:') or name.startswith('(')):
        name = name.split(')', 1)[-1].strip()

    # Strategy 1: Exact match (original)
    if name in cost_map:
        return cost_map[name]
    
    # Strategy 2: Exact match (normalized)
    normalized_name = normalize_text(name)
    if normalized_name in normalized_map:
        return normalized_map[normalized_name]
    
    # Strategy 3: Partial match
    for prod_name, cost in cost_map.items():
        if not prod_name: continue
        norm_prod = normalize_text(prod_name)
        if normalized_name and norm_prod and (normalized_name in norm_prod or norm_prod in normalized_name):
            return cost
            
    # Strategy 4: SKU match if SKU exists in name or SKU map is provided
    if sku_cost_map:
        sku_match = re.search(r'\[?([A-Za-z0-9\-_]{4,})\]?', name)
        if sku_match:
            potential_sku = sku_match.group(1).upper()
            if potential_sku in sku_cost_map:
                return sku_cost_map[potential_sku]
        # Also try direct SKU match if item_name is just a SKU
        if name.upper() in sku_cost_map:
            return sku_cost_map[name.upper()]
    
    return 0


def recalculate_snapshot_costs(snapshot_id):
    """Re-run cost calculation for all orders in a snapshot based on updated product_costs.
    
    Enhanced with fuzzy matching for Arabic/English product names.
    Includes robustness against database locks and errors.
    """
    conn = None
    try:
        conn = get_db()
        
        # Fetch cost map (original names)
        rows = conn.execute("SELECT product_name, cost FROM product_costs").fetchall()
        cost_map = {r[0].strip(): r[1] for r in rows if r[0]}
        
        # Create normalized map for fuzzy matching
        normalized_map = {normalize_text(name): cost for name, cost in cost_map.items() if name}
        
        # Also fetch SKU-based costs
        sku_rows = conn.execute("SELECT sku, cost FROM product_costs WHERE sku IS NOT NULL AND sku != ''").fetchall()
        sku_cost_map = {r[0].strip(): r[1] for r in sku_rows if r[0]}
        
        # Fetch orders
        cur = conn.cursor()
        cur.execute("SELECT order_id, items_summary FROM orders WHERE snapshot_id=?", (snapshot_id,))
        orders = cur.fetchall()
        
        updates = []
        matched_count = 0
        total_items_checked = 0
        
        for oid, summary in orders:
            if not summary:
                continue
            
            # Consistent splitting with generate_report.py
            parts = summary.split('|')
            total_c = 0.0
            found = False
            
            for part in parts:
                part = part.strip()
                if not part:
                    continue
                
                # Extract item name and quantity (e.g., "Product Name x2")
                qty = 1
                name = part
                if ' x' in part:
                    try:
                        name_part, qty_part = part.rsplit(' x', 1)
                        name = name_part.strip()
                        qty = int(qty_part)
                    except: pass
                
                total_items_checked += 1
                
                # Use the improved find_best_cost_match
                item_cost = find_best_cost_match(name, cost_map, normalized_map, sku_cost_map)
                
                if item_cost > 0:
                    total_c += item_cost * qty
                    found = True
                    matched_count += 1
            
            # Only update if we found at least one matching item
            if found:
                updates.append((total_c, oid, snapshot_id))
        
        if updates:
            cur.executemany("UPDATE orders SET cost=? WHERE order_id=? AND snapshot_id=?", updates)
            conn.commit()
            print(f"Recalculated costs for {len(updates)} orders. Matched {matched_count}/{total_items_checked} items.")
        else:
            print(f"No cost matches found. Checked {total_items_checked} items across {len(orders)} orders.")
            
    except Exception as e:
        print(f"[CRITICAL] Error in recalculate_snapshot_costs: {e}")
    finally:
        if conn:
            conn.close()

@app.route("/api/returns/add", methods=["POST"])
def api_returns_add():
    data = request.json
    tracking_id = data.get("tracking_id", "").strip()
    notes = data.get("notes", "").strip()

    if not tracking_id:
        return jsonify({"success": False, "error": "رقم التتبع مطلوب"}), 400

    conn = None
    try:
        conn = get_db()
        # Insert or ignore (if duplicate, it won't crash but won't insert)
        cursor = conn.cursor()
        cursor.execute("SELECT 1 FROM physical_returns WHERE tracking_id=?", (tracking_id,))
        exists = cursor.fetchone()
        
        if exists:
             return jsonify({"success": False, "error": "تم مسح هذا الرقم مسبقاً", "duplicate": True}), 200

        cursor.execute(
            "INSERT INTO physical_returns (tracking_id, notes) VALUES (?, ?)",
            (tracking_id, notes)
        )
        conn.commit()
        return jsonify({"success": True, "message": "تم تسجيل الاستلام بنجاح"}), 200
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/returns/list", methods=["GET"])
def api_returns_list():
    limit = request.args.get("limit", 50, type=int)
    conn = None
    try:
        conn = get_db()
        rows = conn.execute(
            "SELECT id, tracking_id, scanned_at, notes FROM physical_returns ORDER BY scanned_at DESC LIMIT ?", 
            (limit,)
        ).fetchall()
        
        # Format the output
        records = [{
            "id": r["id"],
            "tracking_id": r["tracking_id"],
            "scanned_at": r["scanned_at"],
            "notes": r["notes"]
        } for r in rows]

        # Get total count
        total = conn.execute("SELECT COUNT(*) FROM physical_returns").fetchone()[0]

        return jsonify({"success": True, "data": records, "total": total}), 200
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/returns/delete/<int:item_id>", methods=["DELETE"])
def api_returns_delete(item_id):
    conn = None
    try:
        conn = get_db()
        conn.execute("DELETE FROM physical_returns WHERE id=?", (item_id,))
        conn.commit()
        return jsonify({"success": True}), 200
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/update-cost", methods=["POST"])
def update_cost():
    data = request.json
    sku  = data.get("sku")
    name = data.get("name")
    try:
        cost = float(data.get("cost"))
    except:
        return jsonify({"success": False, "error": "Invalid cost"}), 400
        
    conn = None
    try:
        conn = get_db()
        # Upsert logic
        exists = conn.execute("SELECT 1 FROM product_costs WHERE sku=?", (sku,)).fetchone()
        if exists:
            conn.execute("UPDATE product_costs SET cost = ?, updated_at=CURRENT_TIMESTAMP WHERE sku = ?", (cost, sku))
        else:
            if not name:
                return jsonify({"success": False, "error": "Product name required for new items"}), 400
            conn.execute("INSERT OR REPLACE INTO product_costs (sku, product_name, cost, updated_at) VALUES (?, ?, ?, CURRENT_TIMESTAMP)", (sku, name, cost))
            
        conn.commit()
        return jsonify({"success": True})
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500
    finally:
        if conn:
            conn.close()

@app.route("/api/save-bulk-costs", methods=["POST"])
def api_save_bulk_costs():
    """Save multiple product costs at once and start background recalculation."""
    data = request.json
    items = data.get("items", [])
    if not items:
        return jsonify({"success": True, "message": "No items to save"})

    conn = None
    try:
        conn = get_db()
        for item in items:
            sku = item.get("sku")
            name = item.get("name")
            cost = float(item.get("cost", 0))
            conn.execute("""
                INSERT INTO product_costs (sku, product_name, cost, updated_at) 
                VALUES (?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(sku) DO UPDATE SET cost=excluded.cost, updated_at=excluded.updated_at
            """, (sku, name, cost))
        conn.commit()
    except Exception as e:
        if conn: conn.close()
        return jsonify({"success": False, "error": f"Failed to save items: {e}"}), 500
    finally:
        if conn: conn.close()

    # Recalculate in background to prevent timeout
    sid = get_active_snapshot_id()
    if sid:
        def run_recalc():
            try:
                print(f"Background: Starting cost recalculation for snapshot {sid}...")
                recalculate_snapshot_costs(sid)
                print(f"Background: Cost recalculation complete for snapshot {sid}.")
            except Exception as e:
                print(f"Background Error: Recalculation failed: {e}")
        
        thread = threading.Thread(target=run_recalc)
        thread.daemon = True
        thread.start()

    return jsonify({
        "success": True, 
        "count": len(items), 
        "message": "تم الحفظ بنجاح، جاري تحديث الأرباح في الخلفية..."
    })

@app.route("/api/upload-costs", methods=["POST"])
def api_upload_costs():
    if "files" not in request.files and "file" not in request.files:
        return jsonify({"success": False, "error": "لم يتم اختيار ملف"}), 400
    
    files = request.files.getlist("files") or request.files.getlist("file")
    if not files or all(f.filename == '' for f in files):
        return jsonify({"success": False, "error": "اسم الملف فارغ"}), 400
    
    total_count = 0
    try:
        for file in files:
            if not file.filename: continue
            
            # Save temp
            temp_path = os.path.join(UPLOAD_FOLDER, f"temp_costs_{secure_filename(file.filename)}")
            file.save(temp_path)
            
            try:
                df = process_data.read_file_safe(temp_path)
                if df is not None:
                    count = process_data.process_costs_file(df)
                    total_count += count
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        
        # Trigger Recalculation
        sid = get_active_snapshot_id()
        if sid:
            recalculate_snapshot_costs(sid)
            
        return jsonify({
            "success": True,
            "message": f"تم تحديث تكاليف {total_count} منتج بنجاح ✅"
        })
    except Exception as e:
        return jsonify({"success": False, "error": str(e)}), 500

@app.route("/")
def index():
    accounts = get_all_accounts()
    return render_template("index.html", accounts=accounts)


@app.route("/api/stats")
def api_stats():
    sid = request.args.get("snapshot_id", None, type=int)
    if sid is None:
        sid = get_active_snapshot_id()
    return jsonify({
        "stats":     db_stats(sid),
        "platforms": platform_breakdown(sid),
    })


@app.route("/api/files")
def api_files():
    return jsonify({"files": list_uploaded_files()})


@app.route("/api/reports")
def api_reports():
    return jsonify({"reports": list_reports()})


@app.route("/api/charts")
def api_charts():
    try:
        conn = get_db()
        cur  = conn.cursor()

        # 1. Platform breakdown (orders count + expected + collected)
        sid = request.args.get("snapshot_id", None, type=int)
        if sid is None:
            sid = get_active_snapshot_id()
        cur.execute("""
            SELECT o.platform,
                   COUNT(DISTINCT o.order_id)            AS orders,
                   COALESCE(SUM(o.price),0)              AS expected,
                   COALESCE(SUM(c.collected_amount),0)   AS collected
            FROM orders o
            LEFT JOIN collections c
                ON o.order_id = c.order_id AND c.snapshot_id = ?
            WHERE o.snapshot_id = ?
            GROUP BY o.platform ORDER BY orders DESC
        """, (sid, sid))
        platforms_raw = [dict(r) for r in cur.fetchall()]

        # 2. Payment status distribution
        cur.execute("""
            SELECT o.order_id,
                   o.price,
                   COALESCE(SUM(c.collected_amount),0) AS collected
            FROM orders o
            LEFT JOIN collections c
                ON o.order_id = c.order_id AND c.snapshot_id = ?
            WHERE o.snapshot_id = ?
            GROUP BY o.order_id
        """, (sid, sid))
        status_counts = {"مدفوع": 0, "غير مدفوع": 0, "زيادة": 0}
        for row in cur.fetchall():
            exp, col = row["price"], row["collected"]
            if col == 0:
                status_counts["غير مدفوع"] += 1
            elif col > exp + 0.1:
                status_counts["زيادة"] += 1
            else:
                status_counts["مدفوع"] += 1

        # 3. Weekly trend (orders per week)
        cur.execute("""
            SELECT week_number, COUNT(*) AS cnt, COALESCE(SUM(price),0) AS total
            FROM orders WHERE week_number IS NOT NULL AND snapshot_id = ?
            GROUP BY week_number ORDER BY week_number
        """, (sid,))
        weekly = [dict(r) for r in cur.fetchall()]

        conn.close()

        return jsonify({
            "platforms": {
                "labels":    [p["platform"] for p in platforms_raw],
                "orders":    [p["orders"]   for p in platforms_raw],
                "expected":  [round(p["expected"],  2) for p in platforms_raw],
                "collected": [round(p["collected"], 2) for p in platforms_raw],
            },
            "status": {
                "labels": list(status_counts.keys()),
                "values": list(status_counts.values()),
            },
            "weekly": {
                "labels": [f"أسبوع {w['week_number']}" for w in weekly],
                "orders": [w["cnt"]   for w in weekly],
                "totals": [round(w["total"], 2) for w in weekly],
            }
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/upload", methods=["POST"])
def upload():
    if "files" not in request.files:
        return jsonify({"success": False, "message": "لم يتم إرسال أي ملف"}), 400

    uploaded = []
    errors   = []
    
    # Get selected account from form data (sent by Dropdown)
    account_name = request.form.get("account_name", "").strip()

    for file in request.files.getlist("files"):
        if not file.filename:
            continue
        
        filename = file.filename
        name, ext = os.path.splitext(filename)
        ext = ext.lower()
        
        if ext not in ALLOWED_EXT:
            errors.append(f"{filename}: نوع الملف غير مدعوم")
            continue
            
        # Append account name if provided
        final_filename = filename
        if account_name and account_name != 'Auto Detect':
            # Clean account name
            safe_acc = "".join(c for c in account_name if c.isalnum() or c in (' ', '_', '-')).strip()
            final_filename = f"{name}_[{safe_acc}]{ext}"
        
        dest = os.path.join(UPLOAD_FOLDER, final_filename)
        try:
            file.save(dest)
            uploaded.append(final_filename)
        except PermissionError:
            errors.append(f"{filename}: الملف مفتوح في برنامج آخر، يرجى إغلاقه أولاً")
        except Exception as e:
            errors.append(f"{filename}: خطأ غير متوقع ({e})")

    if uploaded and not errors:
        return jsonify({"success": True,
                        "message": f"تم رفع {len(uploaded)} ملف بنجاح",
                        "files": uploaded, "errors": errors})
    elif uploaded and errors:
        return jsonify({"success": True, 
                        "message": f"تم رفع {len(uploaded)} ملف، ولكن فشل {len(errors)} ملف",
                        "files": uploaded, "errors": errors})
    return jsonify({"success": False, "message": "فشل رفع الملفات", "errors": errors}), 400


@app.route("/delete-file", methods=["POST"])
def delete_file():
    data     = request.get_json()
    filename = data.get("filename", "")
    path     = os.path.join(UPLOAD_FOLDER, filename)
    if os.path.exists(path):
        os.remove(path)
        return jsonify({"success": True, "message": f"تم حذف {filename}"})
    return jsonify({"success": False, "message": "الملف غير موجود"}), 404


@app.route("/process", methods=["POST"])
def process():
    """Run the full pipeline: process_data → process_collections → generate_report"""
    import io, contextlib

    # ── Determine / create snapshot ───────────────────────────
    data_json   = request.get_json(silent=True) or {}
    week_label  = data_json.get("label", "").strip()
    if not week_label:
        now        = datetime.now()
        week_label = f"أسبوع {now.isocalendar()[1]} - {now.year}"

    conn = get_db()
    cur  = conn.cursor()
    cur.execute(
        "INSERT INTO weekly_snapshots (label, week_number, year) VALUES (?, ?, ?)",
        (week_label, datetime.now().isocalendar()[1], datetime.now().year)
    )
    snapshot_id = cur.lastrowid
    conn.commit()
    conn.close()

    log_lines = []

    def capture(fn, *args, **kwargs):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            try:
                fn(*args, **kwargs)
            except Exception as e:
                buf.write(f"\n[ERROR] {e}\n")
        return buf.getvalue()

    log_lines.append("═" * 50)
    log_lines.append(f"▶ Snapshot ID: {snapshot_id} | {week_label}")
    log_lines.append("═" * 50)

    log_lines.append("▶ الخطوة 1/3: معالجة ملفات الطلبات...")
    log_lines.append("═" * 50)
    log_lines.append(capture(process_data.main, snapshot_id=snapshot_id))

    log_lines.append("═" * 50)
    log_lines.append("▶ الخطوة 2/3: معالجة ملفات التحصيل...")
    log_lines.append("═" * 50)
    log_lines.append(capture(process_collections.process_collections, snapshot_id=snapshot_id))
    
    # ── Check for Unmatched Collections ──
    try:
        conn = get_db()
        unmatched_count = conn.execute("""
            SELECT COUNT(DISTINCT c.order_id)
            FROM collections c
            LEFT JOIN orders o ON c.order_id = o.order_id AND c.snapshot_id = o.snapshot_id
            WHERE c.snapshot_id = ? AND o.order_id IS NULL
        """, (snapshot_id,)).fetchone()[0]
        conn.close()
        if unmatched_count > 0:
            log_lines.append("⚠️ [تنبيه هام]:")
            log_lines.append(f"تم العثور على ({unmatched_count}) عمليات تحصيل لا يوجد لها طلبات مطابقة.")
            log_lines.append("ربما تكون هذه الطلبات قديمة، أو لم يتم رفع ملف الطلبات الخاص بها.")
    except Exception as e:
        log_lines.append(f"[ERROR] أخفق فحص التطابق: {e}")

    # ── Step 2.5: Apply saved product costs to orders BEFORE generating report ──
    log_lines.append("═" * 50)
    log_lines.append("▶ الخطوة 2.5/3: تطبيق التكاليف المحفوظة على الطلبات...")
    log_lines.append(capture(recalculate_snapshot_costs, snapshot_id))

    log_lines.append("═" * 50)
    log_lines.append("▶ الخطوة 3/3: إنشاء التقرير النهائي...")
    log_lines.append("═" * 50)
    log_lines.append(capture(generate_report.generate_weekly_report,
                             snapshot_id=snapshot_id, label=week_label))

    log_lines.append("═" * 50)
    log_lines.append("✅ اكتملت جميع الخطوات بنجاح!")

    return jsonify({
        "success":     True,
        "log":         "\n".join(log_lines),
        "snapshot_id": snapshot_id,
        "stats":       db_stats(snapshot_id),
        "platforms":   platform_breakdown(snapshot_id),
        "reports":     list_reports(),
    })


@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(REPORTS_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "الملف غير موجود", 404


@app.route("/regenerate-report", methods=["POST"])
def regenerate_report():
    """Re-apply costs to current snapshot and regenerate the Excel report.
    This is called AFTER the user enters product costs via the Products page.
    Does NOT re-process files or create a new snapshot.
    """
    import io, contextlib

    sid = get_active_snapshot_id()
    if not sid:
        return jsonify({"success": False, "message": "لا يوجد snapshot نشط. يرجى معالجة الملفات أولاً."}), 400

    log_lines = []

    def capture(fn, *args, **kwargs):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            try:
                fn(*args, **kwargs)
            except Exception as e:
                buf.write(f"\n[ERROR] {e}\n")
        return buf.getvalue()

    log_lines.append("═" * 50)
    log_lines.append(f"▶ إعادة توليد التقرير للـ Snapshot ID: {sid}")
    log_lines.append("═" * 50)

    # Step 1: Re-apply costs from product_costs → orders.cost
    log_lines.append("▶ الخطوة 1/2: إعادة حساب التكاليف على الطلبات...")
    log_lines.append(capture(recalculate_snapshot_costs, sid))

    # Step 2: Re-generate report (will use updated orders.cost from DB)
    log_lines.append("═" * 50)
    log_lines.append("▶ الخطوة 2/2: إعادة توليد التقرير النهائي...")
    
    # Get label from existing snapshot
    try:
        conn = get_db()
        snap_row = conn.execute("SELECT label FROM weekly_snapshots WHERE snapshot_id=?", (sid,)).fetchone()
        conn.close()
        week_label = snap_row[0] if snap_row else f"snapshot_{sid}"
    except:
        week_label = f"snapshot_{sid}"

    log_lines.append(capture(generate_report.generate_weekly_report,
                             snapshot_id=sid, label=week_label))

    log_lines.append("═" * 50)
    log_lines.append("✅ تم إعادة توليد التقرير بنجاح مع التكاليف المحدّثة!")

    return jsonify({
        "success":  True,
        "log":      "\n".join(log_lines),
        "snapshot_id": sid,
        "reports": list_reports(),
    })


@app.route("/reset-db", methods=["POST"])
def reset_db():
    """Reset ONLY the active (latest) snapshot — does NOT delete historical data."""
    try:
        sid  = get_active_snapshot_id()
        conn = get_db()
        conn.execute("DELETE FROM orders      WHERE snapshot_id=?", (sid,))
        conn.execute("DELETE FROM collections WHERE snapshot_id=?", (sid,))
        conn.execute("DELETE FROM weekly_reports WHERE snapshot_id=?", (sid,))
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": f"تم إعادة تعيين البيانات للـ snapshot الحالي (id={sid})"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/new-week", methods=["POST"])
def new_week():
    """
    Archive current week and start fresh:
    1. Create a new snapshot entry in weekly_snapshots
    2. Delete uploaded sample files (data already saved in DB under old snapshot)
    NOTE: Historical data is NEVER deleted — it lives under its snapshot_id.
    """
    try:
        data_json  = request.get_json(silent=True) or {}
        week_label = data_json.get("label", "").strip()
        if not week_label:
            now        = datetime.now()
            week_label = f"أسبوع {now.isocalendar()[1]} - {now.year}"

        # Create new snapshot for the upcoming week
        conn = get_db()
        cur  = conn.cursor()
        cur.execute(
            "INSERT INTO weekly_snapshots (label, week_number, year, notes) VALUES (?, ?, ?, ?)",
            (week_label,
             datetime.now().isocalendar()[1],
             datetime.now().year,
             "بدأ تلقائياً عند الضغط على زر أسبوع جديد")
        )
        new_snapshot_id = cur.lastrowid
        conn.commit()
        conn.close()

        # Delete sample files (data is safely stored in DB)
        deleted_files = []
        for f in os.listdir(UPLOAD_FOLDER):
            if f.startswith("~$"): continue
            ext = os.path.splitext(f)[1].lower()
            if ext in ALLOWED_EXT:
                os.remove(os.path.join(UPLOAD_FOLDER, f))
                deleted_files.append(f)

        return jsonify({
            "success":         True,
            "message":         f"✅ تم بدء أسبوع جديد! البيانات السابقة محفوظة. حُذف {len(deleted_files)} ملف.",
            "new_snapshot_id": new_snapshot_id,
            "label":           week_label,
            "deleted_files":   deleted_files,
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


# ─────────────────────────────────────────────────────────────────────────────
# Snapshots History API
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/api/snapshots")
def api_snapshots():
    """Return list of all weekly snapshots with their KPIs."""
    try:
        conn = get_db()
        rows = conn.execute("""
            SELECT
                s.snapshot_id,
                s.label,
                s.week_number,
                s.year,
                s.created_at,
                COALESCE(r.total_orders,    0) AS total_orders,
                COALESCE(r.total_sales,     0) AS total_sales,
                COALESCE(r.total_collected, 0) AS total_collected,
                COALESCE(r.net_profit,      0) AS net_profit,
                COALESCE(r.collection_rate, 0) AS collection_rate,
                COALESCE(r.report_path,    '') AS report_path
            FROM weekly_snapshots s
            LEFT JOIN weekly_reports r ON s.snapshot_id = r.snapshot_id
            ORDER BY s.snapshot_id DESC
        """).fetchall()
        conn.close()
        return jsonify({"snapshots": [dict(r) for r in rows]})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/snapshot/<int:sid>")
def api_snapshot_detail(sid):
    """Return detailed stats + platform breakdown for a specific snapshot."""
    return jsonify({
        "stats":     db_stats(sid),
        "platforms": platform_breakdown(sid),
    })


# ─────────────────────────────────────────────────────────────────────────────
# Accounts (Stores / Branches) Management API
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/api/accounts")
def api_accounts():
    """Return all configured accounts."""
    return jsonify({"accounts": get_all_accounts()})


@app.route("/api/accounts/add", methods=["POST"])
def api_add_account():
    """Add a new store/branch account."""
    data = request.get_json(force=True)
    platform  = (data.get("platform_name") or "").strip()
    account   = (data.get("account_name") or "").strip()
    country   = (data.get("country") or "SA").strip()
    fixed_shipping = float(data.get("fixed_shipping_cost", 0) or 0)
    cost_inc_tax   = 1 if data.get("cost_includes_tax") else 0
    pm_commission  = float(data.get("payment_commission_rate", 0) or 0)
    tax_rate       = float(data.get("tax_rate", 0) or 0)
    client_shipping = float(data.get("client_shipping_cost", 0) or 0)

    if not platform or not account:
        return jsonify({"success": False, "message": "يجب إدخال اسم المنصة واسم الحساب"})

    try:
        conn = get_db()
        conn.execute(
            """INSERT INTO accounts 
               (platform_name, account_name, country, fixed_shipping_cost, cost_includes_tax, payment_commission_rate, tax_rate, client_shipping_cost) 
               VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
            (platform, account, country, fixed_shipping, cost_inc_tax, pm_commission, tax_rate, client_shipping),
        )
        # Also ensure the platform exists in the platforms table
        conn.execute(
            "INSERT OR IGNORE INTO platforms (platform_name, commission_rate, tax_rate, shipping_default) VALUES (?, 0, 0, 0)",
            (platform,),
        )
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": f"تمت إضافة الحساب: {account} ({platform})"})
    except Exception as e:
        if "UNIQUE constraint" in str(e):
            return jsonify({"success": False, "message": "هذا الحساب موجود بالفعل لنفس المنصة"})
        return jsonify({"success": False, "message": f"خطأ: {e}"})


@app.route("/api/accounts/update", methods=["POST"])
def api_update_account():
    """Update an existing account."""
    data = request.get_json(force=True)
    acc_id    = data.get("id")
    platform  = (data.get("platform_name") or "").strip()
    account   = (data.get("account_name") or "").strip()
    country   = (data.get("country") or "SA").strip()
    fixed_shipping = float(data.get("fixed_shipping_cost", 0) or 0)
    cost_inc_tax   = 1 if data.get("cost_includes_tax") else 0
    pm_commission  = float(data.get("payment_commission_rate", 0) or 0)
    tax_rate       = float(data.get("tax_rate", 0) or 0)
    client_shipping = float(data.get("client_shipping_cost", 0) or 0)

    if not acc_id or not platform or not account:
        return jsonify({"success": False, "message": "بيانات غير مكتملة"})

    try:
        conn = get_db()
        conn.execute(
            """UPDATE accounts SET 
               platform_name=?, account_name=?, country=?, 
               fixed_shipping_cost=?, cost_includes_tax=?,
               payment_commission_rate=?, tax_rate=?, client_shipping_cost=? 
               WHERE id=?""",
            (platform, account, country, fixed_shipping, cost_inc_tax, pm_commission, tax_rate, client_shipping, acc_id),
        )
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": f"تم تحديث الحساب بنجاح"})
    except Exception as e:
        if "UNIQUE constraint" in str(e):
            return jsonify({"success": False, "message": "يوجد حساب آخر بنفس الاسم في هذه المنصة"})
        return jsonify({"success": False, "message": f"خطأ: {e}"})

@app.route("/api/accounts/delete", methods=["POST"])
def api_delete_account():
    """Delete an account."""
    data   = request.get_json(force=True)
    acc_id = data.get("id")

    if not acc_id:
        return jsonify({"success": False, "message": "لم يتم تحديد الحساب"})

    try:
        conn = get_db()
        conn.execute("DELETE FROM accounts WHERE id=?", (acc_id,))
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": "تم حذف الحساب بنجاح"})
    except Exception as e:
        return jsonify({"success": False, "message": f"خطأ: {e}"})


@app.route("/api/platforms/list")
def api_platforms_list():
    """Return distinct platform names for dropdown."""
    try:
        conn = get_db()
        rows = conn.execute("SELECT DISTINCT platform_name FROM platforms ORDER BY platform_name").fetchall()
        conn.close()
        return jsonify({"platforms": [r["platform_name"] for r in rows]})
    except:
        return jsonify({"platforms": []})


if __name__ == "__main__":
    import webbrowser, threading
    def open_browser():
        import time; time.sleep(1.2)
        webbrowser.open("http://127.0.0.1:5000")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(debug=False, port=5000)
