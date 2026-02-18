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
import os, sys, sqlite3, json, shutil
from datetime import datetime
from werkzeug.utils import secure_filename

# ── Add project root to path so we can import our scripts ──────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

import process_data
import process_collections
import generate_report

app = Flask(__name__)
app.secret_key = "weekly_reconciliation_2026"

UPLOAD_FOLDER  = os.path.join(BASE_DIR, "samples")
REPORTS_FOLDER = os.path.join(BASE_DIR, "reports")
DB_PATH        = os.path.join(BASE_DIR, "finance_system.db")
ALLOWED_EXT    = {".csv", ".xlsx", ".xls"}

os.makedirs(UPLOAD_FOLDER,  exist_ok=True)
os.makedirs(REPORTS_FOLDER, exist_ok=True)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def db_stats():
    """Return quick counts from the database."""
    try:
        conn = get_db()
        cur  = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM orders")
        orders = cur.fetchone()[0]
        cur.execute("SELECT COUNT(*) FROM collections")
        collections = cur.fetchone()[0]
        cur.execute("SELECT COALESCE(SUM(price),0) FROM orders")
        total_expected = cur.fetchone()[0]
        cur.execute("SELECT COALESCE(SUM(collected_amount),0) FROM collections")
        total_collected = cur.fetchone()[0]
        conn.close()
        rate = round(total_collected / total_expected * 100, 1) if total_expected else 0
        return {
            "orders": orders,
            "collections": collections,
            "total_expected": round(total_expected, 2),
            "total_collected": round(total_collected, 2),
            "collection_rate": rate,
        }
    except Exception as e:
        return {"orders": 0, "collections": 0, "total_expected": 0,
                "total_collected": 0, "collection_rate": 0, "error": str(e)}


def platform_breakdown():
    try:
        conn = get_db()
        cur  = conn.cursor()
        cur.execute("""
            SELECT o.platform,
                   COUNT(o.order_id)                        AS orders,
                   COALESCE(SUM(o.price),0)                 AS expected,
                   COALESCE(SUM(c.collected_amount),0)      AS collected
            FROM orders o
            LEFT JOIN collections c ON o.order_id = c.order_id
            GROUP BY o.platform
        """)
        rows = [dict(r) for r in cur.fetchall()]
        conn.close()
        for r in rows:
            r["rate"] = round(r["collected"] / r["expected"] * 100, 1) if r["expected"] else 0
            r["expected"]  = round(r["expected"],  2)
            r["collected"] = round(r["collected"], 2)
        return rows
    except:
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

@app.route("/")
def index():
    return render_template("index.html",
                           stats=db_stats(),
                           platforms=platform_breakdown(),
                           files=list_uploaded_files(),
                           reports=list_reports())


@app.route("/api/stats")
def api_stats():
    return jsonify({
        "stats": db_stats(),
        "platforms": platform_breakdown(),
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
        cur.execute("""
            SELECT o.platform,
                   COUNT(DISTINCT o.order_id)            AS orders,
                   COALESCE(SUM(o.price),0)              AS expected,
                   COALESCE(SUM(c.collected_amount),0)   AS collected
            FROM orders o
            LEFT JOIN collections c ON o.order_id = c.order_id
            GROUP BY o.platform ORDER BY orders DESC
        """)
        platforms_raw = [dict(r) for r in cur.fetchall()]

        # 2. Payment status distribution
        cur.execute("""
            SELECT o.order_id,
                   o.price,
                   COALESCE(SUM(c.collected_amount),0) AS collected
            FROM orders o
            LEFT JOIN collections c ON o.order_id = c.order_id
            GROUP BY o.order_id
        """)
        status_counts = {"مدفوع": 0, "غير مدفوع": 0, "مدفوع جزئياً": 0, "زيادة": 0}
        for row in cur.fetchall():
            exp, col = row["price"], row["collected"]
            diff = exp - col
            if col == 0:
                status_counts["غير مدفوع"] += 1
            elif abs(diff) < 0.1:
                status_counts["مدفوع"] += 1
            elif col > exp:
                status_counts["زيادة"] += 1
            else:
                status_counts["مدفوع جزئياً"] += 1

        # 3. Weekly trend (orders per week)
        cur.execute("""
            SELECT week_number, COUNT(*) AS cnt, COALESCE(SUM(price),0) AS total
            FROM orders WHERE week_number IS NOT NULL
            GROUP BY week_number ORDER BY week_number
        """)
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
    for file in request.files.getlist("files"):
        if not file.filename:
            continue
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in ALLOWED_EXT:
            errors.append(f"{file.filename}: نوع الملف غير مدعوم")
            continue
        # Keep original Arabic filename (safe for local use)
        dest = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(dest)
        uploaded.append(file.filename)

    if uploaded:
        return jsonify({"success": True,
                        "message": f"تم رفع {len(uploaded)} ملف بنجاح",
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
    log_lines = []

    def capture(fn):
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            try:
                fn()
            except Exception as e:
                buf.write(f"\n[ERROR] {e}\n")
        return buf.getvalue()

    log_lines.append("═" * 50)
    log_lines.append("▶ الخطوة 1/3: معالجة ملفات الطلبات...")
    log_lines.append("═" * 50)
    log_lines.append(capture(process_data.main))

    log_lines.append("═" * 50)
    log_lines.append("▶ الخطوة 2/3: معالجة ملفات التحصيل...")
    log_lines.append("═" * 50)
    log_lines.append(capture(process_collections.process_collections))

    log_lines.append("═" * 50)
    log_lines.append("▶ الخطوة 3/3: إنشاء التقرير النهائي...")
    log_lines.append("═" * 50)
    log_lines.append(capture(generate_report.generate_weekly_report))

    log_lines.append("═" * 50)
    log_lines.append("✅ اكتملت جميع الخطوات بنجاح!")

    return jsonify({
        "success": True,
        "log": "\n".join(log_lines),
        "stats": db_stats(),
        "platforms": platform_breakdown(),
        "reports": list_reports(),
    })


@app.route("/download/<filename>")
def download(filename):
    path = os.path.join(REPORTS_FOLDER, filename)
    if os.path.exists(path):
        return send_file(path, as_attachment=True)
    return "الملف غير موجود", 404


@app.route("/reset-db", methods=["POST"])
def reset_db():
    try:
        conn = get_db()
        conn.execute("DELETE FROM orders")
        conn.execute("DELETE FROM collections")
        conn.commit()
        conn.close()
        return jsonify({"success": True, "message": "تم إعادة تعيين قاعدة البيانات"})
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


@app.route("/new-week", methods=["POST"])
def new_week():
    """
    Start a fresh week:
    1. Clear orders + collections from DB
    2. Delete all uploaded sample files
    """
    try:
        # 1. Clear database
        conn = get_db()
        conn.execute("DELETE FROM orders")
        conn.execute("DELETE FROM collections")
        conn.commit()
        conn.close()

        # 2. Delete all sample files
        deleted_files = []
        for f in os.listdir(UPLOAD_FOLDER):
            if f.startswith("~$"):
                continue
            ext = os.path.splitext(f)[1].lower()
            if ext in ALLOWED_EXT:
                os.remove(os.path.join(UPLOAD_FOLDER, f))
                deleted_files.append(f)

        return jsonify({
            "success": True,
            "message": f"✅ تم بدء أسبوع جديد! حُذف {len(deleted_files)} ملف وتم تفريغ قاعدة البيانات.",
            "deleted_files": deleted_files,
        })
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 500


if __name__ == "__main__":
    import webbrowser, threading
    def open_browser():
        import time; time.sleep(1.2)
        webbrowser.open("http://127.0.0.1:5000")
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(debug=False, port=5000)
