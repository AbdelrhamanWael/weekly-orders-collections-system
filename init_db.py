# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Module   : init_db.py — Database Initialization
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
#  Updated   : February 2026 — Added Weekly Snapshot support
# ============================================================
import sqlite3
import os
from datetime import datetime

DB_NAME = "finance_system.db"

def create_database():
    if os.path.exists(DB_NAME):
        print(f"Database {DB_NAME} already exists. Running migrations...")

    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()

    # ─────────────────────────────────────────────────────────
    # 1. Weekly Snapshots Table  (NEW)
    #    Every time "new week" is triggered, a snapshot is saved
    #    so historical data is never lost.
    # ─────────────────────────────────────────────────────────
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS weekly_snapshots (
        snapshot_id   INTEGER PRIMARY KEY AUTOINCREMENT,
        label         TEXT NOT NULL,          -- e.g. "أسبوع 7 - 2026"
        week_number   INTEGER,
        year          INTEGER,
        created_at    TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        notes         TEXT DEFAULT ''
    )
    ''')

    # ─────────────────────────────────────────────────────────
    # 2. Orders Table  (snapshot_id column added)
    # ─────────────────────────────────────────────────────────
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        order_id     TEXT NOT NULL,
        snapshot_id  INTEGER NOT NULL DEFAULT 0,
        platform     TEXT NOT NULL,
        account_name TEXT DEFAULT '',
        order_date   DATE,
        price        REAL DEFAULT 0,
        cost         REAL DEFAULT 0,
        shipping     REAL DEFAULT 0,
        commission   REAL DEFAULT 0,
        tax          REAL DEFAULT 0,
        items_summary TEXT DEFAULT '',
        week_number  INTEGER,
        year         INTEGER,
        upload_date  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (order_id, snapshot_id),
        FOREIGN KEY (snapshot_id) REFERENCES weekly_snapshots(snapshot_id)
    )
    ''')

    # ─────────────────────────────────────────────────────────
    # 3. Collections Table  (snapshot_id column added)
    # ─────────────────────────────────────────────────────────
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS collections (
        collection_id    INTEGER PRIMARY KEY AUTOINCREMENT,
        snapshot_id      INTEGER NOT NULL DEFAULT 0,
        order_id         TEXT,
        original_amount  REAL DEFAULT 0,
        collection_fee   REAL DEFAULT 0,
        collected_amount REAL,
        collection_date  DATE,
        is_return        INTEGER DEFAULT 0,
        account_name     TEXT DEFAULT '',
        week_number      INTEGER,
        year             INTEGER,
        upload_date      TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (snapshot_id) REFERENCES weekly_snapshots(snapshot_id)
    )
    ''')

    # ─────────────────────────────────────────────────────────
    # 4. Platforms Table
    # ─────────────────────────────────────────────────────────
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS platforms (
        platform_id       INTEGER PRIMARY KEY AUTOINCREMENT,
        platform_name     TEXT UNIQUE,
        commission_rate   REAL DEFAULT 0,
        tax_rate          REAL DEFAULT 0,
        shipping_default  REAL DEFAULT 0
    )
    ''')

    # ─────────────────────────────────────────────────────────
    # 5. Weekly Reports Table  (NOW ACTUALLY USED)
    #    Stores aggregated KPIs per snapshot for fast retrieval
    # ─────────────────────────────────────────────────────────
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS weekly_reports (
        report_id         INTEGER PRIMARY KEY AUTOINCREMENT,
        snapshot_id       INTEGER UNIQUE,
        week_number       INTEGER,
        year              INTEGER,
        label             TEXT,
        total_orders      INTEGER DEFAULT 0,
        total_sales       REAL    DEFAULT 0,
        total_collected   REAL    DEFAULT 0,
        total_uncollected REAL    DEFAULT 0,
        net_profit        REAL    DEFAULT 0,
        collection_rate   REAL    DEFAULT 0,
        paid_count        INTEGER DEFAULT 0,
        unpaid_count      INTEGER DEFAULT 0,
        partial_count     INTEGER DEFAULT 0,
        report_path       TEXT    DEFAULT '',
        report_date       TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (snapshot_id) REFERENCES weekly_snapshots(snapshot_id)
    )
    ''')

    # ─────────────────────────────────────────────────────────
    # Migration: add snapshot_id to existing tables if missing
    # ─────────────────────────────────────────────────────────
    _migrate_add_column(cursor, "orders",         "snapshot_id",     "INTEGER NOT NULL DEFAULT 0")
    _migrate_add_column(cursor, "orders",         "account_name",    "TEXT DEFAULT ''")
    _migrate_add_column(cursor, "orders",         "items_summary",   "TEXT DEFAULT ''")
    _migrate_add_column(cursor, "collections",    "snapshot_id",     "INTEGER NOT NULL DEFAULT 0")
    _migrate_add_column(cursor, "collections",    "original_amount", "REAL DEFAULT 0")
    _migrate_add_column(cursor, "collections",    "collection_fee",  "REAL DEFAULT 0")
    _migrate_add_column(cursor, "collections",    "is_return",       "INTEGER DEFAULT 0")
    _migrate_add_column(cursor, "collections",    "account_name",    "TEXT DEFAULT ''")
    # weekly_reports — migrate old schema to new
    _migrate_add_column(cursor, "weekly_reports", "snapshot_id",     "INTEGER UNIQUE")
    _migrate_add_column(cursor, "weekly_reports", "label",           "TEXT DEFAULT ''")
    _migrate_add_column(cursor, "weekly_reports", "collection_rate", "REAL DEFAULT 0")
    _migrate_add_column(cursor, "weekly_reports", "paid_count",      "INTEGER DEFAULT 0")
    _migrate_add_column(cursor, "weekly_reports", "unpaid_count",    "INTEGER DEFAULT 0")
    _migrate_add_column(cursor, "weekly_reports", "partial_count",   "INTEGER DEFAULT 0")
    _migrate_add_column(cursor, "weekly_reports", "report_path",     "TEXT DEFAULT ''")

    # Initialize default platforms
    platforms = [
        ('Ilasouq',  0, 0, 0),
        ('Noon',     0, 0, 0),
        ('Trendyol', 0, 0, 0),
        ('Amazon',   0, 0, 0),
        ('Website',  0, 0, 0),   # موقع خاص
    ]
    cursor.executemany('''
    INSERT OR IGNORE INTO platforms (platform_name, commission_rate, tax_rate, shipping_default)
    VALUES (?, ?, ?, ?)
    ''', platforms)

    # Create a default snapshot (id=0) for legacy data
    cursor.execute('''
    INSERT OR IGNORE INTO weekly_snapshots (snapshot_id, label, week_number, year, notes)
    VALUES (0, 'بيانات قديمة', 0, 0, 'Default snapshot for pre-migration data')
    ''')

    conn.commit()
    conn.close()
    print(f"✅ Database {DB_NAME} initialized/migrated successfully.")
    print("   Tables: orders, collections, platforms, weekly_reports, weekly_snapshots")


def _migrate_add_column(cursor, table, column, col_def):
    """Add a column to an existing table if it doesn't already exist."""
    try:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {col_def}")
        print(f"   Migration: added '{column}' to '{table}'")
    except Exception:
        pass  # Column already exists — safe to ignore


if __name__ == "__main__":
    create_database()
