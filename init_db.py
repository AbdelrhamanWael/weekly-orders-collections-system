# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Module   : init_db.py — Database Initialization
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
# ============================================================
import sqlite3
import os
from datetime import datetime

DB_NAME = "finance_system.db"

def create_database():
    if os.path.exists(DB_NAME):
        print(f"Database {DB_NAME} already exists.")
        # Optional: Add logic to back up existing DB or confirm overwrite
    
    conn = sqlite3.connect(DB_NAME)
    cursor = conn.cursor()
    
    # 1. Orders Table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS orders (
        order_id TEXT PRIMARY KEY,
        platform TEXT NOT NULL,
        order_date DATE,
        price REAL,
        cost REAL DEFAULT 0,
        shipping REAL DEFAULT 0,
        commission REAL DEFAULT 0,
        tax REAL DEFAULT 0,
        week_number INTEGER,
        year INTEGER,
        upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # 2. Collections Table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS collections (
        collection_id INTEGER PRIMARY KEY AUTOINCREMENT,
        order_id TEXT,
        collected_amount REAL,
        collection_date DATE,
        week_number INTEGER,
        year INTEGER,
        upload_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY (order_id) REFERENCES orders(order_id)
    )
    ''')
    
    # 3. Platforms Table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS platforms (
        platform_id INTEGER PRIMARY KEY AUTOINCREMENT,
        platform_name TEXT UNIQUE,
        commission_rate REAL DEFAULT 0,
        tax_rate REAL DEFAULT 0,
        shipping_default REAL DEFAULT 0
    )
    ''')
    
    # 4. Weekly Reports Table
    cursor.execute('''
    CREATE TABLE IF NOT EXISTS weekly_reports (
        report_id INTEGER PRIMARY KEY AUTOINCREMENT,
        week_number INTEGER,
        year INTEGER,
        total_orders INTEGER DEFAULT 0,
        total_sales REAL DEFAULT 0,
        total_collected REAL DEFAULT 0,
        total_uncollected REAL DEFAULT 0,
        net_profit REAL DEFAULT 0,
        report_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )
    ''')
    
    # Initialize default platforms
    platforms = [
        ('Ilasouq', 0, 0, 0),
        ('Noon', 0, 0, 0),
        ('Trendyol', 0, 0, 0),
        ('Amazon', 0, 0, 0)
    ]
    
    cursor.executemany('''
    INSERT OR IGNORE INTO platforms (platform_name, commission_rate, tax_rate, shipping_default)
    VALUES (?, ?, ?, ?)
    ''', platforms)
    
    conn.commit()
    conn.close()
    print(f"Database {DB_NAME} initialized successfully with tables: orders, collections, platforms, weekly_reports.")

if __name__ == "__main__":
    create_database()
