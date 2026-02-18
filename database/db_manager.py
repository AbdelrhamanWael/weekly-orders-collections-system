"""
مدير قاعدة البيانات - Database Manager
يدير جميع عمليات قاعدة البيانات SQLite
"""

import sqlite3
from datetime import datetime
from typing import List, Optional, Tuple
from pathlib import Path
import pandas as pd

from .models import Order, Collection, Platform, OrderStatus, WeeklyReport


class DatabaseManager:
    """مدير قاعدة البيانات"""
    
    def __init__(self, db_path: str = "data/orders.db"):
        """
        تهيئة مدير قاعدة البيانات
        
        Args:
            db_path: مسار ملف قاعدة البيانات
        """
        self.db_path = db_path
        # إنشاء مجلد data إذا لم يكن موجوداً
        Path(db_path).parent.mkdir(parents=True, exist_ok=True)
        self._create_tables()
        self._initialize_platforms()
    
    def _get_connection(self) -> sqlite3.Connection:
        """إنشاء اتصال بقاعدة البيانات"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn
    
    def _create_tables(self):
        """إنشاء جداول قاعدة البيانات"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        # جدول الطلبات
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS orders (
                order_id TEXT PRIMARY KEY,
                platform TEXT NOT NULL,
                order_date DATE NOT NULL,
                price REAL NOT NULL,
                cost REAL NOT NULL,
                shipping REAL NOT NULL,
                commission REAL NOT NULL,
                tax REAL NOT NULL,
                week_number INTEGER NOT NULL,
                year INTEGER NOT NULL,
                upload_date TIMESTAMP NOT NULL
            )
        """)
        
        # جدول التحصيل
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS collections (
                collection_id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id TEXT NOT NULL,
                collected_amount REAL NOT NULL,
                collection_date DATE NOT NULL,
                week_number INTEGER NOT NULL,
                year INTEGER NOT NULL,
                upload_date TIMESTAMP NOT NULL,
                FOREIGN KEY (order_id) REFERENCES orders (order_id)
            )
        """)
        
        # جدول المنصات
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS platforms (
                platform_id INTEGER PRIMARY KEY AUTOINCREMENT,
                platform_name TEXT UNIQUE NOT NULL,
                commission_rate REAL NOT NULL,
                tax_rate REAL NOT NULL,
                shipping_default REAL NOT NULL
            )
        """)
        
        # جدول التقارير الأسبوعية
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS weekly_reports (
                report_id INTEGER PRIMARY KEY AUTOINCREMENT,
                week_number INTEGER NOT NULL,
                year INTEGER NOT NULL,
                total_orders INTEGER NOT NULL,
                total_sales REAL NOT NULL,
                total_collected REAL NOT NULL,
                total_uncollected REAL NOT NULL,
                net_profit REAL NOT NULL,
                collection_rate REAL NOT NULL,
                report_date TIMESTAMP NOT NULL,
                UNIQUE(week_number, year)
            )
        """)
        
        conn.commit()
        conn.close()
    
    def _initialize_platforms(self):
        """تهيئة المنصات الافتراضية"""
        default_platforms = [
            Platform(None, "أمازون", 0.15, 0.15, 25.0),
            Platform(None, "نون", 0.12, 0.15, 20.0),
            Platform(None, "سلة", 0.10, 0.15, 15.0),
            Platform(None, "زد", 0.10, 0.15, 15.0),
            Platform(None, "أخرى", 0.15, 0.15, 20.0),
        ]
        
        for platform in default_platforms:
            self.add_platform(platform)
    
    # ==================== عمليات المنصات ====================
    
    def add_platform(self, platform: Platform) -> bool:
        """إضافة منصة جديدة"""
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR IGNORE INTO platforms (platform_name, commission_rate, tax_rate, shipping_default)
                VALUES (?, ?, ?, ?)
            """, (platform.platform_name, platform.commission_rate, platform.tax_rate, platform.shipping_default))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"خطأ في إضافة المنصة: {e}")
            return False
    
    def get_platform(self, platform_name: str) -> Optional[Platform]:
        """الحصول على بيانات منصة"""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM platforms WHERE platform_name = ?", (platform_name,))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return Platform(
                platform_id=row['platform_id'],
                platform_name=row['platform_name'],
                commission_rate=row['commission_rate'],
                tax_rate=row['tax_rate'],
                shipping_default=row['shipping_default']
            )
        return None
    
    def get_all_platforms(self) -> List[str]:
        """الحصول على جميع أسماء المنصات"""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT platform_name FROM platforms ORDER BY platform_name")
        platforms = [row['platform_name'] for row in cursor.fetchall()]
        conn.close()
        return platforms
    
    # ==================== عمليات الطلبات ====================
    
    def add_orders_bulk(self, orders: List[Order]) -> int:
        """إضافة مجموعة من الطلبات دفعة واحدة"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        added_count = 0
        for order in orders:
            try:
                cursor.execute("""
                    INSERT OR REPLACE INTO orders 
                    (order_id, platform, order_date, price, cost, shipping, commission, tax, week_number, year, upload_date)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    order.order_id, order.platform, order.order_date, order.price,
                    order.cost, order.shipping, order.commission, order.tax,
                    order.week_number, order.year, order.upload_date
                ))
                added_count += 1
            except Exception as e:
                print(f"خطأ في إضافة الطلب {order.order_id}: {e}")
        
        conn.commit()
        conn.close()
        return added_count
    
    def get_orders_by_week(self, week_number: int, year: int) -> pd.DataFrame:
        """الحصول على طلبات أسبوع معين"""
        conn = self._get_connection()
        query = """
            SELECT * FROM orders 
            WHERE week_number = ? AND year = ?
            ORDER BY order_date DESC
        """
        df = pd.read_sql_query(query, conn, params=(week_number, year))
        conn.close()
        return df
    
    # ==================== عمليات التحصيل ====================
    
    def add_collections_bulk(self, collections: List[Collection]) -> int:
        """إضافة مجموعة من التحصيلات دفعة واحدة"""
        conn = self._get_connection()
        cursor = conn.cursor()
        
        added_count = 0
        for collection in collections:
            try:
                cursor.execute("""
                    INSERT INTO collections 
                    (order_id, collected_amount, collection_date, week_number, year, upload_date)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    collection.order_id, collection.collected_amount, collection.collection_date,
                    collection.week_number, collection.year, collection.upload_date
                ))
                added_count += 1
            except Exception as e:
                print(f"خطأ في إضافة التحصيل للطلب {collection.order_id}: {e}")
        
        conn.commit()
        conn.close()
        return added_count
    
    def get_collections_by_week(self, week_number: int, year: int) -> pd.DataFrame:
        """الحصول على تحصيلات أسبوع معين"""
        conn = self._get_connection()
        query = """
            SELECT * FROM collections 
            WHERE week_number = ? AND year = ?
            ORDER BY collection_date DESC
        """
        df = pd.read_sql_query(query, conn, params=(week_number, year))
        conn.close()
        return df
    
    # ==================== عمليات التقارير ====================
    
    def get_matched_orders(self, week_number: int, year: int) -> pd.DataFrame:
        """
        الحصول على الطلبات مع التحصيل (Matched Orders)
        يقوم بعمل LEFT JOIN بين الطلبات والتحصيل
        """
        conn = self._get_connection()
        query = """
            SELECT 
                o.order_id,
                o.platform,
                o.order_date,
                o.price,
                o.cost,
                o.shipping,
                o.commission,
                o.tax,
                COALESCE(SUM(c.collected_amount), 0) as collected_amount,
                MAX(c.collection_date) as collection_date,
                o.week_number,
                o.year
            FROM orders o
            LEFT JOIN collections c ON o.order_id = c.order_id
            WHERE o.week_number = ? AND o.year = ?
            GROUP BY o.order_id
            ORDER BY o.order_date DESC
        """
        df = pd.read_sql_query(query, conn, params=(week_number, year))
        conn.close()
        
        # حساب الحالة وصافي الربح
        if not df.empty:
            df['status'] = df.apply(self._determine_status, axis=1)
            df['net_profit'] = df['collected_amount'] - (df['cost'] + df['shipping'] + df['commission'] + df['tax'])
            df['days_since_order'] = (datetime.now() - pd.to_datetime(df['order_date'])).dt.days
        
        return df
    
    def _determine_status(self, row) -> str:
        """تحديد حالة الطلب"""
        if row['collected_amount'] == 0:
            return 'غير محصل'
        elif row['collected_amount'] >= row['price']:
            return 'محصل بالكامل'
        elif row['collected_amount'] < 0:
            return 'مرتجع'
        else:
            return 'محصل جزئياً'
    
    def save_weekly_report(self, report: WeeklyReport) -> bool:
        """حفظ التقرير الأسبوعي"""
        try:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute("""
                INSERT OR REPLACE INTO weekly_reports 
                (week_number, year, total_orders, total_sales, total_collected, 
                 total_uncollected, net_profit, collection_rate, report_date)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                report.week_number, report.year, report.total_orders, report.total_sales,
                report.total_collected, report.total_uncollected, report.net_profit,
                report.collection_rate, report.report_date
            ))
            conn.commit()
            conn.close()
            return True
        except Exception as e:
            print(f"خطأ في حفظ التقرير: {e}")
            return False
    
    def get_weekly_report(self, week_number: int, year: int) -> Optional[WeeklyReport]:
        """الحصول على تقرير أسبوعي"""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT * FROM weekly_reports 
            WHERE week_number = ? AND year = ?
        """, (week_number, year))
        row = cursor.fetchone()
        conn.close()
        
        if row:
            return WeeklyReport(
                report_id=row['report_id'],
                week_number=row['week_number'],
                year=row['year'],
                total_orders=row['total_orders'],
                total_sales=row['total_sales'],
                total_collected=row['total_collected'],
                total_uncollected=row['total_uncollected'],
                net_profit=row['net_profit'],
                collection_rate=row['collection_rate'],
                report_date=row['report_date']
            )
        return None
    
    def get_all_weeks(self) -> List[Tuple[int, int]]:
        """الحصول على جميع الأسابيع المتاحة"""
        conn = self._get_connection()
        cursor = conn.cursor()
        cursor.execute("""
            SELECT DISTINCT week_number, year 
            FROM orders 
            ORDER BY year DESC, week_number DESC
        """)
        weeks = [(row['week_number'], row['year']) for row in cursor.fetchall()]
        conn.close()
        return weeks
