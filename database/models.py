"""
نماذج البيانات لنظام ربط الطلبات والتحصيل
Data Models for Orders & Collections System
"""

from dataclasses import dataclass
from datetime import datetime
from typing import Optional


@dataclass
class Order:
    """نموذج الطلب - Order Model"""
    order_id: str
    platform: str
    order_date: datetime
    price: float
    cost: float
    shipping: float
    commission: float
    tax: float
    week_number: int
    year: int
    upload_date: datetime
    
    @property
    def total_deductions(self) -> float:
        """إجمالي الخصومات"""
        return self.cost + self.shipping + self.commission + self.tax
    
    @property
    def expected_profit(self) -> float:
        """الربح المتوقع قبل التحصيل"""
        return self.price - self.total_deductions


@dataclass
class Collection:
    """نموذج التحصيل - Collection Model"""
    collection_id: Optional[int]
    order_id: str
    collected_amount: float
    collection_date: datetime
    week_number: int
    year: int
    upload_date: datetime


@dataclass
class Platform:
    """نموذج المنصة - Platform Model"""
    platform_id: Optional[int]
    platform_name: str
    commission_rate: float  # نسبة العمولة (0.15 = 15%)
    tax_rate: float  # نسبة الضريبة (0.15 = 15%)
    shipping_default: float  # تكلفة الشحن الافتراضية


@dataclass
class OrderStatus:
    """حالة الطلب مع التحصيل - Order Status with Collection"""
    order_id: str
    platform: str
    order_date: datetime
    price: float
    cost: float
    shipping: float
    commission: float
    tax: float
    collected_amount: float
    collection_date: Optional[datetime]
    status: str  # 'محصل بالكامل', 'محصل جزئياً', 'غير محصل', 'مرتجع'
    net_profit: float
    days_since_order: int
    week_number: int
    year: int
    
    @property
    def collection_percentage(self) -> float:
        """نسبة التحصيل"""
        if self.price == 0:
            return 0.0
        return (self.collected_amount / self.price) * 100


@dataclass
class WeeklyReport:
    """التقرير الأسبوعي - Weekly Report"""
    report_id: Optional[int]
    week_number: int
    year: int
    total_orders: int
    total_sales: float
    total_collected: float
    total_uncollected: float
    net_profit: float
    collection_rate: float
    report_date: datetime
    
    @property
    def average_order_value(self) -> float:
        """متوسط قيمة الطلب"""
        if self.total_orders == 0:
            return 0.0
        return self.total_sales / self.total_orders
    
    @property
    def profit_margin(self) -> float:
        """هامش الربح"""
        if self.total_sales == 0:
            return 0.0
        return (self.net_profit / self.total_sales) * 100
