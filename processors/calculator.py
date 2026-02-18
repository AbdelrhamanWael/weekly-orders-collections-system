"""
محرك الحسابات - Calculator Engine
يقوم بحساب العمولات والضرائب وصافي الربح
"""

import pandas as pd
from typing import Dict
from database.models import Platform


class Calculator:
    """محرك الحسابات المالية"""
    
    @staticmethod
    def calculate_commission(price: float, commission_rate: float) -> float:
        """
        حساب العمولة
        
        Args:
            price: سعر المنتج
            commission_rate: نسبة العمولة (0.15 = 15%)
            
        Returns:
            قيمة العمولة
        """
        return price * commission_rate
    
    @staticmethod
    def calculate_tax(price: float, tax_rate: float) -> float:
        """
        حساب الضريبة
        
        Args:
            price: سعر المنتج
            tax_rate: نسبة الضريبة (0.15 = 15%)
            
        Returns:
            قيمة الضريبة
        """
        return price * tax_rate
    
    @staticmethod
    def calculate_net_profit(
        collected_amount: float,
        cost: float,
        shipping: float,
        commission: float,
        tax: float
    ) -> float:
        """
        حساب صافي الربح
        
        Formula: Profit = Collected - (Cost + Shipping + Commission + Tax)
        
        Args:
            collected_amount: المبلغ المحصل
            cost: تكلفة المنتج
            shipping: تكلفة الشحن
            commission: العمولة
            tax: الضريبة
            
        Returns:
            صافي الربح
        """
        total_deductions = cost + shipping + commission + tax
        return collected_amount - total_deductions
    
    @staticmethod
    def apply_platform_rates(df: pd.DataFrame, platform: Platform) -> pd.DataFrame:
        """
        تطبيق نسب العمولة والضريبة على DataFrame
        
        Args:
            df: DataFrame يحتوي على الطلبات
            platform: بيانات المنصة
            
        Returns:
            DataFrame مع العمولات والضرائب المحسوبة
        """
        df = df.copy()
        
        # حساب العمولة إذا كانت صفر
        if 'commission' in df.columns:
            df.loc[df['commission'] == 0, 'commission'] = df.loc[df['commission'] == 0, 'price'] * platform.commission_rate
        
        # حساب الضريبة إذا كانت صفر
        if 'tax' in df.columns:
            df.loc[df['tax'] == 0, 'tax'] = df.loc[df['tax'] == 0, 'price'] * platform.tax_rate
        
        # تطبيق الشحن الافتراضي إذا كان صفر
        if 'shipping' in df.columns:
            df.loc[df['shipping'] == 0, 'shipping'] = platform.shipping_default
        
        return df
    
    @staticmethod
    def calculate_collection_rate(total_collected: float, total_sales: float) -> float:
        """
        حساب نسبة التحصيل
        
        Args:
            total_collected: إجمالي المحصل
            total_sales: إجمالي المبيعات
            
        Returns:
            نسبة التحصيل (%)
        """
        if total_sales == 0:
            return 0.0
        return (total_collected / total_sales) * 100
    
    @staticmethod
    def calculate_profit_margin(net_profit: float, total_sales: float) -> float:
        """
        حساب هامش الربح
        
        Args:
            net_profit: صافي الربح
            total_sales: إجمالي المبيعات
            
        Returns:
            هامش الربح (%)
        """
        if total_sales == 0:
            return 0.0
        return (net_profit / total_sales) * 100
    
    @staticmethod
    def calculate_summary_stats(df: pd.DataFrame) -> Dict:
        """
        حساب الإحصائيات الملخصة
        
        Args:
            df: DataFrame يحتوي على الطلبات المطابقة
            
        Returns:
            قاموس يحتوي على الإحصائيات
        """
        if df.empty:
            return {
                'total_orders': 0,
                'total_sales': 0.0,
                'total_collected': 0.0,
                'total_uncollected': 0.0,
                'net_profit': 0.0,
                'collection_rate': 0.0,
                'profit_margin': 0.0,
                'avg_order_value': 0.0,
                'fully_collected': 0,
                'partially_collected': 0,
                'uncollected': 0,
                'returned': 0
            }
        
        total_orders = len(df)
        total_sales = df['price'].sum()
        total_collected = df['collected_amount'].sum()
        total_uncollected = total_sales - total_collected
        net_profit = df['net_profit'].sum()
        
        # حساب النسب
        collection_rate = Calculator.calculate_collection_rate(total_collected, total_sales)
        profit_margin = Calculator.calculate_profit_margin(net_profit, total_sales)
        avg_order_value = total_sales / total_orders if total_orders > 0 else 0.0
        
        # حساب عدد الطلبات حسب الحالة
        status_counts = df['status'].value_counts().to_dict()
        
        return {
            'total_orders': total_orders,
            'total_sales': round(total_sales, 2),
            'total_collected': round(total_collected, 2),
            'total_uncollected': round(total_uncollected, 2),
            'net_profit': round(net_profit, 2),
            'collection_rate': round(collection_rate, 2),
            'profit_margin': round(profit_margin, 2),
            'avg_order_value': round(avg_order_value, 2),
            'fully_collected': status_counts.get('محصل بالكامل', 0),
            'partially_collected': status_counts.get('محصل جزئياً', 0),
            'uncollected': status_counts.get('غير محصل', 0),
            'returned': status_counts.get('مرتجع', 0)
        }
    
    @staticmethod
    def calculate_platform_stats(df: pd.DataFrame) -> pd.DataFrame:
        """
        حساب الإحصائيات لكل منصة
        
        Args:
            df: DataFrame يحتوي على الطلبات المطابقة
            
        Returns:
            DataFrame يحتوي على إحصائيات كل منصة
        """
        if df.empty:
            return pd.DataFrame()
        
        platform_stats = df.groupby('platform').agg({
            'order_id': 'count',
            'price': 'sum',
            'collected_amount': 'sum',
            'net_profit': 'sum'
        }).reset_index()
        
        platform_stats.columns = ['platform', 'total_orders', 'total_sales', 'total_collected', 'net_profit']
        
        # حساب النسب
        platform_stats['collection_rate'] = platform_stats.apply(
            lambda row: Calculator.calculate_collection_rate(row['total_collected'], row['total_sales']),
            axis=1
        )
        
        platform_stats['profit_margin'] = platform_stats.apply(
            lambda row: Calculator.calculate_profit_margin(row['net_profit'], row['total_sales']),
            axis=1
        )
        
        # تقريب القيم
        for col in ['total_sales', 'total_collected', 'net_profit', 'collection_rate', 'profit_margin']:
            platform_stats[col] = platform_stats[col].round(2)
        
        return platform_stats
