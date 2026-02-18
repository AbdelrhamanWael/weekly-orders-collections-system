"""
محوّل الملفات - File Transformer
يقوم بتحويل ملفات Excel من أي منصة إلى صيغة موحدة
"""

import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Tuple
import re


class FileTransformer:
    """محوّل ملفات Excel إلى صيغة موحدة"""
    
    # خرائط الأعمدة لكل منصة (Column Mappings)
    COLUMN_MAPPINGS = {
        'أمازون': {
            'order_id': ['order id', 'رقم الطلب', 'order-id', 'amazon order id'],
            'order_date': ['order date', 'تاريخ الطلب', 'purchase date', 'التاريخ'],
            'price': ['item price', 'السعر', 'price', 'المبلغ', 'القيمة'],
            'cost': ['cost', 'التكلفة', 'item cost', 'تكلفة المنتج'],
            'shipping': ['shipping', 'الشحن', 'shipping charge', 'رسوم الشحن'],
            'collected_amount': ['collected', 'المحصل', 'amount collected', 'المبلغ المحصل', 'المحصّل', 'المبلغ المحصّل', 'collected amount', 'المبلغ_المحصل'],
            'collection_date': ['collection date', 'تاريخ التحصيل', 'payment date', 'تاريخ الدفع']
        },
        'نون': {
            'order_id': ['order id', 'رقم الطلب', 'order number', 'رقم الأوردر'],
            'order_date': ['order date', 'تاريخ الطلب', 'date', 'التاريخ'],
            'price': ['price', 'السعر', 'total', 'الإجمالي', 'المبلغ'],
            'cost': ['cost', 'التكلفة', 'product cost', 'تكلفة المنتج'],
            'shipping': ['shipping', 'الشحن', 'delivery fee', 'رسوم التوصيل'],
            'collected_amount': ['collected', 'المحصل', 'paid amount', 'المبلغ المدفوع', 'المبلغ المحصل', 'المحصّل', 'المبلغ المحصّل', 'collected amount', 'المبلغ_المحصل'],
            'collection_date': ['collection date', 'تاريخ التحصيل', 'payment date', 'تاريخ الدفع']
        },
        'سلة': {
            'order_id': ['order id', 'رقم الطلب', 'id', 'الرقم'],
            'order_date': ['order date', 'تاريخ الطلب', 'created at', 'تاريخ الإنشاء'],
            'price': ['total', 'الإجمالي', 'price', 'السعر'],
            'cost': ['cost', 'التكلفة', 'product cost', 'تكلفة المنتج'],
            'shipping': ['shipping', 'الشحن', 'shipping cost', 'تكلفة الشحن'],
            'collected_amount': ['collected', 'المحصل', 'paid', 'المدفوع', 'المبلغ المحصل', 'المحصّل', 'المبلغ المحصّل', 'collected amount', 'المبلغ_المحصل'],
            'collection_date': ['collection date', 'تاريخ التحصيل', 'paid at', 'تاريخ الدفع']
        },
        'زد': {
            'order_id': ['order id', 'رقم الطلب', 'id', 'الرقم'],
            'order_date': ['order date', 'تاريخ الطلب', 'date', 'التاريخ'],
            'price': ['total', 'الإجمالي', 'amount', 'المبلغ'],
            'cost': ['cost', 'التكلفة', 'product cost', 'تكلفة المنتج'],
            'shipping': ['shipping', 'الشحن', 'delivery', 'التوصيل'],
            'collected_amount': ['collected', 'المحصل', 'received', 'المستلم', 'المبلغ المحصل', 'المحصّل', 'المبلغ المحصّل', 'collected amount', 'المبلغ_المحصل'],
            'collection_date': ['collection date', 'تاريخ التحصيل', 'received date', 'تاريخ الاستلام']
        },
        'أخرى': {
            'order_id': ['order id', 'رقم الطلب', 'id', 'order number', 'الرقم'],
            'order_date': ['order date', 'تاريخ الطلب', 'date', 'التاريخ', 'created'],
            'price': ['price', 'السعر', 'total', 'الإجمالي', 'amount', 'المبلغ'],
            'cost': ['cost', 'التكلفة', 'product cost', 'تكلفة المنتج'],
            'shipping': ['shipping', 'الشحن', 'delivery', 'التوصيل'],
            'collected_amount': ['collected', 'المحصل', 'paid', 'المدفوع', 'received', 'المبلغ المحصل', 'المحصّل', 'المبلغ المحصّل', 'collected amount', 'المبلغ_المحصل'],
            'collection_date': ['collection date', 'تاريخ التحصيل', 'payment date', 'تاريخ الدفع']
        }
    }
    
    @staticmethod
    def normalize_column_name(col: str) -> str:
        """تطبيع اسم العمود (إزالة المسافات الزائدة والأحرف الخاصة)"""
        cleaned = re.sub(r'[^\w\s\u0600-\u06FF]', '', str(col).strip())
        return ' '.join(cleaned.split()).lower()
    
    @staticmethod
    def find_column(df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
        """
        البحث عن عمود في DataFrame بناءً على قائمة من الأسماء المحتملة
        يدعم المطابقة الدقيقة والجزئية
        
        Args:
            df: DataFrame
            possible_names: قائمة الأسماء المحتملة للعمود
            
        Returns:
            اسم العمود إذا وُجد، None إذا لم يُوجد
        """
        normalized_columns = {FileTransformer.normalize_column_name(col): col for col in df.columns}
        
        # محاولة 1: مطابقة دقيقة
        for name in possible_names:
            normalized_name = FileTransformer.normalize_column_name(name)
            if normalized_name in normalized_columns:
                return normalized_columns[normalized_name]
        
        # محاولة 2: مطابقة جزئية (هل اسم العمود يحتوي على الاسم المطلوب أو العكس)
        for name in possible_names:
            normalized_name = FileTransformer.normalize_column_name(name)
            for norm_col, orig_col in normalized_columns.items():
                if normalized_name in norm_col or norm_col in normalized_name:
                    return orig_col
        
        return None
    
    @staticmethod
    def parse_date(date_value) -> Optional[datetime]:
        """
        تحويل قيمة التاريخ إلى datetime
        يدعم صيغ متعددة
        """
        if pd.isna(date_value):
            return None
        
        # إذا كان التاريخ بالفعل datetime
        if isinstance(date_value, datetime):
            return date_value
        
        # محاولة تحويل النص إلى تاريخ
        date_formats = [
            '%Y-%m-%d',
            '%d/%m/%Y',
            '%m/%d/%Y',
            '%Y/%m/%d',
            '%d-%m-%Y',
            '%Y-%m-%d %H:%M:%S',
            '%d/%m/%Y %H:%M:%S'
        ]
        
        for fmt in date_formats:
            try:
                return datetime.strptime(str(date_value), fmt)
            except:
                continue
        
        # محاولة استخدام pandas
        try:
            return pd.to_datetime(date_value)
        except:
            return None
    
    @staticmethod
    def parse_number(value) -> float:
        """تحويل قيمة إلى رقم"""
        if pd.isna(value):
            return 0.0
        
        # إزالة الفواصل والرموز
        if isinstance(value, str):
            value = value.replace(',', '').replace('SAR', '').replace('ر.س', '').strip()
        
        try:
            return float(value)
        except:
            return 0.0
    
    @classmethod
    def transform_orders_file(
        cls,
        file_path: str,
        platform: str,
        week_number: int,
        year: int
    ) -> Tuple[pd.DataFrame, List[str]]:
        """
        تحويل ملف الطلبات إلى صيغة موحدة
        
        Args:
            file_path: مسار ملف Excel
            platform: اسم المنصة
            week_number: رقم الأسبوع
            year: السنة
            
        Returns:
            (DataFrame موحد, قائمة الأخطاء)
        """
        errors = []
        
        try:
            # قراءة الملف
            df = pd.read_excel(file_path)
            
            if df.empty:
                errors.append("الملف فارغ")
                return pd.DataFrame(), errors
            
            # الحصول على خريطة الأعمدة للمنصة
            column_map = cls.COLUMN_MAPPINGS.get(platform, cls.COLUMN_MAPPINGS['أخرى'])
            
            # البحث عن الأعمدة المطلوبة
            mapped_columns = {}
            for field, possible_names in column_map.items():
                if field in ['collected_amount', 'collection_date']:
                    continue  # هذه الأعمدة خاصة بملف التحصيل
                
                col = cls.find_column(df, possible_names)
                if col:
                    mapped_columns[field] = col
                else:
                    errors.append(f"لم يتم العثور على عمود '{field}'")
            
            # التحقق من وجود الأعمدة الأساسية
            required_fields = ['order_id', 'order_date', 'price']
            missing_required = [f for f in required_fields if f not in mapped_columns]
            
            if missing_required:
                errors.append(f"الأعمدة الأساسية المفقودة: {', '.join(missing_required)}")
                return pd.DataFrame(), errors
            
            # إنشاء DataFrame موحد
            unified_df = pd.DataFrame()
            
            # نسخ الأعمدة الأساسية
            unified_df['order_id'] = df[mapped_columns['order_id']].astype(str)
            unified_df['platform'] = platform
            unified_df['order_date'] = df[mapped_columns['order_date']].apply(cls.parse_date)
            unified_df['price'] = df[mapped_columns['price']].apply(cls.parse_number)
            
            # الأعمدة الاختيارية
            unified_df['cost'] = df[mapped_columns.get('cost', mapped_columns['price'])].apply(cls.parse_number) if 'cost' in mapped_columns else 0.0
            unified_df['shipping'] = df[mapped_columns.get('shipping', mapped_columns['price'])].apply(cls.parse_number) if 'shipping' in mapped_columns else 0.0
            
            # العمولة والضريبة سيتم حسابها لاحقاً بناءً على إعدادات المنصة
            unified_df['commission'] = 0.0
            unified_df['tax'] = 0.0
            
            unified_df['week_number'] = week_number
            unified_df['year'] = year
            unified_df['upload_date'] = datetime.now()
            
            # إزالة الصفوف التي تحتوي على قيم فارغة في الأعمدة الأساسية
            unified_df = unified_df.dropna(subset=['order_id', 'order_date', 'price'])
            
            # إزالة الطلبات المكررة
            duplicates = unified_df['order_id'].duplicated().sum()
            if duplicates > 0:
                errors.append(f"تم العثور على {duplicates} طلب مكرر، سيتم الاحتفاظ بالأول فقط")
                unified_df = unified_df.drop_duplicates(subset=['order_id'], keep='first')
            
            return unified_df, errors
            
        except Exception as e:
            errors.append(f"خطأ في قراءة الملف: {str(e)}")
            return pd.DataFrame(), errors
    
    @classmethod
    def transform_collections_file(
        cls,
        file_path: str,
        platform: str,
        week_number: int,
        year: int
    ) -> Tuple[pd.DataFrame, List[str]]:
        """
        تحويل ملف التحصيل إلى صيغة موحدة
        
        Args:
            file_path: مسار ملف Excel
            platform: اسم المنصة
            week_number: رقم الأسبوع
            year: السنة
            
        Returns:
            (DataFrame موحد, قائمة الأخطاء)
        """
        errors = []
        
        try:
            # قراءة الملف
            df = pd.read_excel(file_path)
            
            if df.empty:
                errors.append("الملف فارغ")
                return pd.DataFrame(), errors
            
            # الحصول على خريطة الأعمدة للمنصة
            column_map = cls.COLUMN_MAPPINGS.get(platform, cls.COLUMN_MAPPINGS['أخرى'])
            
            # البحث عن الأعمدة المطلوبة
            mapped_columns = {}
            for field in ['order_id', 'collected_amount', 'collection_date']:
                col = cls.find_column(df, column_map[field])
                if col:
                    mapped_columns[field] = col
                else:
                    errors.append(f"لم يتم العثور على عمود '{field}'")
            
            # التحقق من وجود الأعمدة الأساسية
            required_fields = ['order_id', 'collected_amount']
            missing_required = [f for f in required_fields if f not in mapped_columns]
            
            if missing_required:
                errors.append(f"الأعمدة الأساسية المفقودة: {', '.join(missing_required)}")
                return pd.DataFrame(), errors
            
            # إنشاء DataFrame موحد
            unified_df = pd.DataFrame()
            
            unified_df['order_id'] = df[mapped_columns['order_id']].astype(str)
            unified_df['collected_amount'] = df[mapped_columns['collected_amount']].apply(cls.parse_number)
            unified_df['collection_date'] = df[mapped_columns.get('collection_date', mapped_columns['order_id'])].apply(cls.parse_date) if 'collection_date' in mapped_columns else datetime.now()
            unified_df['week_number'] = week_number
            unified_df['year'] = year
            unified_df['upload_date'] = datetime.now()
            
            # إزالة الصفوف التي تحتوي على قيم فارغة
            unified_df = unified_df.dropna(subset=['order_id', 'collected_amount'])
            
            # إزالة القيم السالبة (إلا إذا كانت مرتجعات)
            # unified_df = unified_df[unified_df['collected_amount'] >= 0]
            
            return unified_df, errors
            
        except Exception as e:
            errors.append(f"خطأ في قراءة الملف: {str(e)}")
            return pd.DataFrame(), errors
    
    @staticmethod
    def validate_file(file_path: str) -> Tuple[bool, str]:
        """
        التحقق من صحة ملف Excel
        
        Returns:
            (صحيح/خطأ, رسالة)
        """
        try:
            df = pd.read_excel(file_path)
            if df.empty:
                return False, "الملف فارغ"
            return True, "الملف صحيح"
        except Exception as e:
            return False, f"خطأ في قراءة الملف: {str(e)}"
