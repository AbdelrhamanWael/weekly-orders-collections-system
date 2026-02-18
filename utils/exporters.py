"""
مصدّر التقارير - Report Exporter
يقوم بتصدير التقارير إلى ملفات Excel منسقة
"""

import pandas as pd
from datetime import datetime
from typing import Dict
import xlsxwriter
from io import BytesIO


class ReportExporter:
    """مصدّر التقارير إلى Excel"""
    
    @staticmethod
    def export_weekly_report(
        summary_stats: Dict,
        matched_orders_df: pd.DataFrame,
        platform_stats_df: pd.DataFrame,
        week_number: int,
        year: int
    ) -> BytesIO:
        """
        تصدير التقرير الأسبوعي الكامل إلى Excel
        
        Args:
            summary_stats: الإحصائيات الملخصة
            matched_orders_df: DataFrame الطلبات المطابقة
            platform_stats_df: DataFrame إحصائيات المنصات
            week_number: رقم الأسبوع
            year: السنة
            
        Returns:
            BytesIO يحتوي على ملف Excel
        """
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # تنسيقات Excel
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4472C4',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 16,
                'bg_color': '#2E75B6',
                'font_color': 'white',
                'align': 'center',
                'valign': 'vcenter'
            })
            
            currency_format = workbook.add_format({
                'num_format': '#,##0.00',
                'border': 1
            })
            
            percentage_format = workbook.add_format({
                'num_format': '0.00%',
                'border': 1
            })
            
            # ========== صفحة الملخص ==========
            summary_sheet = workbook.add_worksheet('الملخص')
            
            # العنوان
            summary_sheet.merge_range('A1:D1', f'التقرير الأسبوعي - الأسبوع {week_number} / {year}', title_format)
            summary_sheet.write('A2', f'تاريخ التقرير: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
            
            # الإحصائيات الرئيسية
            row = 4
            summary_sheet.write(row, 0, 'المؤشر', header_format)
            summary_sheet.write(row, 1, 'القيمة', header_format)
            
            metrics = [
                ('إجمالي الطلبات', summary_stats['total_orders'], None),
                ('إجمالي المبيعات', summary_stats['total_sales'], currency_format),
                ('إجمالي المحصل', summary_stats['total_collected'], currency_format),
                ('غير المحصل', summary_stats['total_uncollected'], currency_format),
                ('صافي الربح', summary_stats['net_profit'], currency_format),
                ('نسبة التحصيل', summary_stats['collection_rate'] / 100, percentage_format),
                ('هامش الربح', summary_stats['profit_margin'] / 100, percentage_format),
                ('متوسط قيمة الطلب', summary_stats['avg_order_value'], currency_format),
            ]
            
            for metric_name, metric_value, fmt in metrics:
                row += 1
                summary_sheet.write(row, 0, metric_name)
                if fmt:
                    summary_sheet.write(row, 1, metric_value, fmt)
                else:
                    summary_sheet.write(row, 1, metric_value)
            
            # حالات الطلبات
            row += 3
            summary_sheet.write(row, 0, 'حالة الطلب', header_format)
            summary_sheet.write(row, 1, 'العدد', header_format)
            
            statuses = [
                ('محصل بالكامل ✅', summary_stats['fully_collected']),
                ('محصل جزئياً ⚠️', summary_stats['partially_collected']),
                ('غير محصل ❌', summary_stats['uncollected']),
                ('مرتجع 🔄', summary_stats['returned']),
            ]
            
            for status_name, status_count in statuses:
                row += 1
                summary_sheet.write(row, 0, status_name)
                summary_sheet.write(row, 1, status_count)
            
            # تعديل عرض الأعمدة
            summary_sheet.set_column('A:A', 25)
            summary_sheet.set_column('B:B', 20)
            
            # ========== صفحة إحصائيات المنصات ==========
            if not platform_stats_df.empty:
                platform_stats_df.to_excel(writer, sheet_name='إحصائيات المنصات', index=False, startrow=1)
                platform_sheet = writer.sheets['إحصائيات المنصات']
                
                # العنوان
                platform_sheet.merge_range('A1:G1', 'إحصائيات المنصات', title_format)
                
                # تنسيق الرأس
                for col_num, value in enumerate(platform_stats_df.columns.values):
                    platform_sheet.write(1, col_num, value, header_format)
                
                # تنسيق البيانات
                for row_num in range(len(platform_stats_df)):
                    platform_sheet.write(row_num + 2, 2, platform_stats_df.iloc[row_num, 2], currency_format)
                    platform_sheet.write(row_num + 2, 3, platform_stats_df.iloc[row_num, 3], currency_format)
                    platform_sheet.write(row_num + 2, 4, platform_stats_df.iloc[row_num, 4], currency_format)
                    platform_sheet.write(row_num + 2, 5, platform_stats_df.iloc[row_num, 5] / 100, percentage_format)
                    platform_sheet.write(row_num + 2, 6, platform_stats_df.iloc[row_num, 6] / 100, percentage_format)
                
                platform_sheet.set_column('A:A', 15)
                platform_sheet.set_column('B:G', 18)
            
            # ========== صفحة الطلبات المفصلة ==========
            if not matched_orders_df.empty:
                # إعادة ترتيب الأعمدة
                columns_order = [
                    'order_id', 'platform', 'order_date', 'price', 'cost',
                    'shipping', 'commission', 'tax', 'collected_amount',
                    'net_profit', 'status', 'days_since_order'
                ]
                
                export_df = matched_orders_df[columns_order].copy()
                
                # تسمية الأعمدة بالعربية
                export_df.columns = [
                    'رقم الطلب', 'المنصة', 'تاريخ الطلب', 'السعر', 'التكلفة',
                    'الشحن', 'العمولة', 'الضريبة', 'المحصل',
                    'صافي الربح', 'الحالة', 'عدد الأيام'
                ]
                
                export_df.to_excel(writer, sheet_name='الطلبات المفصلة', index=False, startrow=1)
                orders_sheet = writer.sheets['الطلبات المفصلة']
                
                # العنوان
                orders_sheet.merge_range(0, 0, 0, len(export_df.columns) - 1, 'الطلبات المفصلة', title_format)
                
                # تنسيق الرأس
                for col_num, value in enumerate(export_df.columns.values):
                    orders_sheet.write(1, col_num, value, header_format)
                
                # تنسيق الأعمدة المالية
                money_columns = [3, 4, 5, 6, 7, 8, 9]  # السعر، التكلفة، الشحن، العمولة، الضريبة، المحصل، صافي الربح
                for col in money_columns:
                    orders_sheet.set_column(col, col, 15, currency_format)
                
                orders_sheet.set_column('A:A', 20)  # رقم الطلب
                orders_sheet.set_column('B:B', 12)  # المنصة
                orders_sheet.set_column('C:C', 15)  # تاريخ الطلب
                orders_sheet.set_column('K:K', 15)  # الحالة
                orders_sheet.set_column('L:L', 12)  # عدد الأيام
        
        output.seek(0)
        return output
    
    @staticmethod
    def export_uncollected_orders(
        uncollected_df: pd.DataFrame,
        days_threshold: int
    ) -> BytesIO:
        """
        تصدير الطلبات غير المحصلة
        
        Args:
            uncollected_df: DataFrame الطلبات غير المحصلة
            days_threshold: عدد الأيام المحدد
            
        Returns:
            BytesIO يحتوي على ملف Excel
        """
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # تنسيقات
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#C00000',
                'font_color': 'white',
                'border': 1,
                'align': 'center'
            })
            
            title_format = workbook.add_format({
                'bold': True,
                'font_size': 16,
                'bg_color': '#C00000',
                'font_color': 'white',
                'align': 'center'
            })
            
            currency_format = workbook.add_format({
                'num_format': '#,##0.00',
                'border': 1
            })
            
            # إعداد البيانات
            export_df = uncollected_df[[
                'order_id', 'platform', 'order_date', 'price',
                'days_since_order', 'status'
            ]].copy()
            
            export_df.columns = [
                'رقم الطلب', 'المنصة', 'تاريخ الطلب', 'المبلغ',
                'عدد الأيام', 'الحالة'
            ]
            
            export_df.to_excel(writer, sheet_name='الطلبات المتأخرة', index=False, startrow=2)
            sheet = writer.sheets['الطلبات المتأخرة']
            
            # العنوان
            sheet.merge_range(0, 0, 0, len(export_df.columns) - 1,
                            f'الطلبات غير المحصلة بعد {days_threshold} يوم', title_format)
            sheet.write(1, 0, f'تاريخ التقرير: {datetime.now().strftime("%Y-%m-%d %H:%M")}')
            
            # تنسيق الرأس
            for col_num, value in enumerate(export_df.columns.values):
                sheet.write(2, col_num, value, header_format)
            
            # تنسيق الأعمدة
            sheet.set_column('A:A', 20)
            sheet.set_column('B:B', 12)
            sheet.set_column('C:C', 15)
            sheet.set_column('D:D', 15, currency_format)
            sheet.set_column('E:E', 12)
            sheet.set_column('F:F', 15)
        
        output.seek(0)
        return output
