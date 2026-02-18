# ============================================================
#  Project  : Weekly Orders & Collections Reconciliation System
#             نظام ربط الطلبات والتحصيل الأسبوعي
#  Module   : generate_report.py — Excel Report Generator
#  Developer : Abdelrhaman Wael Mohammed
#  Email     : abdelrhamanwael8@gmail.com
#  LinkedIn  : linkedin.com/in/abdelrhaman-wael-mohammed-790171366
#  Created   : February 2026
# ============================================================
import sqlite3
import pandas as pd
import os
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

DB_NAME = "finance_system.db"
OUTPUT_DIR = "reports"

def get_db_connection():
    return sqlite3.connect(DB_NAME)

def generate_weekly_report():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
        
    conn = get_db_connection()
    
    # Query to join Orders and Collections, including cost fields for Net Profit
    query = '''
    SELECT 
        o.order_id,
        o.platform,
        o.order_date,
        o.week_number,
        o.price         AS expected_amount,
        o.cost          AS cost,
        o.shipping      AS shipping,
        o.commission    AS commission,
        o.tax           AS tax,
        COALESCE(SUM(c.collected_amount), 0) AS collected_amount,
        (COALESCE(SUM(c.collected_amount), 0) - o.price) AS difference,
        MAX(c.collection_date) AS last_collection_date,
        COUNT(c.collected_amount) AS transaction_count
    FROM orders o
    LEFT JOIN collections c ON o.order_id = c.order_id
    GROUP BY o.order_id
    ORDER BY o.platform, o.order_date
    '''
    
    df = pd.read_sql_query(query, conn)
    conn.close()
    
    # Calculate Net Profit: collected - (cost + shipping + commission + tax)
    df['net_profit'] = df['collected_amount'] - (
        df['cost'] + df['shipping'] + df['commission'] + df['tax']
    )
    
    # Add Status Column
    def get_status(row):
        diff = row['expected_amount'] - row['collected_amount']
        if row['collected_amount'] == 0:
            return 'غير مدفوع'
        elif abs(diff) < 0.1:
            return 'مدفوع'
        elif row['collected_amount'] > row['expected_amount']:
            return 'زيادة في التحصيل'
        else:
            return 'مدفوع جزئياً'

    df['status'] = df.apply(get_status, axis=1)
    
    # Reorder columns for clarity
    display_cols = [
        'order_id', 'platform', 'order_date', 'week_number',
        'expected_amount', 'collected_amount', 'difference',
        'cost', 'shipping', 'commission', 'tax', 'net_profit',
        'status', 'last_collection_date', 'transaction_count'
    ]
    df = df[display_cols]

    # Arabic column names for display
    col_names_ar = {
        'order_id': 'رقم الطلب',
        'platform': 'المنصة',
        'order_date': 'تاريخ الطلب',
        'week_number': 'رقم الأسبوع',
        'expected_amount': 'المبلغ المتوقع',
        'collected_amount': 'المبلغ المحصل',
        'difference': 'الفرق',
        'cost': 'التكلفة',
        'shipping': 'الشحن',
        'commission': 'العمولة',
        'tax': 'الضريبة',
        'net_profit': 'صافي الربح',
        'status': 'الحالة',
        'last_collection_date': 'آخر تحصيل',
        'transaction_count': 'عدد الحركات'
    }
    df_display = df.rename(columns=col_names_ar)

    # --- Summary Stats ---
    total_expected   = df['expected_amount'].sum()
    total_collected  = df['collected_amount'].sum()
    total_net_profit = df['net_profit'].sum()
    collection_rate  = (total_collected / total_expected * 100) if total_expected > 0 else 0
    paid_count       = (df['status'] == 'مدفوع').sum()
    unpaid_count     = (df['status'] == 'غير مدفوع').sum()
    partial_count    = (df['status'] == 'مدفوع جزئياً').sum()
    total_orders     = len(df)

    # Generate Filename
    current_date = datetime.now().strftime("%Y-%m-%d")
    report_path = os.path.join(OUTPUT_DIR, f"Weekly_Reconciliation_Report_{current_date}.xlsx")
    
    # --- Excel Styles ---
    header_fill    = PatternFill("solid", fgColor="1F3864")   # Dark navy
    subheader_fill = PatternFill("solid", fgColor="2E75B6")   # Blue
    green_fill     = PatternFill("solid", fgColor="E2EFDA")   # Light green
    red_fill       = PatternFill("solid", fgColor="FCE4D6")   # Light red
    yellow_fill    = PatternFill("solid", fgColor="FFF2CC")   # Light yellow
    white_fill     = PatternFill("solid", fgColor="FFFFFF")
    
    header_font    = Font(bold=True, color="FFFFFF", size=12)
    title_font     = Font(bold=True, color="1F3864", size=14)
    kpi_font       = Font(bold=True, color="1F3864", size=22)
    label_font     = Font(bold=True, color="595959", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='BFBFBF'),
        right=Side(style='thin', color='BFBFBF'),
        top=Side(style='thin', color='BFBFBF'),
        bottom=Side(style='thin', color='BFBFBF')
    )
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right_align  = Alignment(horizontal='right', vertical='center')

    with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
        
        # =====================================================
        # SHEET 1: Executive Dashboard (الملخص التنفيذي)
        # =====================================================
        # Write a placeholder df to create the sheet
        pd.DataFrame().to_excel(writer, sheet_name='الملخص التنفيذي', index=False)
        ws = writer.sheets['الملخص التنفيذي']
        ws.sheet_view.rightToLeft = True

        # Title
        ws.merge_cells('B2:H2')
        title_cell = ws['B2']
        title_cell.value = f"تقرير المطابقة الأسبوعي - {current_date}"
        title_cell.font = Font(bold=True, color="1F3864", size=18)
        title_cell.alignment = center_align
        title_cell.fill = PatternFill("solid", fgColor="D6E4F0")

        # KPI Cards layout: Row 4 onwards
        kpis = [
            ("إجمالي المبيعات",        f"{total_expected:,.2f} ر.س",   "D6E4F0", "1F3864"),
            ("إجمالي المحصل",          f"{total_collected:,.2f} ر.س",  "E2EFDA", "375623"),
            ("صافي الربح الإجمالي",    f"{total_net_profit:,.2f} ر.س", "EBF3FB", "1F3864"),
            ("نسبة التحصيل",           f"{collection_rate:.1f}%",       "FFF2CC", "7F6000"),
        ]
        
        col_positions = [2, 4, 6, 8]  # B, D, F, H
        for i, (label, value, bg, fg) in enumerate(kpis):
            col = col_positions[i]
            col_letter = get_column_letter(col)
            next_col   = get_column_letter(col + 1)
            
            # Merge 2 columns for each KPI
            ws.merge_cells(f'{col_letter}4:{next_col}4')
            ws.merge_cells(f'{col_letter}5:{next_col}5')
            ws.merge_cells(f'{col_letter}6:{next_col}6')
            
            label_cell = ws[f'{col_letter}4']
            label_cell.value = label
            label_cell.font = Font(bold=True, color="595959", size=11)
            label_cell.alignment = center_align
            label_cell.fill = PatternFill("solid", fgColor=bg)
            
            value_cell = ws[f'{col_letter}5']
            value_cell.value = value
            value_cell.font = Font(bold=True, color=fg, size=20)
            value_cell.alignment = center_align
            value_cell.fill = PatternFill("solid", fgColor=bg)
            
            ws[f'{col_letter}6'].fill = PatternFill("solid", fgColor=bg)
            
            # Set row heights
            ws.row_dimensions[4].height = 25
            ws.row_dimensions[5].height = 45
            ws.row_dimensions[6].height = 10

        # Order Status Summary (Row 8)
        ws.merge_cells('B8:I8')
        ws['B8'].value = "ملخص حالات الطلبات"
        ws['B8'].font = Font(bold=True, color="FFFFFF", size=12)
        ws['B8'].fill = PatternFill("solid", fgColor="2E75B6")
        ws['B8'].alignment = center_align
        ws.row_dimensions[8].height = 25

        status_data = [
            ("إجمالي الطلبات",    total_orders,   "D6E4F0"),
            ("مدفوع",             paid_count,     "E2EFDA"),
            ("غير مدفوع",         unpaid_count,   "FCE4D6"),
            ("مدفوع جزئياً",      partial_count,  "FFF2CC"),
        ]
        
        for i, (label, val, bg) in enumerate(status_data):
            col = get_column_letter(2 + i * 2)
            next_col = get_column_letter(3 + i * 2)
            ws.merge_cells(f'{col}9:{next_col}9')
            ws.merge_cells(f'{col}10:{next_col}10')
            ws[f'{col}9'].value = label
            ws[f'{col}9'].font = Font(bold=True, color="595959", size=10)
            ws[f'{col}9'].alignment = center_align
            ws[f'{col}9'].fill = PatternFill("solid", fgColor=bg)
            ws[f'{col}10'].value = val
            ws[f'{col}10'].font = Font(bold=True, color="1F3864", size=18)
            ws[f'{col}10'].alignment = center_align
            ws[f'{col}10'].fill = PatternFill("solid", fgColor=bg)
            ws.row_dimensions[9].height = 22
            ws.row_dimensions[10].height = 40

        # Platform breakdown (Row 12)
        ws.merge_cells('B12:I12')
        ws['B12'].value = "التفصيل حسب المنصة"
        ws['B12'].font = Font(bold=True, color="FFFFFF", size=12)
        ws['B12'].fill = PatternFill("solid", fgColor="2E75B6")
        ws['B12'].alignment = center_align
        ws.row_dimensions[12].height = 25

        platform_summary = df.groupby('platform').agg(
            orders=('order_id', 'count'),
            expected=('expected_amount', 'sum'),
            collected=('collected_amount', 'sum'),
            net_profit=('net_profit', 'sum')
        ).reset_index()

        headers_ps = ['المنصة', 'عدد الطلبات', 'المبلغ المتوقع', 'المبلغ المحصل', 'صافي الربح']
        for j, h in enumerate(headers_ps):
            cell = ws.cell(row=13, column=2+j)
            cell.value = h
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill("solid", fgColor="1F3864")
            cell.alignment = center_align
            cell.border = thin_border
        ws.row_dimensions[13].height = 22

        for r_idx, row in platform_summary.iterrows():
            row_num = 14 + r_idx
            vals = [row['platform'], row['orders'],
                    f"{row['expected']:,.2f}", f"{row['collected']:,.2f}", f"{row['net_profit']:,.2f}"]
            for j, v in enumerate(vals):
                cell = ws.cell(row=row_num, column=2+j)
                cell.value = v
                cell.alignment = center_align
                cell.border = thin_border
                cell.fill = PatternFill("solid", fgColor="F2F2F2" if r_idx % 2 == 0 else "FFFFFF")
            ws.row_dimensions[row_num].height = 20

        # Column widths
        for col in range(1, 12):
            ws.column_dimensions[get_column_letter(col)].width = 18

        # =====================================================
        # SHEET 2: All Orders (جميع الطلبات)
        # =====================================================
        df_display.to_excel(writer, sheet_name='جميع الطلبات', index=False)
        ws2 = writer.sheets['جميع الطلبات']
        ws2.sheet_view.rightToLeft = True
        _style_data_sheet(ws2, df, header_fill, header_font, thin_border, center_align)

        # =====================================================
        # SHEET 3: Per Platform
        # =====================================================
        for platform in df['platform'].unique():
            p_df = df[df['platform'] == platform]
            p_display = p_df.rename(columns=col_names_ar)
            sheet_name = platform[:31]  # Excel sheet name limit
            p_display.to_excel(writer, sheet_name=sheet_name, index=False)
            ws_p = writer.sheets[sheet_name]
            ws_p.sheet_view.rightToLeft = True
            _style_data_sheet(ws_p, p_df, subheader_fill, header_font, thin_border, center_align)

        # =====================================================
        # SHEET 4: Pending & Discrepancies (متأخرات وفروقات)
        # =====================================================
        pending = df[df['status'] != 'مدفوع'].rename(columns=col_names_ar)
        pending.to_excel(writer, sheet_name='متأخرات وفروقات', index=False)
        ws4 = writer.sheets['متأخرات وفروقات']
        ws4.sheet_view.rightToLeft = True
        _style_data_sheet(ws4, df[df['status'] != 'مدفوع'],
                          PatternFill("solid", fgColor="C00000"), header_font, thin_border, center_align)

    print(f"✅ Report generated: {report_path}")
    print(f"\n{'='*50}")
    print(f"  إجمالي الطلبات:      {total_orders}")
    print(f"  إجمالي المبيعات:     {total_expected:,.2f} ر.س")
    print(f"  إجمالي المحصل:       {total_collected:,.2f} ر.س")
    print(f"  صافي الربح:          {total_net_profit:,.2f} ر.س")
    print(f"  نسبة التحصيل:        {collection_rate:.1f}%")
    print(f"{'='*50}")
    print(f"\nتفصيل حسب الحالة:")
    print(df.groupby(['platform', 'status']).size().unstack(fill_value=0))


def _style_data_sheet(ws, source_df, header_fill, header_font, thin_border, center_align):
    """Apply consistent styling to a data worksheet."""
    status_colors = {
        'مدفوع':              'E2EFDA',
        'غير مدفوع':          'FCE4D6',
        'مدفوع جزئياً':       'FFF2CC',
        'زيادة في التحصيل':   'D6E4F0',
    }
    
    # Style header row
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_align
        cell.border = thin_border
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = 'A2'

    # Style data rows based on status
    status_col_idx = None
    for i, cell in enumerate(ws[1]):
        if cell.value in ('الحالة', 'status'):
            status_col_idx = i + 1
            break

    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        status_val = ''
        if status_col_idx:
            status_val = ws.cell(row=row_idx, column=status_col_idx).value or ''
        
        bg = status_colors.get(status_val, 'FFFFFF')
        fill = PatternFill("solid", fgColor=bg)
        
        for cell in row:
            cell.border = thin_border
            cell.alignment = center_align
            if status_val:
                cell.fill = fill
        ws.row_dimensions[row_idx].height = 18

    # Auto-fit column widths
    for col in ws.columns:
        max_len = max((len(str(c.value or '')) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 35)


if __name__ == "__main__":
    generate_weekly_report()
