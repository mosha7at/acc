import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart

def validate_data(data):
    """
    Validate the input data to ensure it contains all required keys and values.
    """
    required_keys = {
        'income': ['إجمالي الإيرادات | Total Revenue', 'صافي الربح | Net Profit'],
        'balance': ['إجمالي الأصول | Total Assets', 'إجمالي الخصوم | Total Liabilities', 'إجمالي حقوق الملكية | Total Equity'],
        'cash_flow': ['النقد وما في حكمه في نهاية السنة | Cash and cash equivalents at end of year']
    }
    errors = []
    for section, keys in required_keys.items():
        if section not in data:
            errors.append(f"المدخل '{section}' مفقود.")
            continue
        for key in keys:
            if key not in data[section]:
                errors.append(f"البند '{key}' مفقود في القسم '{section}'.")
            elif not isinstance(data[section][key].get('current', 0), (int, float)) or not isinstance(data[section][key].get('previous', 0), (int, float)):
                errors.append(f"القيم الخاصة بالبند '{key}' في القسم '{section}' ليست أرقامًا صالحة.")
    return errors

def generate_financial_statements(data, output_path):
    """Generate financial statements based on the provided data."""
    # Validate input data
    validation_errors = validate_data(data)
    if validation_errors:
        raise ValueError(f"أخطاء في البيانات المدخلة: {validation_errors}")
    
    wb = openpyxl.Workbook()
    # Create sheets for different financial statements
    sheets = {
        'تقرير عام | Overview': wb.active,
        'قائمة الدخل | Income Statement': wb.create_sheet(),
        'قائمة المركز المالي | Balance Sheet': wb.create_sheet(),
        'قائمة التغيرات في حقوق الملكية | Equity': wb.create_sheet(),
        'قائمة التدفقات النقدية | Cash Flow': wb.create_sheet(),
        'الملاحظات | Notes': wb.create_sheet(),
        'الرسوم البيانية | Charts': wb.create_sheet()
    }
    # Rename the default sheet
    sheets['تقرير عام | Overview'].title = 'تقرير عام | Overview'
    
    # Generate each statement
    generate_overview(sheets['تقرير عام | Overview'], data)
    generate_income_statement(sheets['قائمة الدخل | Income Statement'], data['income'])
    generate_balance_sheet(sheets['قائمة المركز المالي | Balance Sheet'], data['balance'])
    generate_equity_statement(sheets['قائمة التغيرات في حقوق الملكية | Equity'], data['equity'])
    generate_cash_flow_statement(sheets['قائمة التدفقات النقدية | Cash Flow'], data['cash_flow'])
    generate_notes(sheets['الملاحظات | Notes'], data.get('notes', {}))
    generate_charts(sheets['الرسوم البيانية | Charts'], data)
    
    # Save the workbook
    wb.save(output_path)
    return output_path

def generate_overview(sheet, data):
    """Generate an overview sheet with key financial metrics."""
    # Set up header
    sheet['A1'] = 'التقرير المالي الشامل | Comprehensive Financial Report'
    sheet['A1'].font = Font(bold=True, size=16)
    sheet['A3'] = 'المؤشرات المالية الرئيسية | Key Financial Indicators'
    sheet['A3'].font = Font(bold=True, size=14)
    
    # Format cells
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 20
    
    # Add headers
    sheet['A5'] = 'المؤشر | Indicator'
    sheet['B5'] = 'السنة الحالية | Current Year'
    sheet['C5'] = 'السنة السابقة | Previous Year'
    sheet['D5'] = 'التغيير٪ | Change%'
    format_header_row(sheet, row=5)
    
    # Extract key metrics from data
    try:
        metrics = [
            ('إجمالي الإيرادات | Total Revenue', 
             data['income']['إجمالي الإيرادات | Total Revenue']['current'], 
             data['income']['إجمالي الإيرادات | Total Revenue']['previous']),
            ('صافي الربح | Net Profit', 
             data['income']['صافي الربح | Net Profit']['current'], 
             data['income']['صافي الربح | Net Profit']['previous']),
            ('إجمالي الأصول | Total Assets', 
             data['balance']['إجمالي الأصول | Total Assets']['current'], 
             data['balance']['إجمالي الأصول | Total Assets']['previous']),
            ('إجمالي الخصوم | Total Liabilities', 
             data['balance']['إجمالي الخصوم | Total Liabilities']['current'], 
             data['balance']['إجمالي الخصوم | Total Liabilities']['previous']),
            ('إجمالي حقوق الملكية | Total Equity', 
             data['balance']['إجمالي حقوق الملكية | Total Equity']['current'], 
             data['balance']['إجمالي حقوق الملكية | Total Equity']['previous']),
        ]
        
        for i, (metric, current, previous) in enumerate(metrics, start=6):
            sheet[f'A{i}'] = metric
            sheet[f'B{i}'] = current
            sheet[f'C{i}'] = previous
            change_percent = ((current - previous) / previous * 100) if previous else 0
            sheet[f'D{i}'] = f"{change_percent:.2f}%"
            apply_conditional_formatting(sheet, cell=f'D{i}', value=change_percent)
    except Exception as e:
        sheet['A15'] = f"خطأ في حساب المؤشرات: {str(e)}"

def format_header_row(sheet, row):
    """Format a header row with bold font, blue background, and centered alignment."""
    for cell in sheet[row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')

def apply_conditional_formatting(sheet, cell, value):
    """Apply conditional formatting based on the value."""
    if value > 0:
        sheet[cell].fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    elif value < 0:
        sheet[cell].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# ... [Rest of the functions remain similar but include additional validations and optimizations]
