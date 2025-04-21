import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# تعريف ألوان التنسيق
YELLOW_FILL = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
ERROR_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

def validate_and_correct_data(data):
    """
    Validate the input data and correct missing or invalid values.
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
                data[section][key] = {'current': 0, 'previous': 0}  # إضافة القيمة كصفر
            elif not isinstance(data[section][key].get('current', 0), (int, float)) or not isinstance(data[section][key].get('previous', 0), (int, float)):
                errors.append(f"القيم الخاصة بالبند '{key}' في القسم '{section}' ليست أرقامًا صالحة.")
                data[section][key] = {'current': 0, 'previous': 0}  # إضافة القيمة كصفر
    return data, errors

def highlight_cell(sheet, cell, value, fill=YELLOW_FILL):
    """
    Highlight a cell with a specific color and set its value.
    """
    sheet[cell] = value
    sheet[cell].fill = fill

def generate_financial_statements(data, output_path):
    """Generate financial statements based on the provided data."""
    # Validate and correct data
    data, validation_errors = validate_and_correct_data(data)
    
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
    generate_overview(sheets['تقرير عام | Overview'], data, validation_errors)
    generate_income_statement(sheets['قائمة الدخل | Income Statement'], data['income'])
    generate_balance_sheet(sheets['قائمة المركز المالي | Balance Sheet'], data['balance'])
    generate_equity_statement(sheets['قائمة التغيرات في حقوق الملكية | Equity'], data['equity'])
    generate_cash_flow_statement(sheets['قائمة التدفقات النقدية | Cash Flow'], data['cash_flow'])
    generate_notes(sheets['الملاحظات | Notes'], data.get('notes', {}))
    generate_charts(sheets['الرسوم البيانية | Charts'], data)
    
    # Save the workbook
    wb.save(output_path)
    return output_path

def generate_overview(sheet, data, validation_errors):
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
            if current == 0:
                highlight_cell(sheet, f'B{i}', current)
            else:
                sheet[f'B{i}'] = current
            if previous == 0:
                highlight_cell(sheet, f'C{i}', previous)
            else:
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

def generate_income_statement(sheet, income_data):
    """Generate income statement."""
    sheet['A1'] = 'قائمة الدخل | Income Statement'
    sheet['A1'].font = Font(bold=True, size=16)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'السنة الحالية | Current Year'
    sheet['C3'] = 'السنة السابقة | Previous Year'
    sheet['D3'] = 'التغيير | Change'
    sheet['E3'] = 'التغيير٪ | Change%'
    format_header_row(sheet, row=3)
    
    row = 4
    for item, values in income_data.items():
        sheet[f'A{row}'] = item
        current = values.get('current', 0)
        previous = values.get('previous', 0)
        if current == 0:
            highlight_cell(sheet, f'B{row}', current)
        else:
            sheet[f'B{row}'] = current
        if previous == 0:
            highlight_cell(sheet, f'C{row}', previous)
        else:
            sheet[f'C{row}'] = previous
        change = current - previous
        sheet[f'D{row}'] = change
        change_percent = (change / previous * 100) if previous else 0
        sheet[f'E{row}'] = f"{change_percent:.2f}%"
        if "إجمالي" in item or "صافي" in item:
            for col in ['A', 'B', 'C', 'D', 'E']:
                sheet[f'{col}{row}'].font = Font(bold=True)
                sheet[f'{col}{row}'].fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        row += 1

def generate_balance_sheet(sheet, balance_data):
    """Generate balance sheet."""
    sheet['A1'] = 'قائمة المركز المالي | Balance Sheet'
    sheet['A1'].font = Font(bold=True, size=16)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'السنة الحالية | Current Year'
    sheet['C3'] = 'السنة السابقة | Previous Year'
    sheet['D3'] = 'التغيير | Change'
    sheet['E3'] = 'التغيير٪ | Change%'
    format_header_row(sheet, row=3)
    
    row = 4
    for item, values in balance_data.items():
        sheet[f'A{row}'] = item
        current = values.get('current', 0)
        previous = values.get('previous', 0)
        if current == 0:
            highlight_cell(sheet, f'B{row}', current)
        else:
            sheet[f'B{row}'] = current
        if previous == 0:
            highlight_cell(sheet, f'C{row}', previous)
        else:
            sheet[f'C{row}'] = previous
        change = current - previous
        sheet[f'D{row}'] = change
        change_percent = (change / previous * 100) if previous else 0
        sheet[f'E{row}'] = f"{change_percent:.2f}%"
        if "إجمالي" in item:
            for col in ['A', 'B', 'C', 'D', 'E']:
                sheet[f'{col}{row}'].font = Font(bold=True)
                sheet[f'{col}{row}'].fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        row += 1
    
    # Validate balance sheet (Assets = Liabilities + Equity)
    assets = balance_data.get('إجمالي الأصول | Total Assets', {}).get('current', 0)
    liab_equity = balance_data.get('إجمالي الخصوم وحقوق الملكية | Total Liabilities and Equity', {}).get('current', 0)
    sheet[f'A{row+2}'] = 'التحقق من توازن قائمة المركز المالي | Balance Sheet Check'
    sheet[f'A{row+2}'].font = Font(bold=True)
    if abs(assets - liab_equity) < 0.01:
        sheet[f'B{row+2}'] = 'متوازن ✓ | Balanced ✓'
        sheet[f'B{row+2}'].fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    else:
        sheet[f'B{row+2}'] = f'غير متوازن ✗ | Not Balanced ✗ (فرق | Difference: {assets - liab_equity})'
        sheet[f'B{row+2}'].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

# Example usage
if __name__ == "__main__":
    data = {
        'income': {
            'إجمالي الإيرادات | Total Revenue': {'current': 2522753, 'previous': 323283},
            'صافي الربح | Net Profit': {'current': -756822, 'previous': 541875}
        },
        'balance': {
            'إجمالي الأصول | Total Assets': {'current': 80611, 'previous': 636667},
            'إجمالي الخصوم | Total Liabilities': {'current': 946969, 'previous': 101156},
            'إجمالي حقوق الملكية | Total Equity': {'current': 247002, 'previous': 250298}
        },
        'cash_flow': {
            'النقد وما في حكمه في نهاية السنة | Cash and cash equivalents at end of year': {'current': 84519, 'previous': 386251}
        },
        'notes': {}
    }
    output_path = "financial_statements_corrected.xlsx"
    generate_financial_statements(data, output_path)
