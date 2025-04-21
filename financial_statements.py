import os
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, PieChart, LineChart, Series

def generate_financial_statements(data, output_path):
    """Generate financial statements based on the provided data."""
    wb = openpyxl.Workbook()
    # Create sheets for different financial statements
    sheets = {
        'تقرير عام | Overview': wb.active,
        'قائمة الدخل | Income Statement': wb.create_sheet(),
        'قائمة المركز المالي | Balance Sheet': wb.create_sheet(),
        'قائمة التغيرات في حقوق الملكية | Equity': wb.create_sheet(),
        'قائمة التدفقات النقدية | Cash Flow': wb.create_sheet(),
        'الملاحظات | Notes': wb.create_sheet(),
        'الرسوم البيانية | Charts': wb.create_sheet(),
        'الأخطاء | Errors': wb.create_sheet()  # Add a new sheet for errors
    }
    # Rename the default sheet
    sheets['تقرير عام | Overview'].title = 'تقرير عام | Overview'
    
    # Generate each statement
    generate_overview(sheets['تقرير عام | Overview'], data)
    generate_income_statement(sheets['قائمة الدخل | Income Statement'], data['income'])
    generate_balance_sheet(sheets['قائمة المركز المالي | Balance Sheet'], data['balance'])
    generate_equity_statement(sheets['قائمة التغيرات في حقوق الملكية | Equity'], data['equity'])
    generate_cash_flow_statement(sheets['قائمة التدفقات النقدية | Cash Flow'], data['cash_flow'])
    generate_notes(sheets['الملاحظات | Notes'], data['notes'])
    generate_charts(sheets['الرسوم البيانية | Charts'], data)
    generate_errors_sheet(sheets['الأخطاء | Errors'], data)  # Generate errors sheet
    
    # Save the workbook
    wb.save(output_path)
    return output_path

def generate_errors_sheet(sheet, data):
    """Generate a sheet to collect all errors and missing values."""
    # Set up header
    sheet['A1'] = 'الأخطاء والقيم المفقودة | Errors and Missing Values'
    sheet['A1'].font = Font(bold=True, size=16)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'الوصف | Description'
    sheet['C3'] = 'القيمة المفترضة | Assumed Value'
    
    # Format header row
    for cell in sheet['3:3']:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    
    # Set column widths
    sheet.column_dimensions['A'].width = 50
    sheet.column_dimensions['B'].width = 50
    sheet.column_dimensions['C'].width = 20
    
    # Collect errors from different sections
    row = 4
    for section, items in data.items():
        for item, values in items.items():
            if isinstance(values, dict):  # Check if the item has sub-values
                for key, value in values.items():
                    if value is None or value == "":
                        sheet[f'A{row}'] = f"{section} - {item} ({key})"
                        sheet[f'B{row}'] = "لم يتم إدخال القيمة | Value not entered"
                        sheet[f'C{row}'] = 0  # Assume zero for missing values
                        row += 1
            elif values is None or values == "":
                sheet[f'A{row}'] = f"{section} - {item}"
                sheet[f'B{row}'] = "لم يتم إدخال القيمة | Value not entered"
                sheet[f'C{row}'] = 0  # Assume zero for missing values
                row += 1

# Rest of the functions remain unchanged...

def generate_income_statement(sheet, income_data):
    """Generate income statement."""
    # Set up header
    sheet['A1'] = 'قائمة الدخل | Income Statement'
    sheet['A1'].font = Font(bold=True, size=16)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'السنة الحالية | Current Year'
    sheet['C3'] = 'السنة السابقة | Previous Year'
    sheet['D3'] = 'التغيير | Change'
    sheet['E3'] = 'التغيير٪ | Change%'
    
    # Format header row
    for cell in sheet['3:3']:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal='center')
    
    # Set column widths
    sheet.column_dimensions['A'].width = 40
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 20
    
    # Add income items
    row = 4
    for item, values in income_data.items():
        sheet[f'A{row}'] = item
        sheet[f'B{row}'] = values.get('current', 0)  # Assume zero for missing values
        sheet[f'C{row}'] = values.get('previous', 0)  # Assume zero for missing values
        
        # Calculate change
        current = values.get('current', 0)
        previous = values.get('previous', 0)
        change = current - previous
        sheet[f'D{row}'] = change
        
        # Calculate percentage change
        if previous != 0:
            change_percent = (change / previous) * 100
            sheet[f'E{row}'] = f"{change_percent:.2f}%"
        else:
            sheet[f'E{row}'] = "N/A"
        
        # Format totals and net profit
        if "إجمالي" in item or "صافي" in item or "الربح" in item:
            sheet[f'A{row}'].font = Font(bold=True)
            sheet[f'B{row}'].font = Font(bold=True)
            sheet[f'C{row}'].font = Font(bold=True)
            sheet[f'D{row}'].font = Font(bold=True)
            sheet[f'E{row}'].font = Font(bold=True)
            for col in ['A', 'B', 'C', 'D', 'E']:
                sheet[f'{col}{row}'].fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        
        row += 1

# Similar modifications should be applied to other functions like generate_balance_sheet, 
# generate_equity_statement, and generate_cash_flow_statement to handle missing values.

# Rest of the code remains unchanged...
