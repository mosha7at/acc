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
        'الأخطاء | Errors': wb.create_sheet()  # Add the errors sheet
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
    generate_errors_sheet(sheets['الأخطاء | Errors'], data['errors'])  # Generate the errors sheet
    
    # Save the workbook
    wb.save(output_path)
    return output_path

def generate_overview(sheet, data):
    """Generate an overview sheet with key financial metrics."""
    # Same as before...

def generate_income_statement(sheet, income_data):
    """Generate income statement."""
    # Same as before...

def generate_balance_sheet(sheet, balance_data):
    """Generate balance sheet."""
    # Same as before...

def generate_equity_statement(sheet, equity_data):
    """Generate statement of changes in equity."""
    # Same as before...

def generate_cash_flow_statement(sheet, cash_flow_data):
    """Generate cash flow statement."""
    # Same as before...

def generate_notes(sheet, notes_data):
    """Generate notes to financial statements."""
    # Same as before...

def generate_charts(sheet, data):
    """Generate financial charts."""
    # Same as before...

def generate_errors_sheet(sheet, errors):
    """Generate a sheet to display all errors and missing values."""
    # Set up header
    sheet['A1'] = 'الأخطاء والقيم المفقودة | Errors and Missing Values'
    sheet['A1'].font = Font(bold=True, size=16)
    sheet['A3'] = 'الوصف | Description'
    sheet['B3'] = 'القسم | Section'
    sheet['C3'] = 'البند | Item'
    sheet['D3'] = 'السنة | Year'
    # Format header row
    for cell in sheet['3:3']:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    
    # Set column widths
    sheet.column_dimensions['A'].width = 50
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 25
    sheet.column_dimensions['D'].width = 15
    
    # Add errors to the sheet
    row = 4
    for error in errors:
        description, section, item, year = parse_error(error)
        sheet[f'A{row}'] = description
        sheet[f'B{row}'] = section
        sheet[f'C{row}'] = item
        sheet[f'D{row}'] = year
        row += 1

def parse_error(error_message):
    """Parse an error message into its components."""
    # Example error format: "Income - Missing value for 'إيرادات المبيعات' (Current Year)"
    parts = error_message.split(" - ")
    section = parts[0]
    details = parts[1].split(" for '")
    description = details[0]
    item_year = details[1].strip(")").split("' (")
    item = item_year[0]
    year = item_year[1]
    return description, section, item, year
