import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def create_template(output_path):
    """Create an Excel template for financial data input."""
    wb = openpyxl.Workbook()
    # Create sheets for different financial components
    sheets = {
        'تعليمات | Instructions': wb.active,
        'الإيرادات والمصروفات | Income': wb.create_sheet(),
        'الأصول والخصوم | Balance': wb.create_sheet(),
        'حقوق الملكية | Equity': wb.create_sheet(),
        'التدفقات النقدية | Cash Flow': wb.create_sheet(),
        'الملاحظات | Notes': wb.create_sheet()
    }
    # Rename the default sheet
    sheets['تعليمات | Instructions'].title = 'تعليمات | Instructions'
    
    # Set up Instructions sheet
    instructions = sheets['تعليمات | Instructions']
    instructions['A1'] = 'تعليمات استخدام القالب | Template Instructions'
    instructions['A1'].font = Font(bold=True, size=14)
    instructions['A3'] = 'مرحباً بكم في قالب القوائم المالية! | Welcome to the Financial Statements Template!'
    instructions['A5'] = '1. قم بتعبئة البيانات المالية في كل ورقة من أوراق هذا الملف.'
    instructions['A6'] = '1. Fill in the financial data in each sheet of this file.'
    instructions['A8'] = '2. تأكد من إدخال جميع المبالغ بالأرقام فقط (بدون رموز العملة).'
    instructions['A9'] = '2. Make sure to enter all amounts as numbers only (without currency symbols).'
    instructions['A11'] = '3. أكمل جميع الأوراق للحصول على قوائم مالية كاملة ودقيقة.'
    instructions['A12'] = '3. Complete all sheets to get complete and accurate financial statements.'
    instructions['A14'] = '4. بعد الانتهاء، احفظ الملف وقم برفعه باستخدام أمر /generate في البوت.'
    instructions['A15'] = '4. When finished, save the file and upload it using the /generate command in the bot.'
    
    # Format cells to appropriate width
    for col in range(1, 10):
        instructions.column_dimensions[get_column_letter(col)].width = 30
    
    # Set up other sheets
    setup_income_sheet(sheets['الإيرادات والمصروفات | Income'])
    setup_balance_sheet(sheets['الأصول والخصوم | Balance'])
    setup_equity_sheet(sheets['حقوق الملكية | Equity'])
    setup_cash_flow_sheet(sheets['التدفقات النقدية | Cash Flow'])
    setup_notes_sheet(sheets['الملاحظات | Notes'])
    
    # Save the workbook
    wb.save(output_path)
    return output_path

def setup_income_sheet(sheet):
    # Same as before...

def setup_balance_sheet(sheet):
    # Same as before...

def setup_equity_sheet(sheet):
    # Same as before...

def setup_cash_flow_sheet(sheet):
    # Same as before...

def setup_notes_sheet(sheet):
    # Same as before...

def process_excel_file(file_path):
    """Process the Excel file and extract financial data."""
    try:
        wb = openpyxl.load_workbook(file_path)
        # Extract data from each sheet
        income_data = extract_income_data(wb['الإيرادات والمصروفات | Income'])
        balance_data = extract_balance_data(wb['الأصول والخصوم | Balance'])
        equity_data = extract_equity_data(wb['حقوق الملكية | Equity'])
        cash_flow_data = extract_cash_flow_data(wb['التدفقات النقدية | Cash Flow'])
        notes_data = extract_notes_data(wb['الملاحظات | Notes'])
        
        # Collect errors
        errors = collect_errors({
            'income': income_data,
            'balance': balance_data,
            'equity': equity_data,
            'cash_flow': cash_flow_data,
            'notes': notes_data
        })
        
        return {
            'income': income_data,
            'balance': balance_data,
            'equity': equity_data,
            'cash_flow': cash_flow_data,
            'notes': notes_data,
            'errors': errors  # Add errors to the returned data
        }
    except Exception as e:
        raise Exception(f"Error processing Excel file: {str(e)}")

def extract_income_data(sheet):
    """Extract data from income statement sheet."""
    data = {}
    errors = []
    for row in range(4, 23):  # Adjust range based on your template
        item_name = sheet[f'A{row}'].value
        if item_name:
            current_year = sheet[f'B{row}'].value
            previous_year = sheet[f'C{row}'].value
            
            # Assume zero for missing values
            if current_year is None or current_year == "":
                current_year = 0
                errors.append(f"Income - Missing value for '{item_name}' (Current Year)")
            if previous_year is None or previous_year == "":
                previous_year = 0
                errors.append(f"Income - Missing value for '{item_name}' (Previous Year)")
            
            data[item_name] = {'current': current_year, 'previous': previous_year}
    return data

def extract_balance_data(sheet):
    """Extract data from balance sheet."""
    data = {}
    errors = []
    for row in range(4, 45):  # Adjust range based on your template
        item_name = sheet[f'A{row}'].value
        if item_name:
            current_year = sheet[f'B{row}'].value
            previous_year = sheet[f'C{row}'].value
            
            # Assume zero for missing values
            if current_year is None or current_year == "":
                current_year = 0
                errors.append(f"Balance - Missing value for '{item_name}' (Current Year)")
            if previous_year is None or previous_year == "":
                previous_year = 0
                errors.append(f"Balance - Missing value for '{item_name}' (Previous Year)")
            
            data[item_name] = {'current': current_year, 'previous': previous_year}
    return data

def extract_equity_data(sheet):
    """Extract data from equity statement sheet."""
    data = {}
    errors = []
    for row in range(4, 11):  # Adjust range based on your template
        item_name = sheet[f'A{row}'].value
        if item_name:
            capital = sheet[f'B{row}'].value
            reserves = sheet[f'C{row}'].value
            retained = sheet[f'D{row}'].value
            total = sheet[f'E{row}'].value
            
            # Assume zero for missing values
            if capital is None or capital == "":
                capital = 0
                errors.append(f"Equity - Missing value for '{item_name}' (Capital)")
            if reserves is None or reserves == "":
                reserves = 0
                errors.append(f"Equity - Missing value for '{item_name}' (Reserves)")
            if retained is None or retained == "":
                retained = 0
                errors.append(f"Equity - Missing value for '{item_name}' (Retained Earnings)")
            if total is None or total == "":
                total = 0
                errors.append(f"Equity - Missing value for '{item_name}' (Total)")
            
            data[item_name] = {
                'capital': capital,
                'reserves': reserves,
                'retained': retained,
                'total': total
            }
    return data

def extract_cash_flow_data(sheet):
    """Extract data from cash flow statement sheet."""
    data = {}
    errors = []
    for row in range(4, 31):  # Adjust range based on your template
        item_name = sheet[f'A{row}'].value
        if item_name:
            current_year = sheet[f'B{row}'].value
            previous_year = sheet[f'C{row}'].value
            
            # Assume zero for missing values
            if current_year is None or current_year == "":
                current_year = 0
                errors.append(f"Cash Flow - Missing value for '{item_name}' (Current Year)")
            if previous_year is None or previous_year == "":
                previous_year = 0
                errors.append(f"Cash Flow - Missing value for '{item_name}' (Previous Year)")
            
            data[item_name] = {'current': current_year, 'previous': previous_year}
    return data

def extract_notes_data(sheet):
    """Extract notes data."""
    notes = {}
    errors = []
    note_rows = [
        (4, 'note1'),  # General Information
        (8, 'note2'),  # Basis of Preparation
        (12, 'note3'),  # Significant Accounting Policies
        (16, 'note4'),  # Judgments and Estimates
        (20, 'note5'),  # Financial Risk Management
        (24, 'note6'),  # Additional Information
        (28, 'note7')   # Subsequent Events
    ]
    for row, note_key in note_rows:
        if sheet[f'B{row}'].value:
            notes[note_key] = sheet[f'B{row}'].value
        else:
            notes[note_key] = ""
            errors.append(f"Notes - Missing information for '{note_key}'")
    return notes

def collect_errors(data):
    """Collect all errors and missing values from the extracted data."""
    errors = []
    for section, items in data.items():
        if isinstance(items, dict):
            for item, values in items.items():
                if isinstance(values, dict):
                    for key, value in values.items():
                        if value is None or value == "":
                            errors.append(f"{section} - Missing value for '{item}' ({key})")
    return errors
