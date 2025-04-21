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
    # Set up header
    sheet['A1'] = 'قائمة الدخل | Income Statement'
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'المبلغ (السنة الحالية) | Amount (Current Year)'
    sheet['C3'] = 'المبلغ (السنة السابقة) | Amount (Previous Year)'
    # Format header row
    for cell in sheet['3:3']:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    # Set up income items
    income_items = [
        'الإيرادات | Revenues',
        'إيرادات المبيعات | Sales Revenue',
        'إيرادات الخدمات | Services Revenue',
        'إيرادات أخرى | Other Revenue',
        'إجمالي الإيرادات | Total Revenue',
        '',
        'المصروفات | Expenses',
        'تكلفة البضاعة المباعة | Cost of Goods Sold',
        'مصروفات الرواتب | Salary Expenses',
        'مصروفات الإيجار | Rent Expenses',
        'مصروفات المرافق | Utility Expenses',
        'مصروفات التسويق | Marketing Expenses',
        'الاستهلاك والإطفاء | Depreciation & Amortization',
        'مصروفات أخرى | Other Expenses',
        'إجمالي المصروفات | Total Expenses',
        '',
        'الربح قبل الضرائب | Profit Before Tax',
        'ضريبة الدخل | Income Tax',
        'صافي الربح | Net Profit'
    ]
    for i, item in enumerate(income_items, start=4):
        sheet[f'A{i}'] = item
        if item.startswith('إجمالي') or item.startswith('صافي') or item.startswith('الربح'):
            sheet[f'A{i}'].font = Font(bold=True)
    # Format columns width
    sheet.column_dimensions['A'].width = 35
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 25

def setup_balance_sheet(sheet):
    # Set up header
    sheet['A1'] = 'قائمة المركز المالي | Balance Sheet'
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'المبلغ (السنة الحالية) | Amount (Current Year)'
    sheet['C3'] = 'المبلغ (السنة السابقة) | Amount (Previous Year)'
    # Format header row
    for cell in sheet['3:3']:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    # Set up assets
    assets = [
        'الأصول | Assets',
        'الأصول المتداولة | Current Assets',
        'النقدية وما في حكمها | Cash and Cash Equivalents',
        'الذمم المدينة | Accounts Receivable',
        'المخزون | Inventory',
        'أصول متداولة أخرى | Other Current Assets',
        'إجمالي الأصول المتداولة | Total Current Assets',
        '',
        'الأصول غير المتداولة | Non-Current Assets',
        'الممتلكات والمعدات | Property and Equipment',
        'الأصول غير الملموسة | Intangible Assets',
        'استثمارات طويلة الأجل | Long-term Investments',
        'أصول غير متداولة أخرى | Other Non-Current Assets',
        'إجمالي الأصول غير المتداولة | Total Non-Current Assets',
        '',
        'إجمالي الأصول | Total Assets',
        '',
        'الخصوم وحقوق الملكية | Liabilities and Equity',
        'الخصوم المتداولة | Current Liabilities',
        'الذمم الدائنة | Accounts Payable',
        'القروض قصيرة الأجل | Short-term Loans',
        'الإيرادات المؤجلة | Deferred Revenue',
        'خصوم متداولة أخرى | Other Current Liabilities',
        'إجمالي الخصوم المتداولة | Total Current Liabilities',
        '',
        'الخصوم غير المتداولة | Non-Current Liabilities',
        'القروض طويلة الأجل | Long-term Loans',
        'مخصص مكافأة نهاية الخدمة | End of Service Benefits',
        'خصوم غير متداولة أخرى | Other Non-Current Liabilities',
        'إجمالي الخصوم غير المتداولة | Total Non-Current Liabilities',
        '',
        'إجمالي الخصوم | Total Liabilities',
        '',
        'حقوق الملكية | Equity',
        'رأس المال | Capital',
        'الاحتياطيات | Reserves',
        'الأرباح المحتجزة | Retained Earnings',
        'إجمالي حقوق الملكية | Total Equity',
        '',
        'إجمالي الخصوم وحقوق الملكية | Total Liabilities and Equity'
    ]
    for i, item in enumerate(assets, start=4):
        sheet[f'A{i}'] = item
        if item.startswith('إجمالي') or item == 'الأصول | Assets' or item == 'الخصوم وحقوق الملكية | Liabilities and Equity':
            sheet[f'A{i}'].font = Font(bold=True)
    # Format columns width
    sheet.column_dimensions['A'].width = 35
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 25

def setup_equity_sheet(sheet):
    # Set up header
    sheet['A1'] = 'قائمة التغيرات في حقوق الملكية | Statement of Changes in Equity'
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'رأس المال | Capital'
    sheet['C3'] = 'الاحتياطيات | Reserves'
    sheet['D3'] = 'الأرباح المحتجزة | Retained Earnings'
    sheet['E3'] = 'الإجمالي | Total'
    # Format header row
    for cell in sheet['3:3']:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    # Set up equity items
    equity_items = [
        'الرصيد في بداية السنة | Balance at beginning of year',
        'صافي الربح للسنة | Net profit for the year',
        'توزيعات الأرباح | Dividends',
        'زيادة رأس المال | Capital increase',
        'المحول للاحتياطيات | Transferred to reserves',
        'تغييرات أخرى | Other changes',
        'الرصيد في نهاية السنة | Balance at end of year'
    ]
    for i, item in enumerate(equity_items, start=4):
        sheet[f'A{i}'] = item
        if item.startswith('الرصيد في'):
            sheet[f'A{i}'].font = Font(bold=True)
    # Format columns width
    sheet.column_dimensions['A'].width = 35
    sheet.column_dimensions['B'].width = 20
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 20
    sheet.column_dimensions['E'].width = 20

def setup_cash_flow_sheet(sheet):
    # Set up header
    sheet['A1'] = 'قائمة التدفقات النقدية | Cash Flow Statement'
    sheet['A1'].font = Font(bold=True, size=14)
    sheet['A3'] = 'البند | Item'
    sheet['B3'] = 'المبلغ (السنة الحالية) | Amount (Current Year)'
    sheet['C3'] = 'المبلغ (السنة السابقة) | Amount (Previous Year)'
    # Format header row
    for cell in sheet['3:3']:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
        cell.alignment = Alignment(horizontal='center')
    # Set up cash flow items
    cash_flow_items = [
        'التدفقات النقدية من الأنشطة التشغيلية | Cash flows from operating activities',
        'صافي الربح | Net profit',
        'تعديلات لـ: | Adjustments for:',
        'الاستهلاك والإطفاء | Depreciation and amortization',
        'التغير في الذمم المدينة | Change in accounts receivable',
        'التغير في المخزون | Change in inventory',
        'التغير في الذمم الدائنة | Change in accounts payable',
        'تعديلات أخرى | Other adjustments',
        'صافي النقد من الأنشطة التشغيلية | Net cash from operating activities',
        '',
        'التدفقات النقدية من الأنشطة الاستثمارية | Cash flows from investing activities',
        'شراء ممتلكات ومعدات | Purchase of property and equipment',
        'بيع ممتلكات ومعدات | Sale of property and equipment',
        'استثمارات جديدة | New investments',
        'بيع استثمارات | Sale of investments',
        'صافي النقد من الأنشطة الاستثمارية | Net cash from investing activities',
        '',
        'التدفقات النقدية من الأنشطة التمويلية | Cash flows from financing activities',
        'توزيعات أرباح مدفوعة | Dividends paid',
        'قروض جديدة | New loans',
        'سداد قروض | Loan repayments',
        'زيادة رأس المال | Capital increase',
        'صافي النقد من الأنشطة التمويلية | Net cash from financing activities',
        '',
        'صافي التغير في النقد وما في حكمه | Net change in cash and cash equivalents',
        'النقد وما في حكمه في بداية السنة | Cash and cash equivalents at beginning of year',
        'النقد وما في حكمه في نهاية السنة | Cash and cash equivalents at end of year'
    ]
    for i, item in enumerate(cash_flow_items, start=4):
        sheet[f'A{i}'] = item
        if item.startswith('صافي النقد') or item.startswith('التدفقات النقدية') or item == 'النقد وما في حكمه في نهاية السنة | Cash and cash equivalents at end of year':
            sheet[f'A{i}'].font = Font(bold=True)
    # Format columns width
    sheet.column_dimensions['A'].width = 45
    sheet.column_dimensions['B'].width = 25
    sheet.column_dimensions['C'].width = 25

def setup_notes_sheet(sheet):
    # Set up header
    sheet['A1'] = 'الملاحظات على القوائم المالية | Notes to Financial Statements'
    sheet['A1'].font = Font(bold=True, size=14)
    # Set up sections
    notes_sections = [
        ('A3', 'ملاحظة 1: معلومات عامة | Note 1: General Information'),
        ('A7', 'ملاحظة 2: أسس الإعداد | Note 2: Basis of Preparation'),
        ('A11', 'ملاحظة 3: السياسات المحاسبية الهامة | Note 3: Significant Accounting Policies'),
        ('A15', 'ملاحظة 4: الأحكام والتقديرات المحاسبية الهامة | Note 4: Significant Accounting Judgments and Estimates'),
        ('A19', 'ملاحظة 5: إدارة المخاطر المالية | Note 5: Financial Risk Management'),
        ('A23', 'ملاحظة 6: معلومات إضافية حول بنود القوائم المالية | Note 6: Additional Information on Financial Statement Items'),
        ('A27', 'ملاحظة 7: أحداث لاحقة | Note 7: Subsequent Events')
    ]
    for cell_ref, title in notes_sections:
        sheet[cell_ref] = title
        sheet[cell_ref].font = Font(bold=True)
    # Format columns width
    sheet.column_dimensions['A'].width = 50
    sheet.column_dimensions['B'].width = 50
    # Add instructions
    sheet['A4'] = 'أدخل وصفًا موجزًا للمنشأة وطبيعة أنشطتها الرئيسية. | Enter a brief description of the entity and the nature of its main activities.'
    sheet['A8'] = 'اذكر المعايير المحاسبية المتبعة وأساس القياس. | Mention the accounting standards followed and the measurement basis.'
    sheet['A12'] = 'اشرح السياسات المحاسبية الرئيسية المطبقة. | Explain the main accounting policies applied.'
    sheet['A16'] = 'اذكر الأحكام والتقديرات الهامة المستخدمة. | Mention significant judgments and estimates used.'
    sheet['A20'] = 'وضح كيفية إدارة المخاطر المالية مثل مخاطر الائتمان والسيولة والسوق. | Explain how financial risks such as credit, liquidity, and market risks are managed.'
    sheet['A24'] = 'قدم تفاصيل إضافية عن البنود الهامة في القوائم المالية. | Provide additional details about important items in the financial statements.'
    sheet['A28'] = 'اذكر أي أحداث هامة وقعت بعد تاريخ التقرير. | Mention any significant events that occurred after the reporting date.'

def process_excel_file(file_path):
    """Process the Excel file and extract financial data."""
    try:
        wb = openpyxl.load_workbook(file_path)
        # Extract data from each sheet
        data = {
            'income': extract_income_data(wb['الإيرادات والمصروفات | Income']),
            'balance': extract_balance_data(wb['الأصول والخصوم | Balance']),
            'equity': extract_equity_data(wb['حقوق الملكية | Equity']),
            'cash_flow': extract_cash_flow_data(wb['التدفقات النقدية | Cash Flow']),
            'notes': extract_notes_data(wb['الملاحظات | Notes'])
        }
        # Collect errors
        errors = collect_errors(data)
        data['errors'] = errors  # Add errors to the returned data
        return data
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
