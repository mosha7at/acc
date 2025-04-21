import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

def validate_data(data):
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
    return errors

def generate_income_statement(sheet, income_data):
    """Generate income statement."""
    sheet['A1'] = 'قائمة الدخل | Income Statement'
    sheet['A1'].font = Font(bold=True, size=16)
    # تفاصيل إضافية للدالة...

def generate_financial_statements(data, output_path):
    """Generate financial statements based on the provided data."""
    validation_errors = validate_data(data)
    if validation_errors:
        raise ValueError(f"أخطاء في البيانات المدخلة: {validation_errors}")
    
    wb = openpyxl.Workbook()
    sheets = {
        'قائمة الدخل | Income Statement': wb.create_sheet(),
        # باقي الأوراق...
    }
    generate_income_statement(sheets['قائمة الدخل | Income Statement'], data['income'])
    # استدعاء باقي الدوال...
    wb.save(output_path)
    return output_path

# بيانات اختبارية
data = {
    'income': {
        'إجمالي الإيرادات | Total Revenue': {'current': 100000, 'previous': 90000},
        'صافي الربح | Net Profit': {'current': 20000, 'previous': 18000}
    },
    'balance': {
        'إجمالي الأصول | Total Assets': {'current': 500000, 'previous': 480000},
        'إجمالي الخصوم | Total Liabilities': {'current': 200000, 'previous': 190000},
        'إجمالي حقوق الملكية | Total Equity': {'current': 300000, 'previous': 290000}
    },
    'cash_flow': {
        'النقد وما في حكمه في نهاية السنة | Cash and cash equivalents at end of year': {'current': 50000, 'previous': 45000}
    }
}

# تشغيل الكود
output_path = "financial_statements.xlsx"
generate_financial_statements(data, output_path)
