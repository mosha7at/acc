import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Telegram Bot Token (get from BotFather)
TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN", "YOUR_TELEGRAM_TOKEN")

# File paths
TEMPLATE_DIR = "templates"
OUTPUT_DIR = "output"

# Ensure directories exist
os.makedirs(TEMPLATE_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Bot messages
WELCOME_MESSAGE = """
مرحباً بك في بوت القوائم المالية! 👋
هذا البوت يساعدك في إعداد القوائم المالية الخمسة تلقائياً.

استخدم الأوامر التالية:
/start - بدء استخدام البوت
/help - عرض المساعدة
/template - الحصول على قالب إكسل للتعبئة
/generate - رفع ملف إكسل لإنشاء القوائم المالية

Welcome to the Financial Statements Bot! 👋
This bot helps you prepare the five financial statements automatically.

Use the following commands:
/start - Start using the bot
/help - Display help
/template - Get Excel template to fill
/generate - Upload Excel file to generate financial statements
"""

HELP_MESSAGE = """
كيفية استخدام البوت:
1. استخدم الأمر /template للحصول على قالب إكسل
2. قم بتعبئة البيانات المالية في القالب
3. استخدم الأمر /generate وقم برفع ملف الإكسل المعبأ
4. انتظر حتى يتم إنشاء القوائم المالية وتحميلها

How to use the bot:
1. Use /template command to get the Excel template
2. Fill in the financial data in the template
3. Use /generate command and upload the filled Excel file
4. Wait until the financial statements are generated and downloaded
"""

TEMPLATE_MESSAGE = "يرجى استخدام هذا القالب لتعبئة البيانات المالية. / Please use this template to fill in the financial data."
UPLOAD_MESSAGE = "يرجى رفع ملف الإكسل المعبأ. / Please upload the filled Excel file."
PROCESSING_MESSAGE = "جاري معالجة البيانات... / Processing data..."
SUCCESS_MESSAGE = "تم إنشاء القوائم المالية بنجاح! / Financial statements have been successfully generated!"
ERROR_MESSAGE = "حدث خطأ أثناء معالجة البيانات. يرجى التأكد من صحة البيانات المدخلة. / An error occurred while processing data. Please make sure the entered data is correct."
