import logging
import os
from telegram import Update, ReplyKeyboardMarkup
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters
from config import TELEGRAM_TOKEN, TEMPLATE_DIR, OUTPUT_DIR
from config import WELCOME_MESSAGE, HELP_MESSAGE, TEMPLATE_MESSAGE, UPLOAD_MESSAGE, PROCESSING_MESSAGE, SUCCESS_MESSAGE, ERROR_MESSAGE
from excel_processor import create_template, process_excel_file
from financial_statements import generate_financial_statements

# Enable logging
logger = logging.getLogger(__name__)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /start is issued."""
    user_name = update.message.from_user.first_name  # جلب الاسم الأول للمستخدم
    welcome_message = f"مرحبا بك يا {user_name} في بوت المحاسب! 👋\n\nكيف يمكنني مساعدتك اليوم؟"
    
    # إنشاء لوحة المفاتيح
    keyboard = [
        ["الحصول على القالب", "إنشاء القوائم المالية"],
        ["تعليمات", "مساعدة"]  # زران جديدان بجانب بعض
    ]
    reply_markup = ReplyKeyboardMarkup(keyboard, resize_keyboard=True)
    
    await update.message.reply_text(welcome_message, reply_markup=reply_markup)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /help or 'مساعدة' button is pressed."""
    help_message = """
كيفية استخدام البوت:
1. اضغط على "الحصول على القالب" للحصول على قالب إكسل.
2. قم بتعبئة البيانات المالية في القالب.
3. اضغط على "إنشاء القوائم المالية" وقم برفع ملف الإكسل المعبأ.
4. انتظر حتى يتم إنشاء القوائم المالية وتحميلها.

How to use the bot:
1. Press "Get Template" to obtain the Excel template.
2. Fill in the financial data in the template.
3. Press "Generate Financial Statements" and upload the filled Excel file.
4. Wait until the financial statements are generated and downloaded.
"""
    await update.message.reply_text(help_message)

async def instructions_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle the 'تعليمات' button."""
    instructions_message = """
هذه هي التعليمات الأساسية لاستخدام البوت:
- استخدم الزر "الحصول على القالب" للحصول على قالب إكسل.
- استخدم الزر "إنشاء القوائم المالية" لرفع ملف إكسل وإنشاء القوائم المالية.
- إذا كنت بحاجة إلى المساعدة، اضغط على زر "مساعدة".
"""
    await update.message.reply_text(instructions_message)

async def send_template(update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send the Excel template to the user."""
    template_path = os.path.join(TEMPLATE_DIR, "financial_template.xlsx")
    if not os.path.exists(template_path):
        create_template(template_path)
    await update.message.reply_text(TEMPLATE_MESSAGE)
    await update.message.reply_document(document=open(template_path, 'rb'))

async def generate_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Wait for Excel file upload when the command /generate is issued."""
    await update.message.reply_text(UPLOAD_MESSAGE)
    context.user_data["waiting_for_excel"] = True

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle Excel file uploads."""
    if not context.user_data.get("waiting_for_excel", False):
        return
    context.user_data["waiting_for_excel"] = False
    file = update.message.document
    file_name = file.file_name
    if not file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("يرجى رفع ملف إكسل فقط. / Please upload only Excel files.")
        return
    await update.message.reply_text(PROCESSING_MESSAGE)
    new_file = await context.bot.get_file(file.file_id)
    input_path = os.path.join(TEMPLATE_DIR, f"input_{update.message.chat_id}.xlsx")
    await new_file.download_to_drive(input_path)
    try:
        data = process_excel_file(input_path)
        output_path = os.path.join(OUTPUT_DIR, f"financial_statements_{update.message.chat_id}.xlsx")
        generate_financial_statements(data, output_path)
        await update.message.reply_text(SUCCESS_MESSAGE)
        await update.message.reply_document(document=open(output_path, 'rb'))
        try:
            os.remove(input_path)
            os.remove(output_path)
        except:
            pass
    except Exception as e:
        logger.error(f"Error processing file: {e}")
        await update.message.reply_text(f"{ERROR_MESSAGE}\nError details: {str(e)}")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle text messages from custom keyboard buttons."""
    text = update.message.text
    
    if text == "الحصول على القالب":
        await send_template(update, context)
    elif text == "إنشاء القوائم المالية":
        await generate_command(update, context)
    elif text == "تعليمات":
        await instructions_command(update, context)
    elif text == "مساعدة":
        await help_command(update, context)

def start_bot() -> None:
    """Start the bot."""
    application = Application.builder().token(TELEGRAM_TOKEN).build()
    
    # Add command handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    
    # Add message handler for custom keyboard buttons
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    
    # Start the Bot
    application.run_polling()
    logger.info("Bot started")
