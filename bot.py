import logging
import os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters, CallbackQueryHandler
from config import TELEGRAM_TOKEN, TEMPLATE_DIR, OUTPUT_DIR
from config import HELP_MESSAGE, TEMPLATE_MESSAGE, UPLOAD_MESSAGE, PROCESSING_MESSAGE, SUCCESS_MESSAGE, ERROR_MESSAGE
from excel_processor import create_template, process_excel_file
from financial_statements import generate_financial_statements

# Enable logging
logger = logging.getLogger(__name__)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /start is issued."""
    user_name = update.message.from_user.first_name  # جلب الاسم الأول للمستخدم
    welcome_message = f"مرحبا بك يا {user_name} في بوت المحاسب! 👋\n\nكيف يمكنني مساعدتك اليوم؟"
    
    # إنشاء الأزرار
    keyboard = [
        [InlineKeyboardButton("الحصول على القالب", callback_data="template")],
        [InlineKeyboardButton("إنشاء القوائم المالية", callback_data="generate")],
        [InlineKeyboardButton("مساعدة", callback_data="help")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    
    await update.message.reply_text(welcome_message, reply_markup=reply_markup)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /help is issued."""
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
    if isinstance(update, Update):  # إذا كان الطلب من رسالة نصية
        await update.message.reply_text(help_message)
    else:  # إذا كان الطلب من زر
        await update.edit_message_text(help_message)

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle button callbacks."""
    query = update.callback_query
    await query.answer()
    
    if query.data == "template":
        await send_template(query, context)
    elif query.data == "generate":
        await query.edit_message_text(UPLOAD_MESSAGE)
        context.user_data["waiting_for_excel"] = True
    elif query.data == "help":
        await help_command(query, context)  # استدعاء دالة المساعدة

async def send_template(update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send the Excel template to the user."""
    # Create a template if it doesn't exist
    template_path = os.path.join(TEMPLATE_DIR, "financial_template.xlsx")
    if not os.path.exists(template_path):
        create_template(template_path)
    # Send the template file
    if isinstance(update, Update):
        await update.message.reply_text(TEMPLATE_MESSAGE)
        await update.message.reply_document(document=open(template_path, 'rb'))
    else:  # Called from button callback
        await update.edit_message_text(TEMPLATE_MESSAGE)
        await context.bot.send_document(
            chat_id=update.from_user.id,
            document=open(template_path, 'rb')
        )

async def template_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send the Excel template when the command /template is issued."""
    await send_template(update, context)

async def generate_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Wait for Excel file upload when the command /generate is issued."""
    await update.message.reply_text(UPLOAD_MESSAGE)
    context.user_data["waiting_for_excel"] = True

async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle Excel file uploads."""
    # Check if we're waiting for an Excel file
    if not context.user_data.get("waiting_for_excel", False):
        return
    # Reset waiting state
    context.user_data["waiting_for_excel"] = False
    # Get file info
    file = update.message.document
    file_name = file.file_name
    # Check if it's an Excel file
    if not file_name.endswith(('.xlsx', '.xls')):
        await update.message.reply_text("يرجى رفع ملف إكسل فقط. / Please upload only Excel files.")
        return
    # Download the file
    await update.message.reply_text(PROCESSING_MESSAGE)
    new_file = await context.bot.get_file(file.file_id)
    input_path = os.path.join(TEMPLATE_DIR, f"input_{update.message.chat_id}.xlsx")
    await new_file.download_to_drive(input_path)
    try:
        # Process the Excel file
        data = process_excel_file(input_path)
        # Generate financial statements
        output_path = os.path.join(OUTPUT_DIR, f"financial_statements_{update.message.chat_id}.xlsx")
        generate_financial_statements(data, output_path)
        # Send the result back to the user
        await update.message.reply_text(SUCCESS_MESSAGE)
        await update.message.reply_document(document=open(output_path, 'rb'))
        # Clean up
        try:
            os.remove(input_path)
            os.remove(output_path)
        except:
            pass
    except Exception as e:
        logger.error(f"Error processing file: {e}")
        await update.message.reply_text(f"{ERROR_MESSAGE}\nError details: {str(e)}")

def start_bot() -> None:
    """Start the bot."""
    # Create the Application
    application = Application.builder().token(TELEGRAM_TOKEN).build()
    # Add command handlers
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("template", template_command))
    application.add_handler(CommandHandler("generate", generate_command))
    # Add callback query handler for inline buttons
    application.add_handler(CallbackQueryHandler(button_callback))
    # Add document handler
    application.add_handler(MessageHandler(filters.ATTACHMENT, handle_document))
    # Start the Bot
    application.run_polling()
    logger.info("Bot started")
