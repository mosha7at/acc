import logging
import os
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, ContextTypes, MessageHandler, filters, CallbackQueryHandler
from config import TELEGRAM_TOKEN, TEMPLATE_DIR, OUTPUT_DIR
from config import WELCOME_MESSAGE, HELP_MESSAGE, TEMPLATE_MESSAGE, UPLOAD_MESSAGE, PROCESSING_MESSAGE, SUCCESS_MESSAGE, ERROR_MESSAGE
from excel_processor import create_template, process_excel_file
from financial_statements import generate_financial_statements

# Enable logging
logger = logging.getLogger(__name__)

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /start is issued."""
    await update.message.reply_text(WELCOME_MESSAGE)
    # Add keyboard with options
    keyboard = [
        [InlineKeyboardButton("الحصول على القالب / Get Template", callback_data="template")],
        [InlineKeyboardButton("إنشاء القوائم المالية / Generate Statements", callback_data="generate")]
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text("اختر إحدى الخيارات: / Choose one option:", reply_markup=reply_markup)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send a message when the command /help is issued."""
    await update.message.reply_text(HELP_MESSAGE)

async def button_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Handle button callbacks."""
    query = update.callback_query
    await query.answer()
    if query.data == "template":
        await send_template(query, context)
    elif query.data == "generate":
        await query.edit_message_text(UPLOAD_MESSAGE)
        context.user_data["waiting_for_excel"] = True

async def send_template(update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """Send the Excel template to the user."""
    template_path = os.path.join(TEMPLATE_DIR, "financial_template.xlsx")
    if not os.path.exists(template_path):
        create_template(template_path)
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

def start_bot() -> None:
    """Start the bot."""
    application = Application.builder().token(TELEGRAM_TOKEN).build()
    application.add_handler(CommandHandler("start", start))
    application.add_handler(CommandHandler("help", help_command))
    application.add_handler(CommandHandler("template", template_command))
    application.add_handler(CommandHandler("generate", generate_command))
    application.add_handler(CallbackQueryHandler(button_callback))
    application.add_handler(MessageHandler(filters.ATTACHMENT, handle_document))
    application.run_polling()
    logger.info("Bot started")
