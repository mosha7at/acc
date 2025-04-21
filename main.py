import logging
from bot import start_bot

if __name__ == "__main__":
    # Set up logging
    logging.basicConfig(
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        level=logging.WARNING  # تغيير مستوى التسجيل إلى WARNING لتقليل الرسائل غير الضرورية
    )
    
    # تعطيل تسجيل الطلبات الناجحة من مكتبة httpx
    logging.getLogger("httpx").setLevel(logging.WARNING)
    
    # Start the Telegram bot
    start_bot()
