import logging
from bot import start_bot

if __name__ == "__main__":
    # Set up logging
    logging.basicConfig(
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )
    # Start the Telegram bot
    start_bot()
