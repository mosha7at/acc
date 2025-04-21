import logging
from bot import start_bot

if __name__ == "__main__":
    # Set up logging
    logging.basicConfig(
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        level=logging.INFO
    )
    logger = logging.getLogger(__name__)
    logger.info("Starting the bot...")
    
    try:
        # Start the Telegram bot
        start_bot()
    except Exception as e:
        logger.error(f"Failed to start the bot: {str(e)}")
