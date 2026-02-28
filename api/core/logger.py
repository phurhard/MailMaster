import logging
from logging.handlers import RotatingFileHandler
import os

# Create a logs directory if it doesn't exist
LOGS_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), "logs")
os.makedirs(LOGS_DIR, exist_ok=True)

LOG_FILE_PATH = os.path.join(LOGS_DIR, "mailmaster.log")

def setup_logger(name: str) -> logging.Logger:
    """
    Sets up a rich logger with a rotating file handler that keeps the latest 10 copies
    at 10MB each, along with stream out to the console.
    """
    logger = logging.getLogger(name)
    
    # Only configure if it hasn't been configured yet
    if not logger.handlers:
        logger.setLevel(logging.INFO)

        # Log format
        formatter = logging.Formatter(
            '%(asctime)s - %(name)s - [%(levelname)s] - %(message)s'
        )

        # Console Handler
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

        # Rotating File Handler (10MB per file, maintain 10 backups)
        file_handler = RotatingFileHandler(
            LOG_FILE_PATH,
            maxBytes=10 * 1024 * 1024,
            backupCount=10
        )
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    return logger
