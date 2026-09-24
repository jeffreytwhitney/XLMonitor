import logging
import os

from env_config import config

BASE_DIR = config.BASE_DIR
LOG_PATH = os.path.join(BASE_DIR, config.LOGGER_FILE)

logger = logging.getLogger('XLMonitor')
logger.setLevel(getattr(logging, config.LOGGER_LEVEL, logging.INFO))

if not logger.handlers:
    handler = logging.FileHandler(LOG_PATH, encoding='utf-8')
    handler.setFormatter(logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
    ))
    logger.addHandler(handler)
