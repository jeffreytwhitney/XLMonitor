import logging
import os
import sys

from dotenv import load_dotenv


def _get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))


BASE_DIR = _get_base_dir()

load_dotenv(os.path.join(BASE_DIR, '.env'))

LOGGER_LEVEL = os.getenv('LOGGER_LEVEL', 'INFO').upper()
LOGGER_FILE = os.getenv('LOGGER_FILE', 'app.log')
LOG_PATH = os.path.join(BASE_DIR, LOGGER_FILE)

logger = logging.getLogger('XLMonitor')
logger.setLevel(getattr(logging, LOGGER_LEVEL, logging.INFO))

if not logger.handlers:
    handler = logging.FileHandler(LOG_PATH, encoding='utf-8')
    handler.setFormatter(logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s'
    ))
    logger.addHandler(handler)
