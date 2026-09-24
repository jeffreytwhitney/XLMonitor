import os
import sys

from dotenv import load_dotenv, dotenv_values


def _get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(sys.argv[0]))


class EnvConfig:
    """Centralized, typed access to the application's environment variables.

    Settings that stay fixed for the life of the process (paths, intervals,
    TEST_MODE, ...) are parsed once and cached as attributes on construction.
    KILL_ME_NOW is exposed as a property so it is re-read live on every
    access, since it is polled repeatedly from XLMonitor's long-running loop.
    """

    def __init__(self):
        self.BASE_DIR = _get_base_dir()
        self.ENV_PATH = os.path.join(self.BASE_DIR, '.env')
        load_dotenv(self.ENV_PATH)

        self.LOGGER_LEVEL = self.get_str('LOGGER_LEVEL', 'INFO').upper()
        self.LOGGER_FILE = self.get_str('LOGGER_FILE', 'app.log')

        self.WATCH_DIR = self.get_str('WATCH_DIR')
        self.OUTPUT_1F_DIR = self.get_str('OUTPUT_1F_DIR')
        self.OUTPUT_CSV_DIR = self.get_str('OUTPUT_CSV_DIR')
        self.ARCHIVE_DIR = self.get_str('ARCHIVE_DIR')
        self.POLL_INTERVAL = self.get_int('POLL_INTERVAL', 5)
        self.MAX_ARCHIVE_FILE_AGE = self.get_int('MAX_ARCHIVE_FILE_AGE', 30)
        self.TRIM_INTERVAL = self.get_int('TRIM_INTERVAL', 86400)
        self.TEST_MODE = self.get_bool('TEST_MODE')

    @staticmethod
    def get_str(name, default=''):
        return os.getenv(name, default)

    @staticmethod
    def get_int(name, default=0):
        value = os.getenv(name)
        if value is None or value.strip() == '':
            return default
        try:
            return int(value)
        except ValueError:
            return default

    @staticmethod
    def get_bool(name, default=False):
        """Parse a value as a boolean.

        Recognizes '1', 'true', and 'yes' (case-insensitive) as True;
        anything else -- including the literal string 'False' -- is False.
        """
        value = os.getenv(name)
        if value is None:
            return default
        return value.strip().lower() in ('1', 'true', 'yes')

    @property
    def kill_me_now(self):
        """Live-read so KILL_ME_NOW can be toggled without restarting.

        os.environ is only populated once at startup by load_dotenv, so
        editing the .env file on disk while the app is running would
        never be reflected there. Re-parse the .env file directly on
        every access instead, falling back to os.environ (e.g. for
        variables set on the real process environment or via tests).
        """
        values = dotenv_values(self.ENV_PATH)
        value = values.get('KILL_ME_NOW')
        if value is None:
            value = os.getenv('KILL_ME_NOW')
        if value is None:
            return False
        return value.strip().lower() in ('1', 'true', 'yes')


config = EnvConfig()
