import os
import sys

from dotenv import load_dotenv

# Ensure the project root is importable as `XLMonitor` regardless of the
# directory pytest is invoked from.
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Load test-specific settings (.envtest) before XLMonitor/logger are imported,
# so the whole test suite - including the integration test that opens a real
# Excel file - runs against the tests/files fixture directories instead of
# whatever is configured in the production .env file. `override=True` ensures
# these values win even if .env is loaded afterwards (e.g. by logger.py).
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".envtest"), override=True)
