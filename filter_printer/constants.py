import os

# Limits
HASH_LIMIT_BYTES = 100 * 1024 * 1024  # 100MB

# Retries & Timeouts
COPY_RETRY_COUNT = 3
COPY_RETRY_DELAY = 1.0  # seconds

PRINT_RETRY_COUNT = 3
PRINT_RETRY_DELAY = 2.0  # seconds
SUMATRA_TIMEOUT = 60  # seconds

SCAN_UPDATE_INTERVAL = 50  # progress update every N files

# Scanner ignored directories
IGNORED_DIRS = {
    "node_modules",
    ".git",
    "__pycache__",
    "System Volume Information",
    "$Recycle.Bin"
}

# Config default values
DEFAULT_CONFIG = {
    "window_geometry": "880x600",
    "last_keyword": "",
    "last_active_types": {},
    "sort_column": "name",
    "sort_reverse": False,
    "last_folder": "",
    "last_printer": "",
    "last_destination_folder": ""
}
