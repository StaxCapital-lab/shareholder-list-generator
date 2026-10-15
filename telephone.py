# ============================================================
# FORCE UTF-8 OUTPUT (fixes UnicodeEncodeError when this script is
# launched via subprocess, e.g. from app3.py/app5.py). When output is
# piped/captured instead of going to a real console, Python falls back to
# the system codepage (e.g. cp1251 on Windows), which can't encode the
# emoji used in the print() messages below. Must run before any print().
# ============================================================
import sys
import io

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
else:
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import re
from openpyxl import load_workbook
import sys
import os
import random
import logging
import platform
import atexit

# Get the file path from the command-line argument
if len(sys.argv) < 2:
    print("Error: No file path provided!")
    print("Usage: python3 telephone_1881_batch_updated.py <excel_file_path>")
    sys.exit(1)

file_path = sys.argv[1]  # dynamically gets the correct file path from app.py
sheets = ["Private 2024", "Selskap 2024"]  # Sheets to process
# "Selskap 2024" is kept in the workbook for feature parity, but company
# shareholders are intentionally NOT searched on 1881. Their pages use a
# different layout/URL and this lookup is person-only.

# Setup logging for debugging. The log file is written next to the input
# Excel file (inside this run's own session folder) instead of a single
# shared "telephone_debug.log" in the program folder - otherwise several
# people running this on the server at the same time would overwrite each
# other's logs.
log_path = os.path.join(os.path.dirname(file_path) or '.', 'telephone_debug.log')
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Debug: Print file path and check if file exists
print(f"Received file path: {file_path}")
print(f"File exists: {os.path.exists(file_path)}")
if os.path.exists(file_path):
    print(f"File size: {os.path.getsize(file_path)} bytes")
    print(f"File extension: {os.path.splitext(file_path)[1]}")
else:
    print("ERROR: File does not exist!")
    sys.exit(1)

# ============================================================
# VIRTUAL DISPLAY (Xvfb) FOR SERVER DEPLOYMENT
# ============================================================
# This script never uses Chrome's --headless flag, because headless Chrome
# is much more likely to get flagged as a bot by sites like this one. On a
# server (no physical screen), we instead start a virtual display (Xvfb)
# and run Chrome normally against it - Chrome behaves exactly like a
# regular visible browser, it just isn't shown on a real monitor.
# Requires on the server: `sudo apt install xvfb` and
# `pip install pyvirtualdisplay`. Locally on Windows this block is skipped
# entirely, so a real Chrome window still opens, same as before.
virtual_display = None
used_headless_fallback = False

if platform.system() == "Linux":
    try:
        from pyvirtualdisplay import Display
        virtual_display = Display(visible=False, size=(1920, 1080))
        virtual_display.start()
        atexit.register(virtual_display.stop)
        print("🖥️ Virtual display (Xvfb) started for server environment")
    except Exception as e:
        print(f"⚠️ Could not start virtual display ({e}) - falling back to --headless=new")
        used_headless_fallback = True
