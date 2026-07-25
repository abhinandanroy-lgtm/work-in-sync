import json
import os
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LOG_DIR = os.path.join(BASE_DIR, "logs")
DEBUG_DIR = os.path.join(BASE_DIR, os.getenv("DEBUG_DIR", "debug_output"))
SCREENSHOT_DIR = os.path.join(DEBUG_DIR, "screenshots")
HTML_DIR = os.path.join(DEBUG_DIR, "html")
TEXT_DIR = os.path.join(DEBUG_DIR, "text")

os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(DEBUG_DIR, exist_ok=True)
os.makedirs(SCREENSHOT_DIR, exist_ok=True)
os.makedirs(HTML_DIR, exist_ok=True)
os.makedirs(TEXT_DIR, exist_ok=True)

WIS_URL = os.getenv("WIS_URL")
HEADLESS = os.getenv("HEADLESS", "false").lower() == "true"
TIMEOUT_MS = int(os.getenv("TIMEOUT_MS", "25000"))
SEARCH_WAIT_MS = int(os.getenv("SEARCH_WAIT_MS", "3000"))

INPUT_FILE = os.path.join(BASE_DIR, "employee_input.xlsx")
OUTPUT_FILE = os.path.join(BASE_DIR, "attendance_output.xlsx")
SELECTORS_FILE = os.path.join(BASE_DIR, "selectors.json")

EDGE_EXE = os.getenv("EDGE_EXE", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")
CDP_PORT = int(os.getenv("CDP_PORT", "9222"))

def load_selectors():
    with open(SELECTORS_FILE, "r", encoding="utf-8") as f:
        return json.load(f)