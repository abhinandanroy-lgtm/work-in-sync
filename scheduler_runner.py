import schedule
import subprocess
import time
import sys
import threading
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
SCRIPT_PATH = BASE_DIR / "setup_profile.py"
PYTHON_EXE = sys.executable

RUN_TIME = "15:50"

is_running = False
lock = threading.Lock()


def log(msg):
    print(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}")


def run_attendance_job():
    global is_running

    with lock:
        if is_running:
            log("Previous job is still running. Skipping this run.")
            return
        is_running = True

    try:
        log(f"Starting attendance automation: {SCRIPT_PATH}")
        result = subprocess.run(
            [PYTHON_EXE, str(SCRIPT_PATH)],
            cwd=str(BASE_DIR),
            capture_output=True,
            text=True
        )

        if result.stdout:
            print(result.stdout)

        if result.stderr:
            log("ERROR OUTPUT:")
            print(result.stderr)

        log(f"Attendance automation finished with return code {result.returncode}")

    except Exception as e:
        log(f"Scheduler error: {e}")

    finally:
        with lock:
            is_running = False


schedule.every().monday.at(RUN_TIME).do(run_attendance_job)
schedule.every().tuesday.at(RUN_TIME).do(run_attendance_job)
schedule.every().wednesday.at(RUN_TIME).do(run_attendance_job)
schedule.every().thursday.at(RUN_TIME).do(run_attendance_job)
schedule.every().friday.at(RUN_TIME).do(run_attendance_job)

log("Python scheduler started.")
log(f"Project folder: {BASE_DIR}")
log(f"Script to run: {SCRIPT_PATH}")
log(f"Scheduled time: Monday to Friday at {RUN_TIME}")

while True:
    try:
        schedule.run_pending()
        time.sleep(15)
    except KeyboardInterrupt:
        log("Scheduler stopped manually.")
        break
    except Exception as e:
        log(f"Main scheduler loop error: {e}")
        time.sleep(30)