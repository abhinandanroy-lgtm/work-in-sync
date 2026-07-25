from datetime import datetime
from playwright.sync_api import sync_playwright
from app.config import TIMEOUT_MS, CDP_PORT
from app.logger import get_logger
from app.excel_io import read_employees, write_results
from app.wis import WorkInSyncBot

log = get_logger("main")

def get_active_page(browser):
    contexts = browser.contexts
    if not contexts:
        raise RuntimeError("No browser contexts found from CDP connection.")
    context = contexts[0]
    pages = context.pages
    if not pages:
        raise RuntimeError("No pages found in connected Edge browser.")
    return pages[-1]

def run():
    employees = read_employees()
    results = []

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
        page = get_active_page(browser)
        page.set_default_timeout(TIMEOUT_MS)

        bot = WorkInSyncBot(page)
        bot.goto_dashboard()

        if not bot.ensure_logged_in():
            log.error("Login session not active in opened Edge window.")
            return

        bot.open_team_calendar()

        for _, row in employees.iterrows():
            emp_id = row.get("employee_id", "")
            emp_name = row.get("employee_name", "")

            try:
                result = bot.process_employee(emp_id, emp_name)
                result["date"] = datetime.now().strftime("%Y-%m-%d")
                results.append(result)
                log.info(f"Processed: {result}")
            except Exception as e:
                error_row = {
                    "date": datetime.now().strftime("%Y-%m-%d"),
                    "employee_id": str(emp_id).strip(),
                    "employee_name": str(emp_name).strip(),
                    "checked_in": "Error",
                    "checked_out": "Error"
                }
                results.append(error_row)
                log.exception(f"Failed for {emp_id}/{emp_name}: {e}")

    write_results(results)
    log.info("attendance_output.xlsx saved successfully.")

if __name__ == "__main__":
    run()