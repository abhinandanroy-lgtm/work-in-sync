import time
from app.config import WIS_URL, TIMEOUT_MS, SEARCH_WAIT_MS, load_selectors
from app.logger import get_logger
from app.debug_tools import save_all

log = get_logger("wis")

class WorkInSyncBot:
    def __init__(self, page):
        self.page = page
        self.sel = load_selectors()
        self.page.set_default_timeout(TIMEOUT_MS)

    def goto_dashboard(self):
        self.page.goto(WIS_URL, wait_until="load")
        time.sleep(3)
        save_all(self.page, "bot_01_dashboard_loaded")

    def ensure_logged_in(self):
        marker = self.sel.get("logged_in_marker", "").strip()
        if not marker:
            return True
        try:
            self.page.locator(marker).wait_for(timeout=10000)
            save_all(self.page, "bot_02_logged_in_confirmed")
            return True
        except Exception:
            save_all(self.page, "bot_02_logged_in_not_confirmed")
            return False

    def open_team_calendar(self):
        selector = self.sel.get("view_team_calendar", "").strip()
        if not selector:
            return
        try:
            loc = self.page.locator(selector)
            if loc.is_visible(timeout=4000):
                save_all(self.page, "bot_03_before_click_view_team_calendar")
                loc.click()
                time.sleep(3)
                save_all(self.page, "bot_04_after_click_view_team_calendar")
        except Exception as e:
            log.info(f"Team Calendar click skipped: {e}")
            save_all(self.page, "bot_04_team_calendar_click_skipped")

    def clear_search(self):
        selector = self.sel.get("clear_all", "").strip()
        if not selector:
            return
        try:
            loc = self.page.locator(selector)
            if loc.is_visible(timeout=2000):
                save_all(self.page, "bot_05_before_clear_search")
                loc.click()
                time.sleep(1)
                save_all(self.page, "bot_06_after_clear_search")
        except Exception:
            pass

    def search_employee(self, search_value):
        selector = self.sel.get("search_input", "").strip()
        if not selector:
            raise ValueError("search_input selector missing in selectors.json")

        box = self.page.locator(selector)
        save_all(self.page, f"bot_07_before_search_{search_value}")
        box.click()
        box.fill("")
        box.fill(str(search_value))
        time.sleep(SEARCH_WAIT_MS / 1000)
        save_all(self.page, f"bot_08_after_search_{search_value}")
        log.info(f"Searched employee: {search_value}")

    def get_full_page_text(self):
        try:
            txt = self.page.locator("body").inner_text().upper()
            return txt
        except Exception:
            return ""

    def extract_status(self):
        txt = self.get_full_page_text()
        checked_in = "Yes" if "CHECKED IN" in txt else "No"
        checked_out = "Yes" if "CHECKED OUT" in txt else "No"
        return checked_in, checked_out

    def process_employee(self, employee_id, employee_name):
        self.clear_search()
        search_value = str(employee_id).strip() if str(employee_id).strip() else str(employee_name).strip()
        self.search_employee(search_value)
        checked_in, checked_out = self.extract_status()

        return {
            "employee_id": str(employee_id).strip(),
            "employee_name": str(employee_name).strip(),
            "checked_in": checked_in,
            "checked_out": checked_out
        }