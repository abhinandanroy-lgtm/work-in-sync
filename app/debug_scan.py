import time
from playwright.sync_api import sync_playwright
from app.config import WIS_URL, TIMEOUT_MS, SEARCH_WAIT_MS, load_selectors, CDP_PORT
from app.logger import get_logger
from app.debug_tools import save_all, save_all_frames
from app.excel_io import read_employees

log = get_logger("debug_scan")

def safe_step_dump(page, name):
    save_all(page, name)
    save_all_frames(page, name)

def get_active_page(browser):
    contexts = browser.contexts
    if not contexts:
        raise RuntimeError("No browser contexts found from CDP connection.")
    context = contexts[0]
    pages = context.pages
    if not pages:
        raise RuntimeError("No pages found in connected Edge browser.")
    return pages[-1]

def run_debug_scan():
    selectors = load_selectors()
    employees = read_employees()

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
        page = get_active_page(browser)
        page.set_default_timeout(TIMEOUT_MS)

        page.goto(WIS_URL, wait_until="load")
        time.sleep(5)
        safe_step_dump(page, "10_dashboard_loaded")

        marker = selectors.get("logged_in_marker", "").strip()
        if marker:
            try:
                page.locator(marker).wait_for(timeout=10000)
                log.info("Logged in marker found.")
                safe_step_dump(page, "11_logged_in_marker_found")
            except Exception:
                log.warning("Logged in marker not found.")
                safe_step_dump(page, "11_logged_in_marker_not_found")

        view_team_calendar = selectors.get("view_team_calendar", "").strip()
        if view_team_calendar:
            try:
                loc = page.locator(view_team_calendar)
                if loc.is_visible(timeout=5000):
                    safe_step_dump(page, "12_before_click_view_team_calendar")
                    loc.click()
                    time.sleep(4)
                    safe_step_dump(page, "13_after_click_view_team_calendar")
            except Exception as e:
                log.warning(f"Could not click view team calendar: {e}")
                safe_step_dump(page, "13_view_team_calendar_click_failed")

        search_input = selectors.get("search_input", "").strip()
        clear_all = selectors.get("clear_all", "").strip()

        if search_input:
            try:
                box = page.locator(search_input)
                safe_step_dump(page, "14_before_click_search_input")
                box.click()
                time.sleep(1)
                safe_step_dump(page, "15_after_click_search_input")
            except Exception as e:
                log.warning(f"Could not click search input: {e}")
                safe_step_dump(page, "15_search_input_click_failed")

        for idx, row in employees.iterrows():
            emp_id = str(row.get("employee_id", "")).strip()
            emp_name = str(row.get("employee_name", "")).strip()
            search_value = emp_id if emp_id else emp_name
            step_tag = f"employee_{idx+1}_{search_value}"

            try:
                if clear_all:
                    try:
                        clr = page.locator(clear_all)
                        if clr.is_visible(timeout=2000):
                            safe_step_dump(page, f"20_{step_tag}_before_clear_all")
                            clr.click()
                            time.sleep(2)
                            safe_step_dump(page, f"21_{step_tag}_after_clear_all")
                    except Exception:
                        pass

                box = page.locator(search_input)
                safe_step_dump(page, f"22_{step_tag}_before_fill_search")
                box.click()
                box.fill("")
                box.fill(search_value)
                time.sleep(SEARCH_WAIT_MS / 1000)
                safe_step_dump(page, f"23_{step_tag}_after_fill_search")

            except Exception as e:
                log.warning(f"Failed during search for {search_value}: {e}")
                safe_step_dump(page, f"24_{step_tag}_search_failed")

        input("Debug scan completed. Press Enter to finish... ")