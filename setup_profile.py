import os
import subprocess
import time
import requests
import pandas as pd
from playwright.sync_api import sync_playwright

WIS_URL = "https://pwc.moveinsync.com/WP/employee.jsp#WorkInSyncDashboard"
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CDP_PORT = 9222
TIMEOUT_MS = 30000
SEARCH_WAIT_SEC = 4

INPUT_FILE = "employee_input.xlsx"
OUTPUT_FILE = "attendance_output.xlsx"

DEBUG_DIR = "debug_output"
SCREENSHOT_DIR = os.path.join(DEBUG_DIR, "screenshots")
HTML_DIR = os.path.join(DEBUG_DIR, "html")
TEXT_DIR = os.path.join(DEBUG_DIR, "text")

os.makedirs(SCREENSHOT_DIR, exist_ok=True)
os.makedirs(HTML_DIR, exist_ok=True)
os.makedirs(TEXT_DIR, exist_ok=True)


def ts():
    return time.strftime("%Y%m%d_%H%M%S")


def save_step(page, name):
    try:
        page.screenshot(
            path=os.path.join(SCREENSHOT_DIR, f"{ts()}_{name}.png"),
            full_page=True,
            timeout=8000
        )
    except:
        pass

    try:
        with open(os.path.join(HTML_DIR, f"{ts()}_{name}.html"), "w", encoding="utf-8") as f:
            f.write(page.content())
    except:
        pass

    try:
        txt = page.locator("body").inner_text(timeout=5000)
        with open(os.path.join(TEXT_DIR, f"{ts()}_{name}.txt"), "w", encoding="utf-8") as f:
            f.write(txt)
    except:
        pass


def wait_for_cdp(port=9222, retries=30, delay=1):
    url = f"http://127.0.0.1:{port}/json/version"
    for _ in range(retries):
        try:
            r = requests.get(url, timeout=2)
            if r.status_code == 200:
                return True
        except:
            pass
        time.sleep(delay)
    return False


def get_live_moveinsync_page(browser, retries=20, delay=2):
    for _ in range(retries):
        for context in browser.contexts:
            for page in context.pages:
                try:
                    if not page.is_closed() and "moveinsync.com" in page.url:
                        return page
                except:
                    pass
        time.sleep(delay)
    raise RuntimeError("Could not find live MoveInSync page")


def text_visible(page, text_value, timeout=5000):
    try:
        page.get_by_text(text_value, exact=False).wait_for(timeout=timeout)
        return True
    except:
        return False


def click_text(page, text_value, step_name):
    save_step(page, f"{step_name}_before")
    loc = page.get_by_text(text_value, exact=False).first
    loc.click(timeout=10000)
    time.sleep(4)
    save_step(page, f"{step_name}_after")


def get_wis_frame_locator(page):
    page.locator("iframe").first.wait_for(timeout=20000)
    return page.frame_locator("iframe").first


def team_calendar_opened(page):
    frame = get_wis_frame_locator(page)

    checks = [
        lambda: frame.get_by_text("Clear All", exact=False).first.wait_for(timeout=3000),
        lambda: frame.locator("input[placeholder*='Search']").first.wait_for(timeout=3000),
        lambda: frame.locator("input").first.wait_for(timeout=3000)
    ]

    for check in checks:
        try:
            check()
            return True
        except:
            pass

    return False


def open_team_calendar_in_frame(page):
    save_step(page, "04_before_team_calendar_attempt")
    frame = get_wis_frame_locator(page)

    # Widget anchored by exact ID from provided HTML
    widget = frame.locator("div.widget-content-wrapper").filter(
        has=frame.locator("#widget-name-Team\\ Calendar")
    ).first

    # Fallback widget by text
    widget_by_text = frame.locator("mis-employee-web-exp-widget-card").filter(
        has_text="Team Calendar"
    ).first

    # 1. Exact button inside Team Calendar widget
    try:
        btn = widget.locator("button:has-text('View Team Calendar')").first
        btn.wait_for(timeout=10000)
        btn.scroll_into_view_if_needed(timeout=5000)
        time.sleep(1)
        btn.click(timeout=8000, force=True)
        time.sleep(6)
        save_step(page, "05_team_calendar_button_in_widget_clicked")
        if team_calendar_opened(page):
            return True
    except:
        pass

    # 2. Button inside fallback widget by text
    try:
        btn = widget_by_text.locator("button:has-text('View Team Calendar')").first
        btn.wait_for(timeout=8000)
        btn.scroll_into_view_if_needed(timeout=5000)
        time.sleep(1)
        btn.click(timeout=8000, force=True)
        time.sleep(6)
        save_step(page, "05_team_calendar_button_in_fallback_widget_clicked")
        if team_calendar_opened(page):
            return True
    except:
        pass

    # 3. Click .inside container from provided HTML
    try:
        inside = widget.locator("div.inside").first
        inside.wait_for(timeout=8000)
        inside.scroll_into_view_if_needed(timeout=5000)
        time.sleep(1)
        inside.click(timeout=8000, force=True)
        time.sleep(6)
        save_step(page, "05_team_calendar_inside_clicked")
        if team_calendar_opened(page):
            return True
    except:
        pass

    # 4. Click calendar wrapper
    try:
        cal = widget.locator("div.calender-wrapper").first
        cal.wait_for(timeout=8000)
        cal.scroll_into_view_if_needed(timeout=5000)
        time.sleep(1)
        cal.click(timeout=8000, force=True)
        time.sleep(6)
        save_step(page, "05_team_calendar_wrapper_clicked")
        if team_calendar_opened(page):
            return True
    except:
        pass

    # 5. Click the widget card itself
    try:
        widget.wait_for(timeout=8000)
        widget.scroll_into_view_if_needed(timeout=5000)
        time.sleep(1)
        widget.click(timeout=8000, force=True)
        time.sleep(6)
        save_step(page, "05_team_calendar_widget_clicked")
        if team_calendar_opened(page):
            return True
    except:
        pass

    # 6. JS click exact button text inside Team Calendar widget
    try:
        clicked = frame.locator("mis-employee-web-exp-widget-card").evaluate("""
        (nodes) => {
            const list = Array.isArray(nodes) ? nodes : [nodes];
            for (const root of list) {
                if (!root || !root.innerText || !root.innerText.includes('Team Calendar')) continue;
                const btns = root.querySelectorAll('button');
                for (const b of btns) {
                    if ((b.innerText || '').trim() === 'View Team Calendar') {
                        b.click();
                        return true;
                    }
                }
            }
            return false;
        }
        """)
        if clicked:
            time.sleep(6)
            save_step(page, "05_team_calendar_js_widget_button_click")
            if team_calendar_opened(page):
                return True
    except:
        pass

    save_step(page, "05_team_calendar_all_attempts_failed")
    return False


def focus_search_in_frame(page):
    frame = get_wis_frame_locator(page)

    candidates = [
        frame.get_by_placeholder("Search by name, ID or email"),
        frame.get_by_placeholder("Search"),
        frame.locator("input[placeholder*='Search by name']").first,
        frame.locator("input[placeholder*='name']").first,
        frame.locator("input[placeholder*='email']").first,
        frame.locator("input[type='search']").first,
        frame.locator("input[type='text']").first,
        frame.locator("input").first
    ]

    for c in candidates:
        try:
            c.wait_for(timeout=5000)
            c.scroll_into_view_if_needed(timeout=3000)
            c.click(timeout=5000, force=True)
            return c
        except:
            pass

    raise RuntimeError("Search field not found inside iframe")


def clear_and_type(locator, value):
    locator.click(timeout=5000)
    time.sleep(0.5)

    try:
        locator.fill("")
        locator.fill(str(value))
        return
    except:
        pass

    try:
        locator.press("Control+A")
        locator.press("Backspace")
        locator.type(str(value), delay=100)
        return
    except:
        pass

    raise RuntimeError("Could not type into search field")


def extract_status(page):
    frame = get_wis_frame_locator(page)

    try:
        txt = frame.locator("body").inner_text(timeout=5000).upper()
    except:
        txt = ""

    checked_in = "Yes" if "CHECKED IN" in txt else "No"
    checked_out = "Yes" if "CHECKED OUT" in txt else "No"
    return checked_in, checked_out


def main():
    temp_profile = os.path.join(os.getcwd(), "edge_debug_profile")
    os.makedirs(temp_profile, exist_ok=True)

    cmd = [
        EDGE_EXE,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={temp_profile}",
        WIS_URL
    ]
    subprocess.Popen(cmd)

    if not wait_for_cdp(CDP_PORT, retries=30, delay=1):
        print("CDP did not start")
        return

    employees = pd.read_excel(INPUT_FILE).fillna("")
    results = []

    with sync_playwright() as p:
        browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
        page = get_live_moveinsync_page(browser)
        page.set_default_timeout(TIMEOUT_MS)

        try:
            page.goto(WIS_URL, wait_until="domcontentloaded")
        except:
            pass

        time.sleep(5)
        page = get_live_moveinsync_page(browser)
        save_step(page, "01_initial_page")

        if text_visible(page, "Pick an account", 5000):
            click_text(page, "abhinandan.roy@pwc.com", "02_pick_account")

        time.sleep(8)
        page = get_live_moveinsync_page(browser)
        save_step(page, "03_after_login")

        opened = open_team_calendar_in_frame(page)

        if not opened:
            print("Auto-click for Team Calendar failed.")
            input("Open Team Calendar manually, then press Enter here... ")
            save_step(page, "06_team_calendar_opened_manually")
        else:
            save_step(page, "06_team_calendar_opened_auto")

        try:
            search_box = focus_search_in_frame(page)
            save_step(page, "07_search_focused")
        except Exception:
            print("Could not auto-focus search inside iframe.")
            input("Click inside search field manually, then press Enter here... ")
            search_box = focus_search_in_frame(page)
            save_step(page, "07_search_manual_focus_done")

        for _, row in employees.iterrows():
            emp_id = str(row.get("employee_id", "")).strip()
            emp_name = str(row.get("employee_name", "")).strip()
            search_value = emp_id if emp_id else emp_name

            try:
                frame = get_wis_frame_locator(page)

                try:
                    clear_all = frame.get_by_text("Clear All", exact=False).first
                    clear_all.click(timeout=3000, force=True)
                    time.sleep(1)
                except:
                    pass

                search_box = focus_search_in_frame(page)
                save_step(page, f"08_before_search_{search_value}")
                clear_and_type(search_box, search_value)
                time.sleep(SEARCH_WAIT_SEC)
                save_step(page, f"09_after_search_{search_value}")

                checked_in, checked_out = extract_status(page)

                results.append({
                    "date": time.strftime("%Y-%m-%d"),
                    "employee_id": emp_id,
                    "employee_name": emp_name,
                    "checked_in": checked_in,
                    "checked_out": checked_out
                })

            except Exception:
                save_step(page, f"10_search_failed_{search_value}")
                results.append({
                    "date": time.strftime("%Y-%m-%d"),
                    "employee_id": emp_id,
                    "employee_name": emp_name,
                    "checked_in": "Error",
                    "checked_out": "Error"
                })

        pd.DataFrame(results).to_excel(OUTPUT_FILE, index=False)
        save_step(page, "11_finished")
        print(f"Done. Output saved to {OUTPUT_FILE}")
        input("Press Enter to close...")


if __name__ == "__main__":
    main()