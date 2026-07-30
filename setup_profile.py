import os
import subprocess
import time
from datetime import datetime
import requests
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font

from playwright.sync_api import sync_playwright

WIS_URL = "https://pwc.moveinsync.com/WP/employee.jsp#WorkInSyncDashboard"
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CDP_PORT = 9222
TIMEOUT_MS = 30000
SEARCH_WAIT_SEC = 2

INPUT_FILE = "employee_input.xlsx"
OUTPUT_FILE = "attendance_output.xlsx"
TRACKER_FILE = "Tracker.xlsx"

ABHINANDAN_ID = "101675341"
ABHINANDAN_NAME = "Abhinandan Roy"

DEBUG_DIR = "debug_output"
SCREENSHOT_DIR = os.path.join(DEBUG_DIR, "screenshots")
HTML_DIR = os.path.join(DEBUG_DIR, "html")
TEXT_DIR = os.path.join(DEBUG_DIR, "text")

os.makedirs(SCREENSHOT_DIR, exist_ok=True)
os.makedirs(HTML_DIR, exist_ok=True)
os.makedirs(TEXT_DIR, exist_ok=True)

VIEW_TEAM_CALENDAR_XPATH = "xpath=//*[@class='mis-btn mis-none mis-btn-md' and text()='View Team Calendar']"
SEARCH_XPATH = 'xpath=//*[@placeholder="Search by name, ID or email "]'
CLEAR_ALL_XPATH = "xpath=//*[text()='Clear All']"

ABHINANDAN_STATUS_XPATH = "xpath=(//div[@class='t-row ng-star-inserted'])[1]//div[contains(@class,'desktop_checkedin')]"
OTHERS_STATUS_XPATH = "xpath=(//div[@class='t-row ng-star-inserted'])[2]//div[contains(@class,'desktop_checkedin')]"

GREEN_FILL = PatternFill(fill_type="solid", start_color="92D050", end_color="92D050")
HEADER_FILL = PatternFill(fill_type="solid", start_color="BFBFBF", end_color="BFBFBF")
HEADER_FONT = Font(bold=True, size=12)


def ts():
    return time.strftime("%Y%m%d_%H%M%S")


def sheet_base_name():
    return time.strftime("%d-%m-%y")


def make_unique_sheet_name(output_file):
    base = sheet_base_name()

    if not os.path.exists(output_file):
        return f"{base}(1)"

    wb = load_workbook(output_file)
    existing = wb.sheetnames
    wb.close()

    i = 1
    while True:
        candidate = f"{base}({i})"
        if candidate not in existing:
            return candidate
        i += 1


def write_results_to_excel(results, output_file):
    df = pd.DataFrame(results)
    sheet_name = make_unique_sheet_name(output_file)

    if not os.path.exists(output_file):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        ws.append(list(df.columns))
        for row in df.itertuples(index=False, name=None):
            ws.append(list(row))
        wb.save(output_file)
        wb.close()
        return sheet_name

    wb = load_workbook(output_file)
    ws = wb.create_sheet(title=sheet_name)
    ws.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws.append(list(row))
    wb.save(output_file)
    wb.close()
    return sheet_name


def get_tracker_sheet_name(dt):
    return f"Tracker {dt.strftime('%B')}"


def apply_tracker_header_style(ws):
    for col_idx in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL


def ensure_tracker_headers(ws, day_col_name):
    headers = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]

    if ws.max_row == 1 and all(h is None for h in headers):
        ws.cell(row=1, column=1).value = "id"
        ws.cell(row=1, column=2).value = "name"
        ws.cell(row=1, column=3).value = day_col_name
        ws.cell(row=1, column=4).value = "Total"
        apply_tracker_header_style(ws)
        return

    headers = [ws.cell(row=1, column=i).value for i in range(1, ws.max_column + 1)]
    headers = [str(h).strip() if h is not None else "" for h in headers]

    if "id" not in headers:
        ws.insert_cols(1)
        ws.cell(row=1, column=1).value = "id"
        headers = [str(ws.cell(row=1, column=i).value).strip() if ws.cell(row=1, column=i).value is not None else "" for i in range(1, ws.max_column + 1)]

    if "name" not in headers:
        if len(headers) < 2:
            ws.insert_cols(2)
        ws.cell(row=1, column=2).value = "name"
        headers = [str(ws.cell(row=1, column=i).value).strip() if ws.cell(row=1, column=i).value is not None else "" for i in range(1, ws.max_column + 1)]

    headers = [str(ws.cell(row=1, column=i).value).strip() if ws.cell(row=1, column=i).value is not None else "" for i in range(1, ws.max_column + 1)]

    if day_col_name not in headers:
        total_col = None
        for i, h in enumerate(headers, start=1):
            if h == "Total":
                total_col = i
                break

        if total_col is not None:
            ws.insert_cols(total_col)
            ws.cell(row=1, column=total_col).value = day_col_name
        else:
            ws.cell(row=1, column=ws.max_column + 1).value = day_col_name

    headers = [str(ws.cell(row=1, column=i).value).strip() if ws.cell(row=1, column=i).value is not None else "" for i in range(1, ws.max_column + 1)]

    if "Total" not in headers:
        ws.cell(row=1, column=ws.max_column + 1).value = "Total"

    apply_tracker_header_style(ws)


def get_header_map(ws):
    return {
        str(ws.cell(row=1, column=i).value).strip(): i
        for i in range(1, ws.max_column + 1)
        if ws.cell(row=1, column=i).value is not None
    }


def find_or_create_tracker_row(ws, emp_id, emp_name, header_map):
    id_col = header_map["id"]
    name_col = header_map["name"]

    for row_idx in range(2, ws.max_row + 1):
        existing_id = ws.cell(row=row_idx, column=id_col).value
        if str(existing_id).strip() == str(emp_id).strip():
            if not ws.cell(row=row_idx, column=name_col).value:
                ws.cell(row=row_idx, column=name_col).value = emp_name
            return row_idx

    row_idx = ws.max_row + 1
    ws.cell(row=row_idx, column=id_col).value = emp_id
    ws.cell(row=row_idx, column=name_col).value = emp_name
    return row_idx


def refresh_total_column(ws):
    header_map = get_header_map(ws)
    if "Total" not in header_map:
        return

    total_col = header_map["Total"]

    day_cols = []
    for header, col_idx in header_map.items():
        if header not in ["id", "name", "Total"]:
            day_cols.append(col_idx)

    for row_idx in range(2, ws.max_row + 1):
        total_available = 0
        for col_idx in day_cols:
            value = ws.cell(row=row_idx, column=col_idx).value
            if str(value).strip().lower() == "available":
                total_available += 1
        ws.cell(row=row_idx, column=total_col).value = total_available


def update_tracker_excel(results, tracker_file):
    if not results:
        return

    first_date = str(results[0]["date"]).strip()
    try:
        tracker_date = datetime.strptime(first_date, "%Y-%m-%d")
    except:
        tracker_date = datetime.today()

    tracker_sheet_name = get_tracker_sheet_name(tracker_date)
    day_col_name = first_date

    if os.path.exists(tracker_file):
        wb = load_workbook(tracker_file)
    else:
        wb = Workbook()

    if tracker_sheet_name in wb.sheetnames:
        ws = wb[tracker_sheet_name]
    else:
        if len(wb.sheetnames) == 1 and wb.active.max_row == 1 and wb.active.max_column == 1 and wb.active["A1"].value is None:
            ws = wb.active
            ws.title = tracker_sheet_name
        else:
            ws = wb.create_sheet(title=tracker_sheet_name)

    ensure_tracker_headers(ws, day_col_name)
    header_map = get_header_map(ws)

    day_col = header_map[day_col_name]

    for item in results:
        emp_id = item["employee_id"]
        emp_name = item["employee_name"]
        attendance_value = item["attendance"]

        tracker_value = "Available" if attendance_value == "Yes" else "NA"

        row_idx = find_or_create_tracker_row(ws, emp_id, emp_name, header_map)
        cell = ws.cell(row=row_idx, column=day_col)
        cell.value = tracker_value

        if tracker_value == "Available":
            cell.fill = GREEN_FILL
        else:
            cell.fill = PatternFill(fill_type=None)

    refresh_total_column(ws)
    apply_tracker_header_style(ws)
    wb.save(tracker_file)
    wb.close()


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


def get_any_live_page(browser, retries=30, delay=2):
    for _ in range(retries):
        pages = []
        for context in browser.contexts:
            for page in context.pages:
                try:
                    if not page.is_closed():
                        pages.append(page)
                except:
                    pass

        for page in pages:
            try:
                if "moveinsync.com" in page.url:
                    return page
            except:
                pass

        if pages:
            return pages[-1]

        time.sleep(delay)

    raise RuntimeError("Could not find any live browser page")


def close_extra_pages(browser, keep_page):
    try:
        for context in browser.contexts:
            for pg in context.pages[:]:
                try:
                    if pg != keep_page and not pg.is_closed():
                        pg.close()
                except:
                    pass
    except:
        pass


def text_visible(page, text_value, timeout=5000):
    try:
        page.get_by_text(text_value, exact=False).wait_for(timeout=timeout)
        return True
    except:
        return False


def click_text(page, text_value, step_name):
    save_step(page, f"{step_name}_before")
    page.get_by_text(text_value, exact=False).first.click(timeout=10000)
    time.sleep(4)
    save_step(page, f"{step_name}_after")


def get_wis_frame(page):
    page.locator("iframe").first.wait_for(timeout=20000)
    return page.frame_locator("iframe").first


def open_team_calendar(page):
    frame = get_wis_frame(page)
    save_step(page, "04_before_team_calendar_click")

    try:
        frame.get_by_role("button", name="View Team Calendar").click(timeout=10000, force=True)
        time.sleep(5)
        save_step(page, "05_after_team_calendar_click")
        return True
    except:
        pass

    try:
        frame.locator(VIEW_TEAM_CALENDAR_XPATH).first.click(timeout=10000, force=True)
        time.sleep(5)
        save_step(page, "05_after_team_calendar_xpath_click")
        return True
    except:
        pass

    return False


def get_search_box(page):
    frame = get_wis_frame(page)
    try:
        return frame.get_by_role("textbox", name="Search by name, ID or email")
    except:
        return frame.locator(SEARCH_XPATH)


def clear_and_type(search_box, value):
    search_box.click(timeout=5000)
    time.sleep(0.2)
    try:
        search_box.fill("")
        search_box.fill(str(value))
        return
    except:
        pass

    search_box.press("Control+A")
    search_box.press("Backspace")
    search_box.type(str(value), delay=100)


def click_clear_all(page):
    frame = get_wis_frame(page)
    try:
        frame.get_by_text("Clear All").click(timeout=5000, force=True)
        time.sleep(1.5)
        return True
    except:
        pass

    try:
        frame.locator(CLEAR_ALL_XPATH).click(timeout=5000, force=True)
        time.sleep(1.5)
        return True
    except:
        return False


def select_employee_from_result(page, emp_id):
    frame = get_wis_frame(page)

    try:
        frame.get_by_text(f"ID: {emp_id}", exact=False).first.click(timeout=5000, force=True)
        time.sleep(2)
        return True
    except:
        pass

    try:
        frame.get_by_text("ID:", exact=False).first.click(timeout=5000, force=True)
        time.sleep(2)
        return True
    except:
        pass

    return False


def close_status_popup_if_any(page):
    frame = get_wis_frame(page)
    try:
        frame.locator("#close-btn").click(timeout=3000, force=True)
        time.sleep(1)
        return True
    except:
        return False


def get_container_text(locator, timeout=4000):
    try:
        locator.first.wait_for(timeout=timeout)
        return locator.first.inner_text(timeout=timeout).strip()
    except:
        return ""


def parse_status_text(status_text):
    normalized = " ".join(status_text.upper().split())

    if "CHECKED OUT" in normalized:
        return {
            "checked_in": "Yes",
            "checked_out": "Yes"
        }

    if "CHECKED IN" in normalized:
        return {
            "checked_in": "Yes",
            "checked_out": "No"
        }

    if "BOOKED" in normalized:
        return {
            "checked_in": "No",
            "checked_out": "No"
        }

    return {
        "checked_in": "No",
        "checked_out": "No"
    }


def debug_employee_status(emp_id, emp_name, status_xpath, status_text):
    print("=" * 100)
    print(f"EMPLOYEE   : {emp_name}")
    print(f"EMPLOYEE ID: {emp_id}")
    print(f"XPATH USED : {status_xpath}")
    print(f"RAW TEXT   : {repr(status_text)}")
    print("=" * 100)


def extract_status_for_abhinandan(page, emp_id, emp_name):
    frame = get_wis_frame(page)
    locator = frame.locator(ABHINANDAN_STATUS_XPATH)
    status_text = get_container_text(locator)
    debug_employee_status(emp_id, emp_name, ABHINANDAN_STATUS_XPATH, status_text)
    return parse_status_text(status_text)


def extract_status_for_other_employee(page, emp_id, emp_name):
    frame = get_wis_frame(page)
    locator = frame.locator(OTHERS_STATUS_XPATH)
    status_text = get_container_text(locator)
    debug_employee_status(emp_id, emp_name, OTHERS_STATUS_XPATH, status_text)
    return parse_status_text(status_text)


def get_attendance_value(checked_in, checked_out):
    if checked_in == "Yes" or checked_out == "Yes":
        return "Yes"
    return "No"


def main():
    temp_profile = os.path.join(os.getcwd(), "edge_debug_profile")
    os.makedirs(temp_profile, exist_ok=True)

    subprocess.Popen([
        EDGE_EXE,
        f"--remote-debugging-port={CDP_PORT}",
        f"--user-data-dir={temp_profile}",
        WIS_URL
    ])

    if not wait_for_cdp(CDP_PORT, retries=30, delay=1):
        print("CDP did not start")
        return

    employees = pd.read_excel(INPUT_FILE).fillna("")
    results = []

    browser = None
    page = None

    with sync_playwright() as p:
        try:
            browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
            page = get_any_live_page(browser)
            close_extra_pages(browser, page)
            page.set_default_timeout(TIMEOUT_MS)

            try:
                page.goto(WIS_URL, wait_until="domcontentloaded")
            except:
                pass

            time.sleep(8)
            page = get_any_live_page(browser)
            close_extra_pages(browser, page)
            save_step(page, "01_initial_page")

            if text_visible(page, "Pick an account", 5000):
                try:
                    page.locator('[data-test-id="abhinandan.roy@pwc.com"]').click(timeout=10000, force=True)
                    time.sleep(8)
                    page = get_any_live_page(browser)
                    close_extra_pages(browser, page)
                    save_step(page, "02_pick_account")
                except:
                    click_text(page, "abhinandan.roy@pwc.com", "02_pick_account_fallback")
                    time.sleep(8)
                    page = get_any_live_page(browser)
                    close_extra_pages(browser, page)

            save_step(page, "03_after_login")

            opened = open_team_calendar(page)
            if not opened:
                print("Auto-click for Team Calendar failed.")
                input("Open Team Calendar manually, then press Enter here... ")
                save_step(page, "04_team_calendar_manual")
            else:
                save_step(page, "04_team_calendar_auto")

            for _, row in employees.iterrows():
                emp_id = str(row.get("employee_id", "")).strip()
                emp_name = str(row.get("employee_name", "")).strip()

                if not emp_id and not emp_name:
                    continue

                try:
                    if emp_id == ABHINANDAN_ID or emp_name.strip().lower() == ABHINANDAN_NAME.lower():
                        save_step(page, "05_before_abhinandan_status")
                        status = extract_status_for_abhinandan(page, emp_id, emp_name)
                    else:
                        click_clear_all(page)

                        search_box = get_search_box(page)
                        search_box.click(timeout=5000)
                        clear_and_type(search_box, emp_id if emp_id else emp_name)
                        time.sleep(SEARCH_WAIT_SEC)
                        save_step(page, f"06_after_search_{emp_id or emp_name}")

                        selected = select_employee_from_result(page, emp_id)
                        if not selected:
                            print(f"Employee not found in search results: {emp_name} | {emp_id}")
                            results.append({
                                "date": time.strftime("%Y-%m-%d"),
                                "employee_id": emp_id,
                                "employee_name": emp_name,
                                "attendance": "No",
                                "checked_in": "No",
                                "checked_out": "No"
                            })
                            continue

                        save_step(page, f"07_after_select_{emp_id or emp_name}")
                        status = extract_status_for_other_employee(page, emp_id, emp_name)
                        close_status_popup_if_any(page)

                    attendance = get_attendance_value(status["checked_in"], status["checked_out"])

                    results.append({
                        "date": time.strftime("%Y-%m-%d"),
                        "employee_id": emp_id,
                        "employee_name": emp_name,
                        "attendance": attendance,
                        "checked_in": status["checked_in"],
                        "checked_out": status["checked_out"]
                    })

                except Exception as e:
                    print(f"Error for {emp_name} | {emp_id}: {e}")
                    save_step(page, f"08_failed_{emp_id or emp_name}")
                    results.append({
                        "date": time.strftime("%Y-%m-%d"),
                        "employee_id": emp_id,
                        "employee_name": emp_name,
                        "attendance": "Error",
                        "checked_in": "Error",
                        "checked_out": "Error"
                    })

            sheet_name = write_results_to_excel(results, OUTPUT_FILE)
            update_tracker_excel(results, TRACKER_FILE)
            save_step(page, "09_finished")
            print(f"Done. Output saved to {OUTPUT_FILE}, sheet: {sheet_name}")
            print(f"Tracker updated in {TRACKER_FILE}")

        finally:
            try:
                if browser:
                    for context in browser.contexts:
                        for pg in context.pages[:]:
                            try:
                                pg.close()
                            except:
                                pass
            except:
                pass

            try:
                if browser:
                    browser.close()
            except:
                pass


if __name__ == "__main__":
    main()