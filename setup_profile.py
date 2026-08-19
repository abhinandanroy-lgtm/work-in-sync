import os
import csv
import subprocess
import time
import base64
import calendar
from datetime import datetime

import requests
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font
from playwright.sync_api import sync_playwright
from azure.communication.email import EmailClient
from dotenv import load_dotenv

load_dotenv()

WIS_URL = "https://pwc.moveinsync.com/WP/employee.jsp#WorkInSyncDashboard"
EDGE_EXE = r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
CDP_PORT = 9222
TIMEOUT_MS = 30000
SEARCH_WAIT_SEC = 2

INPUT_FILE = "employee_input.xlsx"
OUTPUT_FILE = "attendance_output.xlsx"
TRACKER_FILE = "Tracker.xlsx"
TEST_TRACKER_FILE = "Test_Tracker.xlsx"
FAILED_AWARENESS_CSV = "failed_awareness_mails.csv"

ABHINANDAN_ID = "101675341"
ABHINANDAN_NAME = "Abhinandan Roy"
ABHINANDAN_EMAIL = "abhinandan.roy@pwc.com"
DEBORSHI_EMAIL = "deborshi.som@pwc.com"

ACS_CONNECTION = os.getenv("ACS_CONNECTION")
ACS_SENDER_EMAIL = "DoNotReply@e6aa12e1-707d-4e66-845f-2f5e8845675c.azurecomm.net"

if not ACS_CONNECTION:
    raise ValueError("ACS_CONNECTION environment variable is not set. Please add it in .env file.")

MONTHLY_TARGET_DAYS = 10
MAIL_RETRY_COUNT = 6
MAIL_THROTTLE_DELAY_SEC = 20

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


def today_str():
    return datetime.today().strftime("%Y-%m-%d")


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
        ws.cell(row=1, column=2).value = "email"
        ws.cell(row=1, column=3).value = "name"
        ws.cell(row=1, column=4).value = day_col_name
        ws.cell(row=1, column=5).value = "Total"
        apply_tracker_header_style(ws)
        return

    headers = [
        str(ws.cell(row=1, column=i).value).strip()
        if ws.cell(row=1, column=i).value is not None else ""
        for i in range(1, ws.max_column + 1)
    ]

    if "id" not in headers:
        ws.insert_cols(1)
        ws.cell(row=1, column=1).value = "id"

    headers = [
        str(ws.cell(row=1, column=i).value).strip()
        if ws.cell(row=1, column=i).value is not None else ""
        for i in range(1, ws.max_column + 1)
    ]

    if "email" not in headers:
        ws.insert_cols(2)
        ws.cell(row=1, column=2).value = "email"

    headers = [
        str(ws.cell(row=1, column=i).value).strip()
        if ws.cell(row=1, column=i).value is not None else ""
        for i in range(1, ws.max_column + 1)
    ]

    if "name" not in headers:
        ws.insert_cols(3)
        ws.cell(row=1, column=3).value = "name"

    headers = [
        str(ws.cell(row=1, column=i).value).strip()
        if ws.cell(row=1, column=i).value is not None else ""
        for i in range(1, ws.max_column + 1)
    ]

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

    headers = [
        str(ws.cell(row=1, column=i).value).strip()
        if ws.cell(row=1, column=i).value is not None else ""
        for i in range(1, ws.max_column + 1)
    ]

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
            existing_name = ws.cell(row=row_idx, column=name_col).value
            if emp_name and not existing_name:
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
    day_cols = [
        col_idx for header, col_idx in header_map.items()
        if header not in ["id", "email", "name", "Total"]
    ]

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
    except Exception:
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
        if (
            len(wb.sheetnames) == 1
            and wb.active.max_row == 1
            and wb.active.max_column == 1
            and wb.active["A1"].value is None
        ):
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


def send_html_email(to_email, subject, html_body, plain_text=None, display_name=None, max_retries=MAIL_RETRY_COUNT):
    attempt = 0

    while attempt < max_retries:
        try:
            client = EmailClient.from_connection_string(ACS_CONNECTION)

            recipient = {"address": to_email}
            if display_name:
                recipient["displayName"] = display_name

            message = {
                "senderAddress": ACS_SENDER_EMAIL,
                "recipients": {"to": [recipient]},
                "content": {
                    "subject": subject,
                    "plainText": plain_text or "Your office availability summary is ready.",
                    "html": html_body,
                },
            }

            poller = client.begin_send(message)
            result = poller.result()
            print(f"Email sent to {to_email}. Message ID: {result['id']}")
            return True

        except Exception as e:
            err = str(e)
            attempt += 1

            if "TooManyRequests" in err:
                wait_time = min(20 * attempt, 120)
                print(f"Rate limited while sending to {to_email}. Retry {attempt}/{max_retries} after {wait_time}s...")
                time.sleep(wait_time)
            else:
                print(f"Failed to send html email to {to_email}: {e}")
                return False

    print(f"Failed to send html email to {to_email} after {max_retries} retries.")
    return False


def send_html_email_multi(to_emails, subject, html_body, plain_text=None, max_retries=MAIL_RETRY_COUNT):
    attempt = 0

    while attempt < max_retries:
        try:
            client = EmailClient.from_connection_string(ACS_CONNECTION)

            recipients = [{"address": mail.strip()} for mail in to_emails if str(mail).strip()]

            message = {
                "senderAddress": ACS_SENDER_EMAIL,
                "recipients": {"to": recipients},
                "content": {
                    "subject": subject,
                    "plainText": plain_text or "Monthly attendance report",
                    "html": html_body,
                },
            }

            poller = client.begin_send(message)
            result = poller.result()
            print(f"Email sent to {', '.join(to_emails)}. Message ID: {result['id']}")
            return True

        except Exception as e:
            err = str(e)
            attempt += 1

            if "TooManyRequests" in err:
                wait_time = min(20 * attempt, 120)
                print(f"Rate limited while sending report. Retry {attempt}/{max_retries} after {wait_time}s...")
                time.sleep(wait_time)
            else:
                print(f"Failed to send report email: {e}")
                return False

    print("Failed to send report email after retries.")
    return False


def get_today_column_name():
    return datetime.today().strftime("%Y-%m-%d")


def get_month_sheet_or_none(workbook_path):
    if not os.path.exists(workbook_path):
        return None, None, None

    wb = load_workbook(workbook_path, data_only=True)
    sheet_name = get_tracker_sheet_name(datetime.today())
    if sheet_name not in wb.sheetnames:
        wb.close()
        return None, None, None

    ws = wb[sheet_name]
    header_map = get_header_map(ws)
    return wb, ws, header_map


def read_employee_input():
    df = pd.read_excel(INPUT_FILE).fillna("")
    rename_map = {}
    for col in df.columns:
        low = str(col).strip().lower()
        if low in ["id", "employee id", "employee_id"]:
            rename_map[col] = "employee_id"
        elif low in ["name", "employee name", "employee_name"]:
            rename_map[col] = "employee_name"
    df = df.rename(columns=rename_map)

    if "employee_id" not in df.columns:
        df["employee_id"] = ""
    if "employee_name" not in df.columns:
        df["employee_name"] = ""

    return df


def read_monthly_employees_from_test_tracker():
    wb, ws, header_map = get_month_sheet_or_none(TEST_TRACKER_FILE)
    if wb is None:
        raise Exception("Test_Tracker monthly sheet not found. Please ensure the current month sheet exists.")

    if "id" not in header_map or "name" not in header_map:
        wb.close()
        raise Exception("Test_Tracker monthly sheet must contain 'id' and 'name' columns.")

    rows = []
    for row_idx in range(2, ws.max_row + 1):
        emp_id = str(ws.cell(row=row_idx, column=header_map["id"]).value or "").strip()
        emp_name = str(ws.cell(row=row_idx, column=header_map["name"]).value or "").strip()

        if not emp_id and not emp_name:
            continue

        rows.append({
            "employee_id": emp_id,
            "employee_name": emp_name
        })

    wb.close()
    return pd.DataFrame(rows)


def merge_unique_employees(df1, df2):
    combined = pd.concat([df1[["employee_id", "employee_name"]], df2[["employee_id", "employee_name"]]], ignore_index=True)
    combined["employee_id"] = combined["employee_id"].astype(str).str.strip()
    combined["employee_name"] = combined["employee_name"].astype(str).str.strip()
    combined = combined[(combined["employee_id"] != "") | (combined["employee_name"] != "")]
    combined = combined.drop_duplicates(subset=["employee_id"], keep="first")
    return combined.reset_index(drop=True)


def build_awareness_mail_html(employee_name, employee_id, days_in_office, days_pending):
    html = f"""<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Employee Office Availability Metrics</title>
  <!--[if mso]>
  <style type="text/css">
    table {{border-collapse: collapse;}}
    .fallback-font {{font-family: Arial, sans-serif !important;}}
  </style>
  <![endif]-->
</head>
<body style="margin:0; padding:0; background-color:#f2f2f2;">

  <div style="display:none; max-height:0; overflow:hidden;">
    Your Office availability summary is ready.
  </div>

  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f2f2f2;">
    <tr>
      <td align="center" style="padding: 40px 15px;">

        <table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:600px; max-width:600px; background-color:#ffffff; box-shadow:0 2px 10px rgba(0,0,0,0.06);">

          <tr>
            <td style="padding:0;">
              <!--[if mso]>
              <v:rect xmlns:v="urn:schemas-microsoft-com:vml" fill="true" stroke="false" style="width:600px;mso-width-percent:0;">
              <v:fill type="gradient" color="#ffffff" color2="#F58025" angle="135" />
              <v:textbox inset="0,0,0,0" style="mso-fit-shape-to-text:false;">
              <![endif]-->
              <div style="background: linear-gradient(135deg, #ffffff 0%, #ffe6d1 25%, #ffb066 60%, #F58025 100%);">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                  <tr>
                    <td style="padding: 36px 40px 0 40px;">
                      <span style="font-family: Arial, sans-serif; font-size:12px; color:#a8340a; letter-spacing:1.2px; font-weight:bold;">
                        PwC&nbsp;&nbsp;|&nbsp;&nbsp;Canada IT
                      </span>
                    </td>
                  </tr>
                  <tr>
                    <td style="padding: 10px 40px 0 40px;">
                      <span style="font-family: Arial, sans-serif; font-size:25px; color:#000000; font-weight:bold; line-height:32px;">
                        Office Availability Summary
                      </span>
                    </td>
                  </tr>
                  <tr>
                    <td style="padding: 16px 40px 32px 40px;">
                      <table role="presentation" cellpadding="0" cellspacing="0">
                        <tr>
                          <td style="background-color:#000000; height:4px; width:60px; line-height:4px; font-size:1px;">&nbsp;</td>
                        </tr>
                      </table>
                    </td>
                  </tr>
                </table>
              </div>
              <!--[if mso]>
              </v:textbox>
              </v:rect>
              <![endif]-->
            </td>
          </tr>

          <tr>
            <td style="padding: 28px 40px 10px 40px;">
              <p style="margin:0; font-family: Arial, sans-serif; font-size:15px; color:#000000; line-height:22px;">
                Hi <strong>{employee_name}</strong>,
              </p>

              <p style="margin:10px 0 12px 0; font-family: Arial, sans-serif; font-size:14px; color:#555555; line-height:21px;">
                Here's a snapshot of your office availability progress for the current month. If you have any remaining in-office days to meet your monthly availability requirement, we encourage you to plan them at your convenience and stay on track before the month concludes.
              </p>

              <p style="margin:0 0 12px 0; font-family: Arial, sans-serif; font-size:14px; color:#555555; line-height:21px;">
                If your current bookings do not reflect your in-office availability, please take a moment to review and book them in WorkInSync from the next time, so your schedule remains up to date.
              </p>

              <p style="margin:0 0 12px 0; font-family: Arial, sans-serif; font-size:14px; color:#555555; line-height:21px;">
                If you have already discussed and aligned your availability plan with your People Manager, please feel free to disregard this notification.
              </p>

              <p style="margin:0; font-family: Arial, sans-serif; font-size:14px; color:#555555; line-height:21px;">
                Thank you for your continued collaboration and support in fostering a positive workplace experience.
              </p>
            </td>
          </tr>

          <tr>
            <td style="padding: 20px 40px 10px 40px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #eeeeee;">
                <tr>
                  <td style="padding: 16px 20px; background-color:#fafafa;" width="50%">
                    <span style="font-family: Arial, sans-serif; font-size:11px; color:#999999; text-transform:uppercase; letter-spacing:0.5px; font-weight:bold;">Employee Name</span>
                    <br>
                    <span style="font-family: Arial, sans-serif; font-size:16px; color:#000000; font-weight:bold;">{employee_name}</span>
                  </td>
                  <td style="padding: 16px 20px; background-color:#fafafa; border-left:1px solid #eeeeee;" width="50%">
                    <span style="font-family: Arial, sans-serif; font-size:11px; color:#999999; text-transform:uppercase; letter-spacing:0.5px; font-weight:bold;">Employee ID</span>
                    <br>
                    <span style="font-family: Arial, sans-serif; font-size:16px; color:#000000; font-weight:bold;">{employee_id}</span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="padding: 24px 40px 36px 40px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="48%" valign="top" style="background: linear-gradient(135deg, #fff0e4 0%, #ffd9b8 100%); border:1px solid #F58025; padding:24px 20px;">
                    <span style="font-family: Arial, sans-serif; font-size:11px; color:#a8340a; font-weight:bold; text-transform:uppercase; letter-spacing:0.5px;">
                      Days in Office
                    </span>
                    <br>
                    <span style="font-family: Arial, sans-serif; font-size:40px; color:#F58025; font-weight:bold; line-height:50px;">
                      {days_in_office}
                    </span>
                    <span style="font-family: Arial, sans-serif; font-size:13px; color:#333333;"> / {MONTHLY_TARGET_DAYS} days</span>
                  </td>

                  <td width="4%">&nbsp;</td>

                  <td width="48%" valign="top" style="background-color:#000000; padding:24px 20px; border:1px solid #000000;">
                    <span style="font-family: Arial, sans-serif; font-size:11px; color:#ffffff; font-weight:bold; text-transform:uppercase; letter-spacing:0.5px;">
                      Days Pending
                    </span>
                    <br>
                    <span style="font-family: Arial, sans-serif; font-size:40px; color:#FF8A3D; font-weight:bold; line-height:50px;">
                      {days_pending}
                    </span>
                    <span style="font-family: Arial, sans-serif; font-size:13px; color:#cccccc;"> days left</span>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <tr>
            <td style="background-color:#000000; padding: 22px 40px;">
              <p style="margin:0; font-family: Arial, sans-serif; font-size:11px; color:#ffffff; line-height:16px; text-align:center;">
                This is an automated message from Canada IT. For questions, contact
                <a href="mailto:shreyasi.dutta@pwc.com" style="color:#F58025; text-decoration:underline;">shreyasi.dutta@pwc.com</a>.
                <br>
                © 2026 PwC. All rights reserved.
              </p>
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>

</body>
</html>"""
    return html


def log_failed_awareness_mail(employee_id, employee_name, employee_email, reason):
    file_exists = os.path.exists(FAILED_AWARENESS_CSV)
    with open(FAILED_AWARENESS_CSV, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(["date", "employee_id", "employee_name", "employee_email", "reason"])
        writer.writerow([today_str(), employee_id, employee_name, employee_email, reason])


def clear_failed_awareness_log():
    if os.path.exists(FAILED_AWARENESS_CSV):
        os.remove(FAILED_AWARENESS_CSV)


def read_failed_awareness_recipients():
    if not os.path.exists(FAILED_AWARENESS_CSV):
        return []

    rows = []
    with open(FAILED_AWARENESS_CSV, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def rewrite_failed_awareness_log(rows):
    if not rows:
        clear_failed_awareness_log()
        return

    with open(FAILED_AWARENESS_CSV, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["date", "employee_id", "employee_name", "employee_email", "reason"])
        for row in rows:
            writer.writerow(row)


def get_days_in_office_from_test_tracker(employee_id):
    wb, ws, header_map = get_month_sheet_or_none(TEST_TRACKER_FILE)
    if wb is None:
        return 0

    try:
        if "id" not in header_map or "Total" not in header_map:
            wb.close()
            return 0

        for row_idx in range(2, ws.max_row + 1):
            existing_id = str(ws.cell(row=row_idx, column=header_map["id"]).value or "").strip()
            if existing_id == str(employee_id).strip():
                try:
                    val = int(ws.cell(row=row_idx, column=header_map["Total"]).value or 0)
                except Exception:
                    val = 0
                wb.close()
                return val
    except Exception:
        pass

    try:
        wb.close()
    except Exception:
        pass
    return 0


def send_awareness_mails_from_test_tracker(test_tracker_file):
    wb, ws, header_map = get_month_sheet_or_none(test_tracker_file)
    if wb is None:
        print("Test tracker monthly sheet not found.")
        return

    required = ["id", "email", "name", "Total"]
    for col in required:
        if col not in header_map:
            print(f"Required column missing in test tracker: {col}")
            wb.close()
            return

    clear_failed_awareness_log()

    sent_count = 0
    fail_count = 0

    for row_idx in range(2, ws.max_row + 1):
        employee_id = str(ws.cell(row=row_idx, column=header_map["id"]).value or "").strip()
        employee_email = str(ws.cell(row=row_idx, column=header_map["email"]).value or "").strip()
        employee_name = str(ws.cell(row=row_idx, column=header_map["name"]).value or "").strip()
        total_value = ws.cell(row=row_idx, column=header_map["Total"]).value

        if not employee_id and not employee_name:
            continue
        if not employee_email:
            continue

        try:
            days_in_office = int(total_value)
        except Exception:
            days_in_office = 0

        days_pending = max(MONTHLY_TARGET_DAYS - days_in_office, 0)

        html_body = build_awareness_mail_html(
            employee_name=employee_name,
            employee_id=employee_id,
            days_in_office=days_in_office,
            days_pending=days_pending
        )

        ok = send_html_email(
            to_email=employee_email,
            subject="Office Availability Summary",
            html_body=html_body,
            plain_text=f"Hi {employee_name}, your office availability summary is ready. Days in office: {days_in_office}. Days pending: {days_pending}.",
            display_name=employee_name,
            max_retries=MAIL_RETRY_COUNT
        )

        if ok:
            sent_count += 1
        else:
            fail_count += 1
            log_failed_awareness_mail(employee_id, employee_name, employee_email, "Failed after retries")

        time.sleep(MAIL_THROTTLE_DELAY_SEC)

    wb.close()
    print(f"Awareness mail run completed. Sent: {sent_count}, Failed: {fail_count}")

    if fail_count > 0:
        print(f"Failed recipient log saved to: {FAILED_AWARENESS_CSV}")


def resend_failed_awareness_mails():
    failed_rows = read_failed_awareness_recipients()
    if not failed_rows:
        print("No failed awareness mail log found or no failed recipients.")
        return

    print(f"Retrying failed awareness mails: {len(failed_rows)} recipients")

    still_failed = []
    success_count = 0

    for row in failed_rows:
        employee_id = str(row.get("employee_id", "")).strip()
        employee_name = str(row.get("employee_name", "")).strip()
        employee_email = str(row.get("employee_email", "")).strip()

        if not employee_email:
            still_failed.append([today_str(), employee_id, employee_name, employee_email, "Missing email"])
            continue

        days_in_office = get_days_in_office_from_test_tracker(employee_id)
        days_pending = max(MONTHLY_TARGET_DAYS - days_in_office, 0)

        html_body = build_awareness_mail_html(
            employee_name=employee_name,
            employee_id=employee_id,
            days_in_office=days_in_office,
            days_pending=days_pending
        )

        ok = send_html_email(
            to_email=employee_email,
            subject="Office Availability Summary",
            html_body=html_body,
            plain_text=f"Hi {employee_name}, your office availability summary is ready. Days in office: {days_in_office}. Days pending: {days_pending}.",
            display_name=employee_name,
            max_retries=MAIL_RETRY_COUNT
        )

        if ok:
            success_count += 1
        else:
            still_failed.append([today_str(), employee_id, employee_name, employee_email, "Failed again"])

        time.sleep(MAIL_THROTTLE_DELAY_SEC)

    rewrite_failed_awareness_log(still_failed)

    if still_failed:
        print(f"Some recipients still failed. Updated log: {FAILED_AWARENESS_CSV}")
    else:
        clear_failed_awareness_log()
        print("All previously failed awareness mails sent successfully.")

    print(f"Resend summary. Success: {success_count}, Remaining failed: {len(still_failed)}")


def get_working_days_in_month(year, month):
    cal = calendar.monthcalendar(year, month)
    working_days = 0
    for week in cal:
        for day_index in range(0, 5):
            if week[day_index] != 0:
                working_days += 1
    return working_days


def read_tracker_summary(tracker_file):
    wb, ws, header_map = get_month_sheet_or_none(tracker_file)
    if wb is None:
        return None, "Tracker monthly sheet not found."

    required = ["id", "name", "Total"]
    for col in required:
        if col not in header_map:
            wb.close()
            return None, f"Required column missing: {col}"

    rows = []
    total_emp = 0
    total_attendance_sum = 0

    for row_idx in range(2, ws.max_row + 1):
        emp_id = str(ws.cell(row=row_idx, column=header_map["id"]).value or "").strip()
        emp_name = str(ws.cell(row=row_idx, column=header_map["name"]).value or "").strip()
        total_val = ws.cell(row=row_idx, column=header_map["Total"]).value

        if not emp_id and not emp_name:
            continue

        try:
            total_attendance = int(total_val)
        except Exception:
            total_attendance = 0

        total_emp += 1
        total_attendance_sum += total_attendance

        rows.append({
            "id": emp_id,
            "name": emp_name,
            "total": total_attendance
        })

    wb.close()

    avg_attendance = round(total_attendance_sum / total_emp, 2) if total_emp else 0
    today = datetime.today()
    working_days = get_working_days_in_month(today.year, today.month)

    summary = {
        "total_emp": total_emp,
        "avg_attendance": avg_attendance,
        "working_days": working_days,
        "rows": rows,
        "month": today.strftime("%B"),
        "year": today.strftime("%Y"),
        "report_date": today.strftime("%Y-%m-%d")
    }
    return summary, None


def build_month_end_report_html(tracker_file):
    summary, error = read_tracker_summary(tracker_file)
    if error:
        return f"<html><body><p>{error}</p></body></html>"

    table_rows = ""
    sorted_rows = sorted(summary["rows"], key=lambda x: x["total"], reverse=True)

    for item in sorted_rows:
        total = item["total"]

        if total < 10:
            bg_color = "#FFE8D9"
            text_color = "#C0562A"
        elif total == 10:
            bg_color = "#E3F7EA"
            text_color = "#2E9E5B"
        else:
            bg_color = "#1A7A4A"
            text_color = "#FFFFFF"

        table_rows += f"""
                    <tr class="data-row" style="border-bottom:1px solid #f0f0f0;">
                        <td style="padding:11px 15px; font-size:13px; color:#333333; font-family:Arial,sans-serif;">{item['name']}</td>
                        <td style="padding:11px 15px; font-size:13px; color:#333333; font-family:Arial,sans-serif;">{item['id']}</td>
                        <td align="center" style="padding:11px 15px; font-size:13px; font-family:Arial,sans-serif;">
                            <span style="display:inline-block; min-width:42px; padding:5px 10px; border-radius:14px; background-color:{bg_color}; color:{text_color}; font-weight:700;">
                                {total}
                            </span>
                        </td>
                    </tr>
        """

    html = f"""<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Monthly Attendance Report</title>
<!--[if mso]>
<noscript><xml><o:OfficeDocumentSettings><o:PixelsPerInch>96</o:PixelsPerInch></o:OfficeDocumentSettings></xml></noscript>
<![endif]-->
<style type="text/css">
    body, table, td, p {{ -webkit-text-size-adjust: 100%; -ms-text-size-adjust: 100%; }}
    table, td {{ mso-table-lspace: 0pt; mso-table-rspace: 0pt; }}
    body {{ margin:0 !important; padding:0 !important; width:100% !important; font-family: Arial, Helvetica, sans-serif; }}
    @media only screen and (max-width: 600px) {{
        .email-container {{ width: 100% !important; }}
        .mobile-padding {{ padding-left: 15px !important; padding-right: 15px !important; }}
        .responsive-table td, .responsive-table th {{ font-size: 12px !important; padding: 6px 4px !important; }}
        .header-text {{ font-size: 20px !important; }}
        .intro-text {{ font-size: 14px !important; }}
        .header-pad {{ padding: 26px 20px 24px 20px !important; }}
    }}
    .data-row:hover {{ background-color: #FFF8F3 !important; }}
</style>
</head>
<body style="margin:0; padding:0; background-color:#f4f4f4;">

<div style="display:none; max-height:0; overflow:hidden; mso-hide:all;">
    Monthly attendance summary report for {summary['total_emp']} employees - Review office attendance data
</div>

<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#f4f4f4;">
<tr><td align="center" style="padding:20px 10px;">

<table role="presentation" class="email-container" width="800" cellpadding="0" cellspacing="0" border="0" style="max-width:800px; width:100%; background-color:#ffffff; border-radius:8px; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.08);">

    <tr>
        <td style="padding:0;">

            <!--[if mso]>
            <v:rect xmlns:v="urn:schemas-microsoft-com:vml" fill="true" stroke="false" style="width:800px;">
            <v:fill type="gradient" color="#FFF6EE" color2="#E8620E" angle="45" />
            <v:textbox inset="0,0,0,0">
            <![endif]-->

            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
                style="background-color:#F2854A; background-image: linear-gradient(120deg, #FFF6EE 0%, #FCD3AC 40%, #E8620E 100%);">
                <tr>
                    <td class="header-pad" style="padding:34px 40px 30px 40px;">

                        <p style="margin:0 0 12px 0; font-size:12px; font-weight:700; letter-spacing:1.6px; color:#B34A17; font-family:Arial,sans-serif;">
                            PwC &nbsp;|&nbsp; Canada IT
                        </p>

                        <h1 class="header-text" style="margin:0 0 8px 0; color:#241f1a; font-size:25px; font-weight:700; line-height:1.3; font-family:Arial,sans-serif;">
                            Monthly Employee Office Availability Report
                        </h1>

                        <div style="width:64px; height:3px; background-color:#241f1a; margin:0 0 14px 0; font-size:0; line-height:0;">&nbsp;</div>

                        <p style="margin:0; color:#5a4230; font-size:14px; font-family:Arial,sans-serif;">
                            {summary['month']} / {summary['year']} &nbsp;|&nbsp; Generated on {summary['report_date']}
                        </p>

                    </td>
                </tr>
            </table>

            <!--[if mso]>
            </v:textbox>
            </v:rect>
            <![endif]-->

        </td>
    </tr>

    <tr>
        <td style="padding:28px 40px 20px 40px;" class="mobile-padding">
            <p style="margin:0 0 15px 0; color:#1a1a1a; font-size:15px; line-height:1.6; font-family:Arial,sans-serif;">
                Dear Shreyasi di,
            </p>
            <p class="intro-text" style="margin:0 0 15px 0; color:#4a4a4a; font-size:15px; line-height:1.6; font-family:Arial,sans-serif;">
                Please find below the consolidated employee office availability report for Canada IT team for <strong style="color:#E8620E;">{summary['month']} / {summary['year']}</strong>. This report presents employee office availability against the monthly benchmark, providing insights into overall alignment with the organization's office presence expectations.
            </p>
            <p class="intro-text" style="margin:0 0 20px 0; color:#4a4a4a; font-size:15px; line-height:1.6; font-family:Arial,sans-serif;">
                Kindly review the data below. For discrepancies, please reach out to Abhinandan Roy.
            </p>

            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                    <td width="33%" style="padding:4px;">
                        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#FFF3EA; border-radius:6px; border:1px solid #FFE0C7;">
                            <tr><td style="padding:15px; text-align:center;">
                                <p style="margin:0; font-size:22px; font-weight:700; color:#C0562A; font-family:Arial,sans-serif;">{summary['total_emp']}</p>
                                <p style="margin:0; font-size:12px; color:#9a6642; font-family:Arial,sans-serif;">Total Employees</p>
                            </td></tr>
                        </table>
                    </td>
                    <td width="33%" style="padding:4px;">
                        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#EEFBF3; border-radius:6px; border:1px solid #D2F0DE;">
                            <tr><td style="padding:15px; text-align:center;">
                                <p style="margin:0; font-size:22px; font-weight:700; color:#1A7A4A; font-family:Arial,sans-serif;">{summary['avg_attendance']}</p>
                                <p style="margin:0; font-size:12px; color:#4a9a6a; font-family:Arial,sans-serif;">Avg. Days/Employee</p>
                            </td></tr>
                        </table>
                    </td>
                    <td width="33%" style="padding:4px;">
                        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color:#FFF8EC; border-radius:6px; border:1px solid #FFE9C4;">
                            <tr><td style="padding:15px; text-align:center;">
                                <p style="margin:0; font-size:22px; font-weight:700; color:#B06A1A; font-family:Arial,sans-serif;">{summary['working_days']}</p>
                                <p style="margin:0; font-size:12px; color:#c08a4a; font-family:Arial,sans-serif;">Working Days</p>
                            </td></tr>
                        </table>
                    </td>
                </tr>
            </table>
        </td>
    </tr>

    <tr>
        <td style="padding:10px 40px 30px 40px;" class="mobile-padding">
            <table role="presentation" class="responsive-table" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse; border:1px solid #e5e5e5; border-radius:6px; overflow:hidden;">

                <thead>
                    <tr>
                        <th align="left" style="background-color:#E8620E; color:#ffffff; padding:12px 15px; font-size:13px; font-family:Arial,sans-serif; font-weight:600;">Employee Name</th>
                        <th align="left" style="background-color:#E8620E; color:#ffffff; padding:12px 15px; font-size:13px; font-family:Arial,sans-serif; font-weight:600;">Employee ID</th>
                        <th align="center" style="background-color:#E8620E; color:#ffffff; padding:12px 15px; font-size:13px; font-family:Arial,sans-serif; font-weight:600;">Total Days in Office</th>
                    </tr>
                </thead>

                <tbody>
{table_rows}
                </tbody>

            </table>

            <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:15px;">
                <tr><td style="font-size:11px; color:#8a8a8a; font-family:Arial,sans-serif; line-height:1.8;">
                    <span style="background-color:#FFE8D9; color:#C0562A; padding:2px 8px; border-radius:10px; font-weight:600;">●</span> Towards Achievement (&lt; 10 days)
                    &nbsp;&nbsp;
                    <span style="background-color:#E3F7EA; color:#2E9E5B; padding:2px 8px; border-radius:10px; font-weight:600;">●</span> Achieved (= 10 days)
                    &nbsp;&nbsp;
                    <span style="background-color:#1A7A4A; color:#FFFFFF; padding:2px 8px; border-radius:10px; font-weight:600;">●</span> Above Achievement (&gt; 10 days)
                </td></tr>
            </table>
        </td>
    </tr>

    <tr>
        <td style="background-color:#fafafa; padding:24px 40px; border-top:1px solid #e5e5e5;" class="mobile-padding">
            <p style="margin:0; font-size:12px; color:#a0a0a0; font-family:Arial,sans-serif;">
                © 2026 PwC. All rights reserved.
            </p>
        </td>
    </tr>

</table>
</td></tr>
</table>

</body>
</html>"""
    return html


def send_month_end_report(tracker_file):
    html_body = build_month_end_report_html(tracker_file)
    recipients = [ABHINANDAN_EMAIL, DEBORSHI_EMAIL]
    send_html_email_multi(
        recipients,
        "Monthly Employee Office Availability Report",
        html_body,
        plain_text="Monthly employee office availability report."
    )


def prompt_manual_awareness_trigger(test_tracker_file):
    try:
        choice = input("Do you want to manually trigger awareness mails to employees from Test_Tracker.xlsx? (yes/no): ").strip().lower()
        if choice == "yes":
            send_awareness_mails_from_test_tracker(test_tracker_file)
    except Exception as e:
        print(f"Manual awareness trigger failed: {e}")


def prompt_resend_failed_awareness_trigger():
    try:
        choice = input("Do you want to retry failed awareness mails from failed_awareness_mails.csv? (yes/no): ").strip().lower()
        if choice == "yes":
            resend_failed_awareness_mails()
    except Exception as e:
        print(f"Retry failed awareness trigger failed: {e}")


def prompt_manual_month_end_report_trigger(tracker_file):
    try:
        choice = input("Do you want to manually trigger month-end overall report to Abhinandan and Deborshi? (yes/no): ").strip().lower()
        if choice == "yes":
            send_month_end_report(tracker_file)
    except Exception as e:
        print(f"Manual month-end report trigger failed: {e}")


def save_step(page, name):
    try:
        page.screenshot(path=os.path.join(SCREENSHOT_DIR, f"{ts()}_{name}.png"), full_page=True, timeout=8000)
    except Exception:
        pass
    try:
        with open(os.path.join(HTML_DIR, f"{ts()}_{name}.html"), "w", encoding="utf-8") as f:
            f.write(page.content())
    except Exception:
        pass
    try:
        txt = page.locator("body").inner_text(timeout=5000)
        with open(os.path.join(TEXT_DIR, f"{ts()}_{name}.txt"), "w", encoding="utf-8") as f:
            f.write(txt)
    except Exception:
        pass


def wait_for_cdp(port=9222, retries=30, delay=1):
    url = f"http://127.0.0.1:{port}/json/version"
    for _ in range(retries):
        try:
            r = requests.get(url, timeout=2)
            if r.status_code == 200:
                return True
        except Exception:
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
                except Exception:
                    pass

        for page in pages:
            try:
                if "moveinsync.com" in page.url:
                    return page
            except Exception:
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
                except Exception:
                    pass
    except Exception:
        pass


def text_visible(page, text_value, timeout=5000):
    try:
        page.get_by_text(text_value, exact=False).wait_for(timeout=timeout)
        return True
    except Exception:
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
    except Exception:
        pass

    try:
        frame.locator(VIEW_TEAM_CALENDAR_XPATH).first.click(timeout=10000, force=True)
        time.sleep(5)
        save_step(page, "05_after_team_calendar_xpath_click")
        return True
    except Exception:
        pass

    return False


def get_search_box(page):
    frame = get_wis_frame(page)
    try:
        return frame.get_by_role("textbox", name="Search by name, ID or email")
    except Exception:
        return frame.locator(SEARCH_XPATH)


def clear_and_type(search_box, value):
    search_box.click(timeout=5000)
    time.sleep(0.2)
    try:
        search_box.fill("")
        search_box.fill(str(value))
        return
    except Exception:
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
    except Exception:
        pass
    try:
        frame.locator(CLEAR_ALL_XPATH).click(timeout=5000, force=True)
        time.sleep(1.5)
        return True
    except Exception:
        return False


def select_employee_from_result(page, emp_id, emp_name):
    frame = get_wis_frame(page)

    if emp_id:
        try:
            frame.get_by_text(f"ID: {emp_id}", exact=False).first.click(timeout=5000, force=True)
            time.sleep(2)
            return True
        except Exception:
            pass

    if emp_name:
        try:
            frame.get_by_text(emp_name, exact=False).first.click(timeout=5000, force=True)
            time.sleep(2)
            return True
        except Exception:
            pass

    try:
        frame.get_by_text("ID:", exact=False).first.click(timeout=5000, force=True)
        time.sleep(2)
        return True
    except Exception:
        pass

    return False


def close_status_popup_if_any(page):
    frame = get_wis_frame(page)
    try:
        frame.locator("#close-btn").click(timeout=3000, force=True)
        time.sleep(1)
        return True
    except Exception:
        return False


def get_container_text(locator, timeout=4000):
    try:
        locator.first.wait_for(timeout=timeout)
        return locator.first.inner_text(timeout=timeout).strip()
    except Exception:
        return ""


def parse_status_text(status_text):
    normalized = " ".join(status_text.upper().split())

    if "CHECKED OUT" in normalized:
        return {"checked_in": "Yes", "checked_out": "Yes"}
    if "CHECKED IN" in normalized:
        return {"checked_in": "Yes", "checked_out": "No"}
    if "BOOKED" in normalized:
        return {"checked_in": "No", "checked_out": "No"}
    return {"checked_in": "No", "checked_out": "No"}


def extract_status_for_abhinandan(page):
    frame = get_wis_frame(page)
    locator = frame.locator(ABHINANDAN_STATUS_XPATH)
    status_text = get_container_text(locator)
    return parse_status_text(status_text)


def extract_status_for_other_employee(page):
    frame = get_wis_frame(page)
    locator = frame.locator(OTHERS_STATUS_XPATH)
    status_text = get_container_text(locator)
    return parse_status_text(status_text)


def get_attendance_value(checked_in, checked_out):
    return "Yes" if checked_in == "Yes" or checked_out == "Yes" else "No"


def main():
    temp_profile = os.path.join(os.getcwd(), "edge_debug_profile")
    os.makedirs(temp_profile, exist_ok=True)

    subprocess.Popen(
        [
            EDGE_EXE,
            f"--remote-debugging-port={CDP_PORT}",
            f"--user-data-dir={temp_profile}",
            WIS_URL,
        ]
    )

    if not wait_for_cdp(CDP_PORT, retries=30, delay=1):
        print("CDP did not start")
        return

    input_employees = read_employee_input()

    try:
        test_tracker_employees = read_monthly_employees_from_test_tracker()
    except Exception as e:
        print(f"Error reading Test_Tracker.xlsx: {e}")
        return

    employees = merge_unique_employees(input_employees, test_tracker_employees)
    results = []

    with sync_playwright() as p:
        browser = None
        try:
            browser = p.chromium.connect_over_cdp(f"http://127.0.0.1:{CDP_PORT}")
            page = get_any_live_page(browser)
            close_extra_pages(browser, page)
            page.set_default_timeout(TIMEOUT_MS)

            try:
                page.goto(WIS_URL, wait_until="domcontentloaded")
            except Exception:
                pass

            time.sleep(8)
            page = get_any_live_page(browser)
            save_step(page, "01_initial_page")

            if text_visible(page, "Pick an account", 5000):
                try:
                    page.locator('[data-test-id="abhinandan.roy@pwc.com"]').click(timeout=10000, force=True)
                    time.sleep(8)
                    page = get_any_live_page(browser)
                except Exception:
                    click_text(page, "abhinandan.roy@pwc.com", "02_pick_account_fallback")
                    time.sleep(8)
                    page = get_any_live_page(browser)

            save_step(page, "03_after_login")

            opened = open_team_calendar(page)
            if not opened:
                print("Auto-click for Team Calendar failed.")
                input("Open Team Calendar manually, then press Enter here... ")

            for _, row in employees.iterrows():
                emp_id = str(row.get("employee_id", "")).strip()
                emp_name = str(row.get("employee_name", "")).strip()

                if not emp_id and not emp_name:
                    continue

                try:
                    if emp_id == ABHINANDAN_ID or emp_name.lower() == ABHINANDAN_NAME.lower():
                        status = extract_status_for_abhinandan(page)
                    else:
                        click_clear_all(page)
                        search_box = get_search_box(page)
                        search_value = emp_id if emp_id else emp_name
                        clear_and_type(search_box, search_value)
                        time.sleep(SEARCH_WAIT_SEC)

                        selected = select_employee_from_result(page, emp_id, emp_name)
                        if not selected:
                            results.append(
                                {
                                    "date": today_str(),
                                    "employee_id": emp_id,
                                    "employee_name": emp_name,
                                    "attendance": "No",
                                    "checked_in": "No",
                                    "checked_out": "No",
                                }
                            )
                            continue

                        status = extract_status_for_other_employee(page)
                        close_status_popup_if_any(page)

                    attendance = get_attendance_value(status["checked_in"], status["checked_out"])

                    results.append(
                        {
                            "date": today_str(),
                            "employee_id": emp_id,
                            "employee_name": emp_name,
                            "attendance": attendance,
                            "checked_in": status["checked_in"],
                            "checked_out": status["checked_out"],
                        }
                    )

                except Exception as e:
                    print(f"Error for {emp_name} | {emp_id}: {e}")
                    results.append(
                        {
                            "date": today_str(),
                            "employee_id": emp_id,
                            "employee_name": emp_name,
                            "attendance": "Error",
                            "checked_in": "Error",
                            "checked_out": "Error",
                        }
                    )

            sheet_name = write_results_to_excel(results, OUTPUT_FILE)

            update_tracker_excel(results, TRACKER_FILE)
            update_tracker_excel(results, TEST_TRACKER_FILE)

            print(f"Updated {TRACKER_FILE} for today: {today_str()}")
            print(f"Updated {TEST_TRACKER_FILE} for today: {today_str()}")

            prompt_manual_awareness_trigger(TEST_TRACKER_FILE)
            prompt_resend_failed_awareness_trigger()
            prompt_manual_month_end_report_trigger(TRACKER_FILE)

            print(f"Done. Output saved to {OUTPUT_FILE}, sheet: {sheet_name}")

        finally:
            try:
                if browser:
                    browser.close()
            except Exception:
                pass


if __name__ == "__main__":
    main() 