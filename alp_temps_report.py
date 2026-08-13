"""
ALP Temperatures Email Report Generator - NJT Tech Services

Email sending now goes through SMTP (smtplib) instead of Outlook COM,
so it works regardless of classic vs. new Outlook.

Set the SMTP_* values in the CONFIG block below before first use.
Leave TEST_MODE checked in the GUI until you've confirmed delivery.
"""

# GUI generation
from tkinter import *
from tkinter.ttk import *
from tkinter import messagebox

# Email
import smtplib
import ssl
from email.message import EmailMessage

# Reading the HTML
from bs4 import BeautifulSoup

# Date and time
from datetime import datetime

# Web
import requests
import urllib3
import threading

# Excel
from openpyxl import load_workbook
import os
import sys
import traceback
from pathlib import Path

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# =============================================================================
# CONFIG - set these once
# =============================================================================

# Option A (most likely to work): Microsoft 365 "direct send" to your tenant's
#   MX endpoint. No password. Only delivers to @njtransit.com recipients, which
#   is all of yours. Find the host with:  nslookup -type=mx njtransit.com
# Option B: an internal relay, e.g. "smtp.njtransit.com" / "mailrelay.njtransit.com", port 25.
# Option C: authenticated M365 - host "smtp.office365.com", port 587, and fill
#   in SMTP_USER / SMTP_PASSWORD. Only works if SMTP AUTH is enabled for your mailbox.
SMTP_HOST = "njtransit-com.mail.protection.outlook.com"
SMTP_PORT = 25
SMTP_USER = None          # e.g. "ggenaoperez@njtransit.com" - leave None for A/B
SMTP_PASSWORD = None      # leave None for A/B
SMTP_USE_STARTTLS = True  # auto-skips if the server doesn't advertise it
SMTP_TIMEOUT = 45

FROM_ADDRESS = "ggenaoperez@njtransit.com"
FROM_DISPLAY = "NJT Rail Mech Tech Services"

# Where the test-mode email goes
TEST_ADDRESS = "ggenaoperez@njtransit.com"

TO_RECIPIENTS = [
    "RailMechTechServices@njtransit.com",
    "RailMechQA_QC@njtransit.com",
    "RailWeekendDutyOfficer@njtransit.com",
    "Rail_Mech_MMC_Locomotive_Shop_Foremen@njtransit.com",
    "RailMechanicalDesk@njtransit.com",
    "Rail_Mech_Dover_Yard_Group@njtransit.com",
    "Rail_Mech_Gladstone_Yard_Group@njtransit.com",
    "Rail_Mech_Great_Notch_Yard_Group@njtransit.com",
    "Rail_Mech_Hoboken_Yard_Group@njtransit.com",
    "Rail_Mech_County_Yard_Group@njtransit.com",
    "Rail_Mech_Long_Branch_Yard_Group@njtransit.com",
    "Rail_Mech_Morrisville_Yard_Group@njtransit.com",
    "Rail_Mech_Port_Morris_Yard_Group@njtransit.com",
    "Rail_Mech_Raritan_Yard_Group@njtransit.com",
    "Rail_Mech_Atlantic_City_Yard_Group@njtransit.com",
    "Rail_Mech_Suffern_Yard_Group@njtransit.com",
    "Rail_Mech_Spring_Valley_Yard_Group@njtransit.com",
    "Rail_Mech_New_York-SSYD_Yard_Group@njtransit.com",
    "Rail_Mech_Port_Jervis_Yard_Group@njtransit.com",
    "Rail_Mech_Bay_Head_Yard_Group@njtransit.com",
]

CC_RECIPIENTS = [
    "DDegennaro@njtransit.com",
    "DRogust@njtransit.com",
    "RBreen@njtransit.com",
    "GKunchandy@njtransit.com",
    "APanza@njtransit.com",
    "MOrtland@njtransit.com",
    "YPatel@njtransit.com",
]

TEMP_THRESHOLD = 125.0
OUTPUT_ROOT = r"F:\42 ALPs Converter Temp\NJTDB Temps"
XLSX_MIME = ("application",
             "vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Fleet definitions: (sheet, label, url template, loco range, conv1 col, conv2 col)
FLEETS = [
    ("Sheet1", "ALP-45DP",
     "https://njt.vehicledb.com/converterReport.php",
     range(4500, 4535), 3, 4),
    ("Sheet2", "ALP-46",
     "https://njt.vehicledb.com/converterReport_ALP46.php",
     range(4600, 4629), 2, 3),
    ("Sheet3", "ALP-46A",
     "https://njt.vehicledb.com/converterReport_ALP46A.php",
     range(4629, 4665), 2, 3),
    ("Sheet4", "ALP-45A",
     "https://njt.vehicledb.com/converterReport_alp45a.php",
     range(4535, 4561), 3, 4),
]

# Cache of the most recent successful run, so the email button doesn't rescrape
LAST_REPORT = {"path": None, "concerns": None, "stamp": None}


# =============================================================================
# Helpers
# =============================================================================

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def safe_float(text, default=float("-inf")):
    """Cell text -> float. Returns default for blank / 'N/A' / junk."""
    try:
        return float(str(text).strip().replace("\u00b0", "").replace("F", ""))
    except (TypeError, ValueError):
        return default


def get_outside_temp():
    """Current local temp in F, or None if NWS is unreachable/reporting null."""
    headers = {"User-Agent": f"NJTransitWeatherApp ({FROM_ADDRESS})"}
    latitude, longitude = 40.74392, -74.1029
    try:
        points = requests.get(
            f"https://api.weather.gov/points/{latitude},{longitude}",
            headers=headers, timeout=15).json()
        stations = requests.get(points["properties"]["observationStations"],
                                headers=headers, timeout=15).json()
        station_id = stations["features"][0]["properties"]["stationIdentifier"]
        obs = requests.get(
            f"https://api.weather.gov/stations/{station_id}/observations/latest",
            headers=headers, timeout=15).json()
        temp_c = obs["properties"]["temperature"]["value"]
        if temp_c is None:
            return None
        return round((temp_c * 9 / 5) + 32, 1)
    except Exception:
        return None


def get_date_and_time():
    return datetime.now().strftime("%m/%d/%Y, %I:%M %p")


# =============================================================================
# Scraping
# =============================================================================

def scrape_fleet(url_base, loco_range, col1, col2, session):
    """Returns (rows, concern_locos) for one fleet. Bad rows are skipped, not fatal."""
    today = datetime.now().strftime("%Y-%m-%d")
    rows, concerns = [], []

    for loco in loco_range:
        try:
            response = session.get(f"{url_base}?loco={loco}&date={today}",
                                   auth=("njt", "njtdb"), verify=False, timeout=30)
            soup = BeautifulSoup(response.text, "html.parser")
            table = soup.find("table", id="table")
            columns = table.find_all("td") if table else []

            con1 = con2 = float("-inf")
            if len(columns) > max(col1, col2):
                con1 = safe_float(columns[col1].text)
                con2 = safe_float(columns[col2].text)
        except Exception:
            con1 = con2 = float("-inf")

        if con1 >= TEMP_THRESHOLD or con2 >= TEMP_THRESHOLD:
            concerns.append(loco)

        rows.append([loco, con1, con2])

    return rows, concerns


# =============================================================================
# Excel generation
# =============================================================================

def generate_excel_file(progress=lambda pct, msg: None):
    """Scrapes all fleets, writes the workbook, returns (filepath, concerns dict)."""
    today = datetime.now().strftime("%m.%d.%y %H%M %p")  # note: %H is 24h, %p is AM/PM
    year = datetime.now().strftime("%Y")
    concerns_by_fleet = {}

    progress(0, "Starting...")
    wb = load_workbook(resource_path("Automated ALP Temps TEMPLATE.xlsx"))

    with requests.Session() as session:
        for i, (sheet, label, url, locos, c1, c2) in enumerate(FLEETS):
            progress(i * 20, f"Collecting {label}...")
            rows, concerns = scrape_fleet(url, locos, c1, c2, session)
            rows.sort(reverse=True, key=lambda x: x[2])

            ws = wb[sheet]
            for row in rows:
                ws.append(row)

            concerns_by_fleet[label] = concerns
            print(f"{label} data collected ({len(concerns)} over threshold)")

    outdir = Path(OUTPUT_ROOT) / year
    outdir.mkdir(parents=True, exist_ok=True)
    filepath = outdir / f"ALP TEMPS {today}.xlsx"
    wb.save(filepath)

    progress(100, "File generated to F: drive")

    LAST_REPORT.update({"path": str(filepath),
                        "concerns": concerns_by_fleet,
                        "stamp": today})
    return str(filepath), concerns_by_fleet


# =============================================================================
# Email
# =============================================================================

def format_concerns(concerns_by_fleet):
    lines = []
    for fleet, locos in concerns_by_fleet.items():
        if locos:
            lines.append(f"    {fleet}: " + ", ".join(str(l) for l in locos))
    return "\n".join(lines)


def build_email(concerns_by_fleet, stamp, temp):
    has_concerns = any(concerns_by_fleet.values())
    temp_line = (f"The current outside temperature is {temp}\u00b0F. "
                 if temp is not None else "")

    subject = f"ALP Temps {stamp}"

    if has_concerns:
        body = (
            f"All,\n\n"
            f"Attached is the ALP converter temperature report for {stamp}. "
            f"{temp_line}The unit(s) listed below have converter temperatures of "
            f"{int(TEMP_THRESHOLD)}\u00b0F or higher and require immediate attention. "
            f"All other converter readings are within normal operating limits.\n\n"
            f"{format_concerns(concerns_by_fleet)}\n\n"
            f"Regards,\n{FROM_DISPLAY}"
        )
    else:
        body = (
            f"All,\n\n"
            f"Attached is the ALP converter temperature report for {stamp}. "
            f"{temp_line}All converter readings are within normal operating limits.\n\n"
            f"Regards,\n{FROM_DISPLAY}"
        )

    return subject, body


def send_report_email(subject, body, attachment_path, to_list, cc_list):
    """Sends via SMTP. Raises on failure so the caller can surface the error."""
    msg = EmailMessage()
    msg["From"] = f"{FROM_DISPLAY} <{FROM_ADDRESS}>"
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject
    msg.set_content(body)

    attachment = Path(attachment_path)
    msg.add_attachment(attachment.read_bytes(),
                       maintype=XLSX_MIME[0],
                       subtype=XLSX_MIME[1],
                       filename=attachment.name)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=SMTP_TIMEOUT) as server:
        server.ehlo()
        if SMTP_USE_STARTTLS and server.has_extn("starttls"):
            server.starttls(context=ssl.create_default_context())
            server.ehlo()
        if SMTP_USER and SMTP_PASSWORD:
            server.login(SMTP_USER, SMTP_PASSWORD)
        server.send_message(msg)  # delivers to both To and Cc


# =============================================================================
# GUI plumbing
# =============================================================================

def ui(fn, *args):
    """Schedule a widget update on the main thread (Tk is not thread-safe)."""
    window.after(0, lambda: fn(*args))


def set_progress(pct, msg):
    ui(pb.configure, {"value": pct})
    ui(status.config, {"text": msg})


def set_busy(busy):
    state = DISABLED if busy else NORMAL
    ui(button1.config, {"state": state})
    ui(button2.config, {"state": state})


def show_preview(subject, body, to_list, cc_list, on_send):
    """Stands in for Outlook's mail.Display() - review before it goes out."""
    win = Toplevel(window)
    win.title("Review before sending")
    win.geometry("620x480")

    text = Text(win, wrap="word")
    text.insert("1.0",
                f"From:    {FROM_ADDRESS}\n"
                f"To:      {', '.join(to_list)}\n"
                f"Cc:      {', '.join(cc_list) if cc_list else '(none)'}\n"
                f"Subject: {subject}\n"
                f"Attach:  {Path(LAST_REPORT['path']).name}\n"
                f"{'-' * 70}\n{body}")
    text.config(state=DISABLED)
    text.pack(fill=BOTH, expand=True, padx=8, pady=8)

    row = Frame(win)
    row.pack(pady=6)

    def confirm():
        win.destroy()
        threading.Thread(target=on_send, daemon=True).start()

    Button(row, text="Send", command=confirm).pack(side=LEFT, padx=6)
    Button(row, text="Cancel", command=win.destroy).pack(side=LEFT, padx=6)


# =============================================================================
# Button commands
# =============================================================================

def do_generate():
    set_busy(True)
    try:
        generate_excel_file(set_progress)
        try:
            os.startfile(LAST_REPORT["path"])
        except OSError:
            pass
    except Exception as exc:
        traceback.print_exc()
        set_progress(0, "Generation failed")
        ui(messagebox.showerror, "Report generation failed", str(exc))
    finally:
        set_busy(False)


def do_email():
    set_busy(True)
    try:
        if not LAST_REPORT["path"] or not Path(LAST_REPORT["path"]).exists():
            generate_excel_file(set_progress)

        set_progress(100, "Building email...")
        temp = get_outside_temp()
        subject, body = build_email(LAST_REPORT["concerns"],
                                    LAST_REPORT["stamp"], temp)

        if test_mode.get():
            to_list, cc_list = [TEST_ADDRESS], []
            subject = "[TEST] " + subject
        else:
            to_list, cc_list = TO_RECIPIENTS, CC_RECIPIENTS

        def deliver():
            try:
                set_progress(100, "Sending...")
                send_report_email(subject, body, LAST_REPORT["path"],
                                  to_list, cc_list)
                set_progress(100, "Email sent")
            except Exception as exc:
                traceback.print_exc()
                set_progress(0, "Send failed")
                ui(messagebox.showerror, "Email failed",
                   f"{type(exc).__name__}: {exc}\n\n"
                   f"Check SMTP_HOST / SMTP_PORT at the top of the script.")
            finally:
                set_busy(False)

        ui(show_preview, subject, body, to_list, cc_list, deliver)
        return  # set_busy(False) happens in deliver() or on cancel below

    except Exception as exc:
        traceback.print_exc()
        set_progress(0, "Failed")
        ui(messagebox.showerror, "Error", str(exc))
        set_busy(False)


def button1commands():
    status.config(text="Generating...")
    threading.Thread(target=do_generate, daemon=True).start()


def button2commands():
    status.config(text="Preparing email...")
    threading.Thread(target=do_email, daemon=True).start()


# =============================================================================
# Window
# =============================================================================

window = Tk()
window.title("ALP Temperatures Email Report Generator - NJT Tech Services")
window.geometry("300x210")

status = Label(window, text="Click to use")
button1 = Button(window, text="Generate ALP Temps Excel File", command=button1commands)
button2 = Button(window, text="Send ALP Temps Email Report", command=button2commands)
pb = Progressbar(window, mode="determinate", length=200, maximum=100)

test_mode = BooleanVar(value=True)
test_check = Checkbutton(window, text=f"Test mode (send only to me)", variable=test_mode)

status.grid(row=0, column=0, padx=10, pady=8)
button1.grid(row=1, column=0, padx=10, pady=6)
button2.grid(row=2, column=0, padx=10, pady=2)
test_check.grid(row=3, column=0, padx=10, pady=6)
pb.grid(row=4, column=0, padx=10, pady=8)

window.mainloop()

# "App works faster than todd gets his monthly report done" -The Intern, 2026
