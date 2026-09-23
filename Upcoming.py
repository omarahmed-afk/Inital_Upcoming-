"""Export today's WebPT visits to one Google Sheet tab.

One-time install:
    py -m pip install selenium gspread google-auth tzdata openpyxl

Google setup:
1. Enable the Google Sheets and Google Drive APIs.
2. Create a service account and download its JSON key.
3. Share the target Google Sheet with the JSON key's client_email as Editor.
4. Set GOOGLE_SHEET_ID through an environment variable. Credentials
   default to feedback-509416-9cbf06d0baf6.json beside this script; override with
   GOOGLE_SERVICE_ACCOUNT_FILE or GOOGLE_SERVICE_ACCOUNT_JSON.

Set WEBPT_USERNAME and WEBPT_PASSWORD through environment variables.
Normal runs write only to the Upcoming tab.
Upcoming is refreshed with one visit per patient: Other, Confirmed, Checked Out,
then Checked In in priority order,
then earliest date. The report covers today through December 31 this year.
Local Excel test (sample data): py visits.py --test-excel
Local Excel test (real CSV): py visits.py --test-excel visits_local.xlsx --input-csv export.csv
Each tab contains: EMR ID | Clinic Name | Patient Name |
Appointment Type | Appointment Date | Visit Status | Phone (blank).
"""

import os
import time
import shutil
import csv
import json
import tempfile
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
from gspread.exceptions import WorksheetNotFound
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, WebDriverException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains



# ── WebPT Credentials ──────────────────────────────────────────
WEBPT_USERNAME = os.getenv("WEBPT_USERNAME", "")
WEBPT_PASSWORD = os.getenv("WEBPT_PASSWORD", "")
WEBPT_URL      = os.getenv("WEBPT_URL", "https://app.webpt.com")

# ── Timezone ───────────────────────────────────────────────────
try:
    from zoneinfo import ZoneInfo
    REPORT_TZ = ZoneInfo("America/New_York")
except Exception:
    REPORT_TZ = None                         # fall back to system local time

DATE_START = DATE_END = DATE_LABEL = None

def compute_dates():
    """Set DATE_START / DATE_END / DATE_LABEL to TODAY in New York time."""
    global DATE_START, DATE_END, DATE_LABEL

    now = datetime.now(REPORT_TZ) if REPORT_TZ else datetime.now()

    DATE_START = now.strftime("%m/%d/%Y")
    DATE_END   = now.strftime("%m/%d/%Y")
    DATE_LABEL = now.strftime("%Y-%m-%d")

    return DATE_LABEL

# ── Google Sheets Output ───────────────────────────────────────
# Sheet ID = the value between /d/ and /edit in the Google Sheet URL.
GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID", "")
GOOGLE_SERVICE_ACCOUNT_JSON = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "")
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv(
    "GOOGLE_SERVICE_ACCOUNT_FILE",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "feedback-509416-9cbf06d0baf6.json"),
)
# ── Cleaning / Counting Rules ──────────────────────────────────
# Appointment types that count as TWO patients instead of one (matched loosely):
DOUBLE_APPT_TYPES = ["initial exam", "new case"]

# Appointment types to drop entirely (matched loosely, case-insensitive):
EXCLUDE_APPT_TYPES = ["sensory freeway", "home care","PTOC - Telehealth"]

# Clinic names to drop entirely (matched loosely, case-insensitive):
EXCLUDE_CLINIC_NAMES = ["sensory", "home care",'PTOC - Telehealth']

# ── Clinic Groups  ────────────────────────────────────
GOOGLE_TAB_NAME = os.getenv("GOOGLE_TAB_NAME", "All")

# ╚══════════════════════════════════════════════════════════════╝

DOWNLOAD_DIR = tempfile.mkdtemp()
driver = None


def validate_configuration():
    missing = [
        name
        for name, value in (
            ("WEBPT_USERNAME", WEBPT_USERNAME),
            ("WEBPT_PASSWORD", WEBPT_PASSWORD),
            ("GOOGLE_SHEET_ID", GOOGLE_SHEET_ID),
        )
        if not value
    ]
    if not GOOGLE_SERVICE_ACCOUNT_JSON and not (
        GOOGLE_SERVICE_ACCOUNT_FILE and os.path.isfile(GOOGLE_SERVICE_ACCOUNT_FILE)
    ):
        missing.append("GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE")
    if missing:
        raise RuntimeError("Missing required environment variable(s): " + ", ".join(missing))


# ── Logging ────────────────────────────────────────────────────
def log(msg):
    ts = datetime.now().strftime("%H:%M:%S")
    print(f"[{ts}] {msg}", flush=True)


# ── Deduplication / Cleaning ───────────────────────────────────

def parse_time(time_str):
    if not time_str:
        return datetime.max
    time_str = time_str.strip()
    for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M"):
        try:
            return datetime.strptime(time_str, fmt)
        except:
            pass
    return datetime.max


def find_col(header, keywords):
    """Return the index of the first header cell containing any keyword."""
    for i, h in enumerate(header):
        if any(k in h.lower() for k in keywords):
            return i
    return None


def _is_double(row, col_type):
    """True if the appointment type counts as two patients."""
    t = row[col_type].strip().lower() if (col_type is not None and col_type < len(row)) else ""
    return any(k.lower() in t for k in DOUBLE_APPT_TYPES)


def _is_excluded_type(row, col_type):
    """True if the appointment type is Sensory / Home Care (drop these entirely)."""
    t = row[col_type].strip().lower() if (col_type is not None and col_type < len(row)) else ""
    return any(k.lower() in t for k in EXCLUDE_APPT_TYPES)


def _is_excluded_clinic(row, col_clinic):
    """True if the clinic name is Sensory / Home Care (drop these entirely)."""
    clinic = row[col_clinic].strip().lower() if (col_clinic is not None and col_clinic < len(row)) else ""
    return any(k.lower() in clinic for k in EXCLUDE_CLINIC_NAMES)


def _status_key(row, col_status):
    status = _norm(_cell(row, col_status))
    return {"checkin": "checkedin", "checkout": "checkedout",
            "canceled": "cancelled"}.get(status, status)


def deduplicate_rows(all_rows):
    """Keep one attended row per displayed visit, preferring Checked Out.

    Match the five exported identity fields, ignoring hidden appointment IDs
    and times. Visits identical in those fields collapse to one row.
    """
    if not all_rows:
        return all_rows

    header = all_rows[0]
    data   = all_rows[1:]

    col_patient_id = _find_exact_col(header, ["EMR ID", "EMRID", "Patient ID", "PatientID"])
    col_date       = find_col(header, ["appointment date", "appt date", "visit date", "date"])
    col_status     = find_col(header, ["visit status", "status"])
    col_type       = find_col(header, ["appointment type", "appt type"])
    col_clinic     = find_clinic_col(header)
    col_patient_name = find_col(header, ["patient name", "patientname"])

    log(f"   ✅ Columns → PatientID:{col_patient_id}  Date:{col_date}  Status:{col_status}  "
        f"ApptType:{col_type}  Clinic:{col_clinic}  PatientName:{col_patient_name}")

    # ── Step A: drop Sensory / Home Care rows by appointment type and clinic name ──
    if col_type is None:
        log("   ⚠️ No Appointment Type column — skipping Sensory/Home Care appointment-type removal")
        data_after_type = data
    else:
        data_after_type, removed_type = [], 0
        for row in data:
            if _is_excluded_type(row, col_type):
                removed_type += 1
            else:
                data_after_type.append(row)
        log(f"   🗑️  Removed {removed_type} Sensory / Home Care appointment rows")
    data = data_after_type

    if col_clinic is None:
        log("   ⚠️ No Clinic column — skipping Sensory/Home Care clinic-name removal")
        data_after_clinic = data
    else:
        data_after_clinic, removed_clinic = [], 0
        for row in data:
            if _is_excluded_clinic(row, col_clinic):
                removed_clinic += 1
            else:
                data_after_clinic.append(row)
        log(f"   🗑️  Removed {removed_clinic} Sensory / Home Care clinic rows")
    data = data_after_clinic
    if col_status is None:
        raise RuntimeError("Visit Status column is required to filter Checked In / Checked Out")
    attended = {"checkedin", "checkedout"}
    data = [row for row in data if _status_key(row, col_status) in attended]

    if col_patient_id is None:
        log("No patient identifier column; keeping all remaining rows")
        return [header] + data

    appointments = {}
    for index, row in enumerate(data):
        # Use exactly the same ID and name selection as the final report.
        displayed = _build_google_values([header, row])[1]
        if not displayed[0]:
            key = ("unidentified", index)
        else:
            key = tuple(" ".join(value.split()).casefold() for value in displayed[:5])
        appointments.setdefault(key, []).append(row)

    result = []
    for rows in appointments.values():
        result.append(next(
            (row for row in rows if _status_key(row, col_status) == "checkedout"),
            rows[0],
        ))

    log(f"Removed {len(data) - len(result)} duplicate appointment rows")
    log(f"After cleaning: {len(result)} rows kept")
    return [header] + result


def find_clinic_col(header):
    for i, h in enumerate(header):
        if "clinic" in h.lower():
            return i
    return None


def _norm(s):
    """Normalize a clinic name for fuzzy matching: lowercase, keep only alphanumerics."""
    return "".join(ch for ch in str(s).lower().strip() if ch.isalnum())


def compute_clinic_counts(rows):
    """
    Count patients per clinic on already-cleaned rows (every clinic present in the
    export). Initial Examination / New Case counts as 2.
    Returns a dict {webpt_clinic_name: total}.
    """
    counts = {}
    if not rows or len(rows) < 2:
        return counts

    header     = rows[0]
    col_clinic = find_clinic_col(header)
    col_type   = find_col(header, ["appointment type", "appt type"])

    if col_clinic is None:
        log("   ⚠️ No clinic column — cannot compute per-clinic counts")
        return counts

    for row in rows[1:]:
        clinic = row[col_clinic].strip() if col_clinic < len(row) else ""
        if not clinic:
            continue
        weight = 2 if _is_double(row, col_type) else 1
        counts[clinic] = counts.get(clinic, 0) + weight

    log("\n── Clinic Counts (Initial/New Case = 2) ──")
    for c in sorted(counts):
        log(f"   • {c}: {counts[c]}")
    log(f"   Σ Total patients: {sum(counts.values())}  |  Clinics with visits: {len(counts)}")
    return counts


# ── Google Sheets ──────────────────────────────────────────────
def _cell(row, col):
    if col is None or col >= len(row):
        return ""
    return str(row[col]).strip()


def _find_exact_col(header, names):
    wanted = {_norm(name) for name in names}
    for i, value in enumerate(header):
        if _norm(value) in wanted:
            return i
    return None


def _patient_name_columns(header):
    """Support either one Patient Name column or separate first/last columns."""
    full_name = _find_exact_col(
        header,
        ["Patient Name", "Patient Full Name", "Full Name", "Patient"],
    )
    first_name = _find_exact_col(
        header,
        ["Patient First Name", "First Name", "PatientFirstName"],
    )
    last_name = _find_exact_col(
        header,
        ["Patient Last Name", "Last Name", "PatientLastName"],
    )
    return full_name, first_name, last_name


def _build_google_values(rows):
    """Build the seven report columns for one Google Sheet tab."""
    output = [[
        "EMR ID",
        "Clinic Name",
        "Patient Name",
        "Appointment Type",
        "Appointment Date",
        "Visit Status",
        "Phone",
    ]]
    if not rows:
        return output

    header = rows[0]
    col_clinic = find_clinic_col(header)
    col_patient_id = _find_exact_col(header, ["EMR ID", "EMRID"])
    if col_patient_id is None:
        col_patient_id = find_col(header, ["patient id", "patientid"])
    col_status = find_col(header, ["visit status", "status"])
    col_appointment_type = find_col(
        header,
        ["appointment type", "appt type"],
    )
    col_appointment_date = find_col(
        header,
        ["appointment date", "appt date", "visit date", "date"],
    )
    col_full_name, col_first_name, col_last_name = _patient_name_columns(header)

    missing = []
    if col_clinic is None:
        missing.append("Clinic")
    if col_patient_id is None:
        missing.append("EMR ID / Patient ID")
    if col_status is None:
        missing.append("Visit Status")
    if col_appointment_type is None:
        missing.append("Appointment Type")
    if col_appointment_date is None:
        missing.append("Appointment Date")
    if col_full_name is None and col_first_name is None and col_last_name is None:
        missing.append("Patient Name")
    if missing:
        raise RuntimeError(
            "Required CSV columns were not found: "
            + ", ".join(missing)
            + f". Exported headers: {header}"
        )

    for row in rows[1:]:
        if col_full_name is not None:
            patient_name = _cell(row, col_full_name)
        else:
            patient_name = " ".join(
                part
                for part in (
                    _cell(row, col_first_name),
                    _cell(row, col_last_name),
                )
                if part
            )

        output.append([
            _cell(row, col_patient_id),
            _cell(row, col_clinic),
            patient_name,
            _cell(row, col_appointment_type),
            _cell(row, col_appointment_date),
            _cell(row, col_status),
            "",  # Phone is intentionally blank.
        ])

    output[1:] = sorted(
        output[1:],
        key=lambda row: (
            row[0].casefold(),
            row[1].casefold(),
            row[2].casefold(),
        ),
    )
    return output


def _report_tabs(cleaned_rows):
    """Split visits into All (excluding Initial Examination) and Initial Examination."""
    values = _build_google_values(cleaned_rows)
    all_values = [values[0]]
    initial_values = [values[0]]
    for row in values[1:]:
        if " ".join(row[3].split()).casefold() == "initial examination":
            initial_values.append(row)
        else:
            all_values.append(row)
    if GOOGLE_TAB_NAME.casefold() == "initial examination":
        raise RuntimeError("GOOGLE_TAB_NAME must differ from Initial Examination")
    return [(GOOGLE_TAB_NAME, all_values), ("Initial Examination", initial_values)]


UPCOMING_HEADERS = [
    "Patient ID", "Clinic Name", "Patient Name", "Treating Therapist",
    "Case Therapist", "Appointment Type", "Appointment Date", "Visit Status",
]


def _appointment_datetime(value):
    value = value.strip()
    try:
        parsed = datetime.fromisoformat(value)
        if parsed.tzinfo is not None:
            parsed = parsed.astimezone(REPORT_TZ).replace(tzinfo=None)
        return parsed
    except ValueError:
        pass
    for fmt in ("%m/%d/%Y", "%m/%d/%Y %I:%M %p", "%m/%d/%Y %I:%M:%S %p",
                "%m/%d/%Y %H:%M", "%m/%d/%Y %H:%M:%S"):
        try:
            return datetime.strptime(value, fmt)
        except ValueError:
            pass
    raise ValueError(f"Unrecognized appointment date: {value!r}")


def _build_upcoming_values(rows, today=None):
    """Choose Other, Confirmed, Checked Out, then Checked In; one row per patient."""
    output = [UPCOMING_HEADERS[:]]
    if not rows:
        return output
    today = today or datetime.now(REPORT_TZ).date()
    header = rows[0]
    patient_id = _find_exact_col(header, ["Patient ID", "EMR ID", "EMRID"])
    clinic = find_clinic_col(header)
    full, first, last = _patient_name_columns(header)
    treating = _find_exact_col(header, ["Treating Therapist"])
    case = _find_exact_col(header, ["Case Therapist"])
    appt_type = find_col(header, ["appointment type", "appt type"])
    date = find_col(header, ["appointment date", "appt date", "visit date"])
    status = find_col(header, ["visit status", "status"])
    columns = [patient_id, clinic, full if full is not None else first if first is not None else last,
               treating, case, appt_type, date, status]
    missing = [name for name, col in zip(UPCOMING_HEADERS, columns) if col is None]
    if missing:
        raise RuntimeError("Upcoming requires CSV columns: " + ", ".join(missing))
    priorities = {"other": 0, "confirmed": 1, "checkedout": 2, "checkedin": 3}
    chosen = {}
    for row in rows[1:]:
        state = _status_key(row, status)
        if state not in priorities:
            continue
        if _is_excluded_type(row, appt_type) or _is_excluded_clinic(row, clinic):
            continue
        when = _appointment_datetime(_cell(row, date))
        if when.date() < today or when.year > today.year:
            continue
        identifier = _cell(row, patient_id)
        if not identifier:
            raise RuntimeError("Upcoming appointment has no Patient ID")
        rank = (priorities[state], when)
        name = _cell(row, full) if full is not None else " ".join(
            part for part in (_cell(row, first), _cell(row, last)) if part
        )
        values = [identifier, _cell(row, clinic), name, _cell(row, treating),
                  _cell(row, case), _cell(row, appt_type), _cell(row, date), _cell(row, status)]
        key = identifier.casefold()
        if key not in chosen or rank < chosen[key][0]:
            chosen[key] = (rank, values)
    output.extend(value for rank, value in sorted(chosen.values(), key=lambda item: (item[0][1], item[1][0])))
    return output


def _today_rows(rows):
    col_date = find_col(rows[0], ["appointment date", "appt date", "visit date"])
    if col_date is None:
        raise RuntimeError("Appointment Date column is required")
    today = datetime.now(REPORT_TZ).date()
    return [rows[0]] + [row for row in rows[1:]
                        if _appointment_datetime(_cell(row, col_date)).date() == today]


def _write_upcoming_tab(spreadsheet, values):
    try:
        worksheet = spreadsheet.worksheet("Upcoming")
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(title="Upcoming", rows=max(100, len(values)), cols=8)
    existing = worksheet.get("A:H")
    count = max(len(values), len(existing))
    if worksheet.row_count < count or worksheet.col_count < 8:
        worksheet.resize(rows=max(worksheet.row_count, count), cols=max(worksheet.col_count, 8))
    # Overwrite the snapshot and blank stale rows in the same request.
    worksheet.update(values=values + [[""] * 8 for _ in range(count - len(values))],
                     range_name=f"A1:H{count}", value_input_option="RAW")
    worksheet.freeze(rows=1)
    log(f"Upcoming: saved {len(values) - 1} patients")


def write_local_excel(cleaned_rows, output_path, upcoming_rows=None):
    """Export All and Initial Examination to a local Excel workbook."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    tabs = [("Upcoming", _build_upcoming_values(
        upcoming_rows if upcoming_rows is not None else cleaned_rows
    ))]
    for index, (tab_name, values) in enumerate(tabs):
        worksheet = workbook.active if index == 0 else workbook.create_sheet()
        worksheet.title = tab_name
        for row in values:
            worksheet.append(row)
        # Preserve IDs, dates, and source text exactly as in the Sheets export.
        for row in worksheet.iter_rows():
            for cell in row:
                cell.data_type = "s"
                cell.number_format = "@"
        for cell in worksheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="2E4F8F")
            cell.alignment = Alignment(horizontal="center")
        for col in range(1, len(values[0]) + 1):
            worksheet.column_dimensions[get_column_letter(col)].width = 28
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
    workbook.save(output_path)
    log(f"Local Excel file saved: {os.path.abspath(output_path)}")


def test_local_excel(output_path, input_csv=None):
    """Run local report processing without browser or Google Sheets access."""
    if input_csv:
        with open(input_csv, "r", encoding="utf-8-sig", newline="") as source:
            rows = list(csv.reader(source))
        if not rows:
            raise RuntimeError("The input CSV is empty")
    else:
        compute_dates()
        rows = [[
            "Patient ID", "Clinic Name", "Patient Name",
            "Appointment Type", "Appointment Date", "Visit Status",
            "Treating Therapist", "Case Therapist",
        ]]
        for index, clinic in enumerate(["Allerton", "Belmont", "Bushwick"], start=1):
            rows.append([
                f"{index:06d}", clinic, f"SAMPLE Patient {index}",
                "Follow Up", DATE_START, "Checked In", "Sample Therapist", "Sample Case Therapist",
            ])
            rows.append([f"{index:06d}", clinic, f"SAMPLE Patient {index}",
                         "Follow Up", DATE_START, "Other", "Sample Therapist", "Sample Case Therapist"])
        log("Using SAMPLE data only; these are not real patient visits.")
    write_local_excel(rows, output_path)


def _write_google_tab(spreadsheet, tab_name, values):
    """Append new visits in A:F and update existing statuses, never writing G."""
    try:
        worksheet = spreadsheet.worksheet(tab_name)
    except WorksheetNotFound:
        worksheet = spreadsheet.add_worksheet(
            title=tab_name,
            rows=max(100, len(values)),
            cols=7,
        )

    existing = worksheet.get("A:G")
    if existing and existing[0][:6] != values[0][:6]:
        raise RuntimeError(f"Unexpected header in {tab_name}; history was not modified")

    # A:E identify a visit. F is mutable; G belongs entirely to the user.
    def history_key(row):
        return tuple(str(value) for value in (list(row) + [""] * 5)[:5])

    saved_rows = {}
    for row_number, row in enumerate(existing[1:], start=2):
        if any(row[:5]):
            saved_rows.setdefault(history_key(row), []).append((row_number, row))
    incoming = {history_key(row): row[:6] for row in values[1:]}
    additions = []
    updates = []
    for key, row in incoming.items():
        if key not in saved_rows:
            additions.append(row)
        else:
            for row_number, saved in saved_rows[key]:
                if _cell(saved, 5) != row[5]:
                    updates.append({"range": f"F{row_number}", "values": [[row[5]]]})

    pending = additions if existing else [values[0][:6]] + additions
    start_row = len(existing) + 1
    required_rows = max(worksheet.row_count, len(existing) + len(pending), 100)
    required_cols = max(worksheet.col_count, 7)
    if required_rows != worksheet.row_count or required_cols != worksheet.col_count:
        worksheet.resize(rows=required_rows, cols=required_cols)

    # Write only below existing report rows; never clear historical data.
    if pending:
        worksheet.update(
            values=pending,
            range_name=f"A{start_row}:F{start_row + len(pending) - 1}",
            value_input_option="RAW",
        )
    if updates:
        worksheet.batch_update(updates, value_input_option="RAW")

    spreadsheet.batch_update({
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": worksheet.id,
                        "gridProperties": {"frozenRowCount": 1},
                    },
                    "fields": "gridProperties.frozenRowCount",
                }
            },
            {
                "repeatCell": {
                    "range": {
                        "sheetId": worksheet.id,
                        "startRowIndex": 0,
                        "endRowIndex": 1,
                        "startColumnIndex": 0,
                        "endColumnIndex": 6,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": 0.18,
                                "green": 0.31,
                                "blue": 0.56,
                            },
                            "textFormat": {
                                "foregroundColor": {
                                    "red": 1,
                                    "green": 1,
                                    "blue": 1,
                                },
                                "bold": True,
                            },
                            "horizontalAlignment": "CENTER",
                        }
                    },
                    "fields": "userEnteredFormat",
                }
            },
            {
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": worksheet.id,
                        "dimension": "COLUMNS",
                        "startIndex": 0,
                        "endIndex": 6,
                    },
                    "properties": {"pixelSize": 220},
                    "fields": "pixelSize",
                }
            },
        ]
    })
    log(f"   ✅ {tab_name}: appended {len(additions)} visits, updated {len(updates)} statuses")


def write_google_sheet_tabs(rows):
    """Refresh Upcoming only; do not write to any other worksheet."""
    if GOOGLE_SHEET_ID in ("", "PASTE_GOOGLE_SHEET_ID_HERE"):
        raise RuntimeError("Set GOOGLE_SHEET_ID before running the script")
    if not rows:
        raise RuntimeError("No rows are available for Google Sheets")

    header = rows[0]
    col_clinic = find_clinic_col(header)
    if col_clinic is None:
        raise RuntimeError("The exported CSV does not contain a Clinic column")

    upcoming_values = _build_upcoming_values(rows)
    if GOOGLE_SERVICE_ACCOUNT_JSON:
        try:
            service_account_info = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
        except json.JSONDecodeError as exc:
            raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON") from exc
        credentials = Credentials.from_service_account_info(
            service_account_info,
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive",
            ],
        )
        client = gspread.authorize(credentials)
    else:
        if not os.path.isfile(GOOGLE_SERVICE_ACCOUNT_FILE):
            raise FileNotFoundError(
                "Google service-account JSON was not found at: "
                f"{GOOGLE_SERVICE_ACCOUNT_FILE}"
            )
        client = gspread.service_account(filename=GOOGLE_SERVICE_ACCOUNT_FILE)
    spreadsheet = client.open_by_key(GOOGLE_SHEET_ID)

    _write_upcoming_tab(spreadsheet, upcoming_values)
    return upcoming_values


# ── Browser ────────────────────────────────────────────────────
def open_browser():
    global driver
    log("🌐 Opening Chrome...")
    prefs = {
        "download.default_directory": DOWNLOAD_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True
    }
    options = webdriver.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-notifications")
    if os.getenv("HEADLESS", "false").lower() in {"1", "true", "yes"}:
        options.add_argument("--headless=new")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    log("✅ Chrome opened")


def wait_click(by, value, timeout=20, desc=""):
    log(f"   ⏳ Clicking: {desc or value}")
    el = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
    try:
        el.click()
    except:
        driver.execute_script("arguments[0].click();", el)
    log(f"   ✅ Clicked: {desc or value}")
    time.sleep(1)
    return el


def wait_visible(by, value, timeout=20, desc=""):
    log(f"   ⏳ Waiting: {desc or value}")
    el = WebDriverWait(driver, timeout).until(EC.visibility_of_element_located((by, value)))
    log(f"   ✅ Found: {desc or value}")
    return el


def clear_type(el, text):
    try:
        el.click()
    except:
        driver.execute_script("arguments[0].click();", el)
    time.sleep(0.3)
    try:
        el.clear()
    except:
        pass
    el.send_keys(Keys.CONTROL + "a")
    el.send_keys(Keys.DELETE)
    time.sleep(0.2)
    el.send_keys(text)
    time.sleep(0.3)


def wait_for_new_window(existing, timeout=10):
    end = time.time() + timeout
    while time.time() < end:
        new = set(driver.window_handles) - set(existing)
        if new:
            return new.pop()
        time.sleep(0.5)
    return None


# ── Login ──────────────────────────────────────────────────────
def login():
    log("\n=== LOGIN ===")
    driver.get(WEBPT_URL)
    time.sleep(4)
    try:
        f = wait_visible(By.ID, "username", desc="Username")
    except:
        f = wait_visible(By.XPATH, "//input[@type='text' or @type='email']", desc="Username")
    clear_type(f, WEBPT_USERNAME)
    try:
        wait_click(By.CSS_SELECTOR, "button[type='submit']", desc="Continue")
    except:
        wait_click(By.XPATH, "//button[contains(text(),'Continue')]", desc="Continue")
    time.sleep(3)
    try:
        f = wait_visible(By.ID, "password", desc="Password")
    except:
        f = wait_visible(By.XPATH, "//input[@type='password']", desc="Password")
    clear_type(f, WEBPT_PASSWORD)
    try:
        wait_click(By.CSS_SELECTOR, "button[type='submit']", desc="Sign In")
    except:
        wait_click(By.XPATH, "//button[contains(text(),'Sign')]", desc="Sign In")
    time.sleep(5)
    try:
        WebDriverWait(driver, 10).until(
            lambda d: "delegator" in d.current_url or "evict" in d.current_url
                      or len(d.find_elements(By.CSS_SELECTOR, "button.eviction-option.ok")) > 0
        )
        log("   ⚠️ Eviction page — clicking 'Yes, oust them!'")
        oust = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.eviction-option.ok"))
        )
        oust.click()
        log("   ✅ Ousted!")
        time.sleep(4)
    except:
        log("   ℹ️ No eviction page")
    log("✅ Logged in")


# ── Navigation ─────────────────────────────────────────────────
def go_to_analytics():
    log("\n=== ANALYTICS ===")
    try:
        nav = wait_visible(By.CSS_SELECTOR, "li:nth-child(3) > a > b", desc="Nav hover")
        ActionChains(driver).move_to_element(nav).perform()
        time.sleep(1)
    except:
        pass
    wins_before = driver.window_handles
    try:
        wait_click(By.CSS_SELECTOR, ".analytics-icon", timeout=60, desc="Analytics icon")
    except TimeoutException:
        try:
            wait_click(
                By.XPATH,
                "//a[contains(translate(normalize-space(.), "
                "'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'analytics')]",
                timeout=60, desc="Analytics",
            )
        except TimeoutException as exc:
            screenshot = os.path.join(os.path.dirname(os.path.abspath(__file__)), "analytics_timeout.png")
            try:
                driver.save_screenshot(screenshot)
            except WebDriverException:
                pass
            raise RuntimeError(
                "Analytics was not clickable after two 60-second attempts. "
                "Check analytics_timeout.png to see whether login completed or the menu is visible."
            ) from exc
    new_win = wait_for_new_window(wins_before, timeout=30)
    if new_win:
        driver.switch_to.window(new_win)
        log("   ✅ Switched to Analytics tab")
    time.sleep(3)


def _click_reports_in_context(browser, depth=0):
    """Find a visible Reports control, including inside embedded frames."""
    label = "translate(normalize-space(.), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz')"
    xpath = (
        "//*[@id='REPORTS'] | "
        "//*[self::a or self::button or self::span or @role='tab' or @role='button']"
        f"[{label}='reports']"
    )
    for element in browser.find_elements(By.XPATH, xpath):
        try:
            if element.is_displayed() and element.is_enabled():
                element.click()
                return True
        except WebDriverException:
            continue
    if depth < 3:
        for frame in browser.find_elements(By.CSS_SELECTOR, "iframe, frame"):
            entered = False
            found = False
            try:
                browser.switch_to.frame(frame)
                entered = True
                found = _click_reports_in_context(browser, depth + 1)
                if found:
                    return True
            except WebDriverException:
                continue
            finally:
                if entered and not found:
                    browser.switch_to.parent_frame()
    return False


def go_to_reports():
    log("\n=== REPORTS ===")
    initial_window = driver.current_window_handle

    def open_reports(browser):
        # Recheck handles on each poll: Analytics may open its tab late.
        handles = list(browser.window_handles)
        handles.sort(key=lambda handle: handle != initial_window)
        for handle in handles:
            try:
                browser.switch_to.window(handle)
                browser.switch_to.default_content()
                if _click_reports_in_context(browser):
                    return True
            except WebDriverException:
                continue
        return False

    try:
        WebDriverWait(driver, 60).until(open_reports)
    except TimeoutException as exc:
        screenshot = os.path.join(os.path.dirname(os.path.abspath(__file__)), "reports_timeout.png")
        try:
            if driver.save_screenshot(screenshot):
                log(f"Reports timeout screenshot saved: {screenshot}")
        except WebDriverException:
            pass
        raise RuntimeError(
            "Reports was not clickable in any open tab or checked frame after 60 seconds. "
            "Check reports_timeout.png to see whether Analytics loaded or login is required."
        ) from exc
    log("Reports opened")
    time.sleep(3)


def go_to_scheduled_visits():
    log("\n=== SCHEDULED VISITS ===")
    try:
        wait_click(By.XPATH,
            "//a[contains(text(),'Scheduled Visits')] | //span[contains(text(),'Scheduled Visits')]",
            timeout=10, desc="Scheduled Visits")
    except:
        wait_click(By.ID, "yui_3_1_1773261475454_281", timeout=10, desc="Scheduled Visits (ID)")
    time.sleep(4)


# ── Date Range ─────────────────────────────────────────────────
def set_date_range(start_date, end_date):
    log(f"\n=== DATE RANGE: {start_date} → {end_date} ===")
    try:
        wait_click(By.CSS_SELECTOR, "#divDateRangeSelect > span", desc="Date Range dropdown")
    except:
        wait_click(By.XPATH, "//*[contains(@id,'DateRange')]", desc="Date Range dropdown")
    time.sleep(2)
    try:
        wait_click(By.CSS_SELECTOR, ".ranges li:nth-child(5)", desc="Custom Range")
    except:
        wait_click(By.XPATH, "//*[contains(text(),'Custom')]", desc="Custom Range")
    time.sleep(2)
    try:
        sf = wait_visible(By.CSS_SELECTOR, "input[name='daterangepicker_start']", desc="Start date")
        clear_type(sf, start_date)
        sf.send_keys(Keys.TAB)
        log(f"   ✅ Start: {start_date}")
    except Exception as e:
        log(f"   ⚠️ Start date error: {e}")
    time.sleep(1)
    try:
        ef = wait_visible(By.CSS_SELECTOR, "input[name='daterangepicker_end']", desc="End date")
        clear_type(ef, end_date)
        ef.send_keys(Keys.TAB)
        log(f"   ✅ End: {end_date}")
    except Exception as e:
        log(f"   ⚠️ End date error: {e}")
    time.sleep(1)
    try:
        wait_click(By.CSS_SELECTOR, ".applyBtn", desc="Apply")
    except:
        wait_click(By.XPATH, "//button[contains(text(),'Apply')]", desc="Apply")
    time.sleep(5)
    log("✅ Date range applied")


# ── Options ────────────────────────────────────────────────────
def configure_options():
    log("\n=== OPTIONS (Patient ID) ===")
    try:
        wait_click(By.ID, "OptionsBtn", desc="Options")
    except:
        wait_click(By.XPATH, "//button[contains(text(),'Options')]", desc="Options")
    time.sleep(2)
    try:
        wait_click(By.CSS_SELECTOR, "td:nth-child(1) > ul > li:nth-child(4) span", desc="Patient ID")
    except:
        try:
            wait_click(By.XPATH, "//span[contains(text(),'Patient ID')]", desc="Patient ID")
        except:
            log("   ⚠️ Patient ID checkbox not found")
    time.sleep(1)
    try:
        wait_click(By.ID, "lblLayoutOk", desc="OK")
    except:
        wait_click(By.XPATH, "//button[contains(text(),'OK')]", desc="OK")
    time.sleep(4)
    log("✅ Options configured")


# ── Export CSV ─────────────────────────────────────────────────
def export_csv():
    log(f"\n=== EXPORTING CSV ===")
    try:
        wait_click(By.ID, "ExportDataBtn", desc="Export Data button")
    except:
        wait_click(By.XPATH, "//span[contains(text(),'Export Data')]", desc="Export Data")
    time.sleep(2)
    try:
        wait_click(By.ID, "lblExportCsv_rdPopupOptionItem", desc="CSV option")
    except:
        wait_click(By.XPATH, "//span[text()='CSV']", desc="CSV option")
    log("   ⏳ Waiting for CSV download...")
    time.sleep(8)
    csv_file = wait_for_download(DOWNLOAD_DIR, ".csv", timeout=30)
    if csv_file:
        ts       = datetime.now().strftime("%Y%m%d_%H%M%S")
        new_name = f"scheduled_visits_{DATE_LABEL}_{ts}.csv"
        dest     = os.path.join(DOWNLOAD_DIR, new_name)
        shutil.move(csv_file, dest)
        log(f"✅ CSV saved: {dest}")
        return dest
    else:
        log("   ⚠️ CSV download not detected")
        return None


def wait_for_download(directory, extension, timeout=30):
    end = time.time() + timeout
    while time.time() < end:
        files = [
            os.path.join(directory, f)
            for f in os.listdir(directory)
            if f.endswith(extension) and not f.endswith(".crdownload")
        ]
        if files:
            return max(files, key=os.path.getmtime)
        time.sleep(1)
    return None


# ── MAIN ───────────────────────────────────────────────────────
def main():
    global driver
    validate_configuration()
    compute_dates()                     # <-- set DATE_* to TODAY (New York) for this run
    log("=" * 60)
    log(f"  WebPT → Google Sheets  |  Date: {DATE_LABEL}  (Today, New York)")
    log("  Report tab: Upcoming")
    log("=" * 60)

    master_csv_path = None
    try:
        open_browser()
        login()
        go_to_analytics()
        go_to_reports()
        go_to_scheduled_visits()
        upcoming_end = datetime.strptime(DATE_END, "%m/%d/%Y").replace(month=12, day=31).strftime("%m/%d/%Y")
        set_date_range(DATE_START, upcoming_end)
        configure_options()
        master_csv_path = export_csv()
    finally:
        if driver:
            time.sleep(2)
            driver.quit()
            driver = None
            log("🔒 Browser closed")

    if not master_csv_path:
        raise RuntimeError("Failed to download the visits CSV")

    with open(master_csv_path, "r", encoding="utf-8-sig") as f:
        all_rows = list(csv.reader(f))

    if not all_rows:
        raise RuntimeError("The downloaded visits CSV is empty")

    log(f"\n📋 Total rows in CSV: {len(all_rows) - 1} (+ header)")

    log("\n── Refresh Upcoming ──")
    write_google_sheet_tabs(all_rows)

    log("\n✅ ALL DONE!")
    log("📊 Upcoming saved successfully")


def run_daily_at_0430():
    """Keep the process alive and run main once daily at 04:30 New York time."""
    try:
        from zoneinfo import ZoneInfo
        schedule_tz = ZoneInfo("America/New_York")
    except Exception:
        schedule_tz = None

    while True:
        now = datetime.now(schedule_tz) if schedule_tz else datetime.now()
        target = now.replace(hour=4, minute=30, second=0, microsecond=0)
        if target <= now:
            target += timedelta(days=1)

        wait_seconds = (target - now).total_seconds()
        log(f"Next run scheduled for {target.isoformat()} ({wait_seconds / 3600:.2f} hours)")
        time.sleep(max(wait_seconds, 1))

        try:
            main()
        except Exception as exc:
            log(f"Scheduled run failed: {exc}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--test-excel", nargs="?", const="visits_local_test.xlsx", metavar="PATH",
        help="Write a local Excel workbook using sample data or --input-csv.",
    )
    parser.add_argument("--input-csv", help="WebPT CSV to use for the local Excel test.")
    args = parser.parse_args()
    if args.input_csv and not args.test_excel:
        parser.error("--input-csv requires --test-excel")
    if args.test_excel:
        test_local_excel(args.test_excel, args.input_csv)
    elif os.getenv("RUN_SCHEDULED", "false").lower() in {"1", "true", "yes"}:
        run_daily_at_0430()
    else:
        main()
