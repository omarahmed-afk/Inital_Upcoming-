# Upcoming

Export WebPT appointments from today through December 31 to the `Upcoming` Google Sheets tab. Each run replaces this tab's results; other tabs are not modified.

One row is selected per Patient ID, preferring Other, Confirmed, Checked Out, then Checked In. The earliest appointment within the preferred status is selected. Dates use America/New_York time. Existing clinic and appointment exclusions apply.

## Setup

Install Python 3.12 and Chrome, then install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Enable the Google Sheets and Google Drive APIs and share the spreadsheet with your service account as Editor. Keep its JSON key outside Git.

Set these environment variables locally:

- `WEBPT_USERNAME`
- `WEBPT_PASSWORD`
- `GOOGLE_SHEET_ID`
- `GOOGLE_SERVICE_ACCOUNT_FILE` (path to your service account JSON key)

The exported report must include Patient ID, Clinic Name, Patient Name, Treating Therapist, Case Therapist, Appointment Type, Appointment Date, and Visit Status.

## Run

```powershell
python -X utf8 Upcoming.py
```

For a local sample without accessing WebPT or Google Sheets:

```powershell
python -X utf8 Upcoming.py --test-excel upcoming_sample.xlsx
```

Set `RUN_SCHEDULED=true` to keep the process running and execute daily at 04:30 New York time.

## GitHub Actions

`.github/workflows/upcoming.yml` runs manually only, from Actions > Update Upcoming > Run workflow. No automatic schedule is configured.

Before running, add repository secrets under Settings > Secrets and variables > Actions:

- `WEBPT_USERNAME`
- `WEBPT_PASSWORD`
- `GOOGLE_SHEET_ID`
- `GOOGLE_SERVICE_ACCOUNT_JSON`: the full contents of your service account JSON key.

The workflow runs Chrome headlessly and updates only Upcoming. It does not upload exported patient data or screenshots as artifacts.
