"""
Automated Scheduling and Payroll Forecasting Bot.
This script extracts shift data from a corporate portal using Playwright,
syncs the shifts to Google Calendar, and maintains an up-to-date payroll 
forecast in Google Sheets.
"""

import os
from datetime import datetime, timedelta
from typing import List, Dict, Tuple, Optional, Any

from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- FINANCIAL CONSTANTS & PAYROLL RATES ---
BASE_HOURLY_RATE = 7.77    
NIGHT_SHIFT_RATE = 1.76
HOLIDAY_RATE = 3.24
SPLIT_SHIFT_BONUS = 20.90
DAILY_TRANSPORT_BONUS = 6.18     

# --- GOOGLE API CONFIGURATION ---
SCOPES = [
    'https://www.googleapis.com/auth/calendar',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def authenticate_google_services() -> Tuple[Any, Any, Any]:
    """
    Authenticates and builds Google Cloud services (Calendar, Sheets, Drive).
    Uses OAuth 2.0 flow and manages token refresh automatically.

    Returns:
        Tuple containing initialized Google API resource objects 
        (calendar_service, sheets_service, drive_service).
    """
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    return (
        build('calendar', 'v3', credentials=creds), 
        build('sheets', 'v4', credentials=creds), 
        build('drive', 'v3', credentials=creds)
    )

def extract_shifts_from_portal() -> Optional[List[Dict[str, Any]]]:
    """
    Navigates the corporate portal via Playwright to extract shift schedules.
    Filters out unassigned or unpublished data.

    Returns:
        List of dictionaries containing parsed shift details, or None if extraction fails.
    """
    raw_shifts = []
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        try:
            # Portal Authentication
            page.goto("https://kiosko.groundforce.aero/EmployeeKiosk/Schedule.aspx?tab=2")
            page.fill("input#tbUsername", os.getenv("GF_USERNAME", "")) 
            page.fill("input#tbPassword", os.getenv("GF_PASSWORD", "")) 
            page.click("input#btSignIn")
            page.wait_for_load_state("networkidle")
            
            # Navigate to Schedule Tab
            page.goto("https://kiosko.groundforce.aero/EmployeeKiosk/Schedule.aspx?tab=2")
            page.wait_for_selector("#dgSchedule_lblTimeDateFrom_0", timeout=20000)

            # Data Extraction Loop
            index = 0
            while True:
                row_selector = f"#dgSchedule_lblTimeDateFrom_{index}"
                if page.locator(row_selector).count() == 0: 
                    break
                
                date_str = page.locator(row_selector).inner_text().strip()
                state = page.locator(f"#dgSchedule_lblState_{index}").inner_text().strip()
                shift_type = page.locator(f"#dgSchedule_lblType_{index}").inner_text().strip()
                
                try:
                    start_time = page.locator(f"#dgSchedule_lblTimeFrom_{index}").inner_text().strip()
                    end_time = page.locator(f"#dgSchedule_lblTimeTo_{index}").inner_text().strip()
                except Exception:
                    start_time, end_time = "", ""

                shift_data = {
                    "date": date_str, 
                    "start": start_time, 
                    "end": end_time, 
                    "state": state, 
                    "is_day_off": False, 
                    "title": ""
                }
                
                # Parse shift types based on Spanish portal outputs
                if "Turnos" in shift_type: 
                    shift_data["title"] = f"✈️ {start_time}-{end_time}"
                elif "Permiso (CU)" in shift_type: 
                    shift_data["title"] = f"📚 Course {start_time}-{end_time}"
                elif any(x in shift_type for x in ["Libre", "Descanso", "Permiso (F)", "Permiso (V)", "Permiso (DCF)", "Permiso (DCH)", "Permiso (AJ)"]):
                    shift_data["is_day_off"] = True
                    shift_data["title"] = f"🌴 {shift_type.replace('Permiso ', '')}"

                if shift_data["title"]: 
                    raw_shifts.append(shift_data)
                index += 1

            # Filter valid week range (Up to the current published Sunday)
            latest_published_date = None
            for row in raw_shifts:
                if "Publicado" in row['state']:
                    dt = datetime.strptime(row['date'], "%d/%m/%Y")
                    if not latest_published_date or dt > latest_published_date: 
                        latest_published_date = dt
                        
            if latest_published_date:
                week_end_boundary = latest_published_date + timedelta(days=(6 - latest_published_date.weekday()))
                raw_shifts = [r for r in raw_shifts if datetime.strptime(r['date'], "%d/%m/%Y") <= week_end_boundary]
            
            return raw_shifts
            
        except Exception:
            return None
        finally: 
            browser.close()

def calculate_night_hours(start_dt: datetime, end_dt: datetime) -> float:
    """
    Calculates the total night hours accrued within a specific shift.
    Night rules apply between 21:00 and 08:00.

    Args:
        start_dt (datetime): Shift start timestamp.
        end_dt (datetime): Shift end timestamp.

    Returns:
        float: Total night hours to be compensated.
    """
    total_duration = (end_dt - start_dt).total_seconds() / 3600
    night_hours = 0.0
    temp_dt = start_dt
    
    while temp_dt < end_dt:
        if temp_dt.hour >= 21 or temp_dt.hour < 8: 
            night_hours += 0.25
        temp_dt += timedelta(minutes=15)
        
    return total_duration if night_hours >= 4 else night_hours

def sync_to_google_calendar(calendar_service: Any, shifts: List[Dict[str, Any]]) -> None:
    """
    Upserts shift data into Google Calendar, avoiding duplicates.
    Assigns specific colors to differentiate working shifts from days off.

    Args:
        calendar_service: Authenticated Google Calendar API instance.
        shifts: List of parsed shift dictionaries.
    """
    timezone = 'Europe/Madrid'
    tracked_event_ids = []
    processed_dates = []
    
    for shift in shifts:
        dt_obj = datetime.strptime(shift['date'], "%d/%m/%Y")
        processed_dates.append(dt_obj)
        
        index = 0
        while f"gf{dt_obj.strftime('%Y%m%d')}{index}" in tracked_event_ids: 
            index += 1
        event_id = f"gf{dt_obj.strftime('%Y%m%d')}{index}"
        tracked_event_ids.append(event_id)
        
        event_body = {
            'id': event_id, 
            'summary': shift['title'], 
            'description': 'Automated sync by Automator Kiosko.'
        }
        
        if shift['is_day_off']:
            event_body['start'] = {'date': dt_obj.strftime("%Y-%m-%d")}
            event_body['end'] = {'date': (dt_obj + timedelta(days=1)).strftime("%Y-%m-%d")}
            event_body['colorId'] = '2' # Green for days off
        else:
            start_dt = datetime.strptime(f"{shift['date']} {shift['start']}", "%d/%m/%Y %H:%M")
            end_dt = datetime.strptime(f"{shift['date']} {shift['end']}", "%d/%m/%Y %H:%M")
            if end_dt < start_dt: 
                end_dt += timedelta(days=1)
                
            event_body['start'] = {'dateTime': start_dt.isoformat(), 'timeZone': timezone}
            event_body['end'] = {'dateTime': end_dt.isoformat(), 'timeZone': timezone}
            event_body['colorId'] = '9' # Blueberry for shifts

        try:
            calendar_service.events().insert(calendarId='primary', body=event_body).execute()
        except HttpError as e:
            if e.resp.status == 409: # Conflict, event already exists -> update
                calendar_service.events().update(calendarId='primary', eventId=event_id, body=event_body).execute()

    # Clean up obsolete synced events within the timeframe
    if processed_dates:
        time_min = min(processed_dates).isoformat() + 'Z'
        time_max = (max(processed_dates) + timedelta(days=60)).isoformat() + 'Z'
        existing_events = calendar_service.events().list(
            calendarId='primary', timeMin=time_min, timeMax=time_max, q='Automated sync'
        ).execute().get('items', [])
        
        for event in existing_events:
            if event['id'].startswith('gf') and event['id'] not in tracked_event_ids:
                calendar_service.events().delete(calendarId='primary', eventId=event['id']).execute()

def get_or_create_monthly_spreadsheet(drive_service: Any, sheets_service: Any, month_name: str) -> Tuple[str, str]:
    """
    Locates or initializes the main directory and the specific monthly spreadsheet.

    Args:
        drive_service: Authenticated Google Drive API instance.
        sheets_service: Authenticated Google Sheets API instance.
        month_name (str): Expected name of the file (e.g., 'March 2026').

    Returns:
        Tuple containing (spreadsheet_id, folder_id).
    """
    folder_query = "name = 'Schedule Forecasts' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    folders = drive_service.files().list(q=folder_query).execute().get('files', [])
    
    if not folders:
        folder_metadata = {'name': 'Schedule Forecasts', 'mimeType': 'application/vnd.google-apps.folder'}
        folder_id = drive_service.files().create(body=folder_metadata, fields='id').execute().get('id')
    else:
        folder_id = folders[0]['id']

    file_query = f"name = '{month_name}' and '{folder_id}' in parents and trashed = false"
    files = drive_service.files().list(q=file_query).execute().get('files', [])
    
    if not files:
        file_metadata = {
            'name': month_name, 
            'mimeType': 'application/vnd.google-apps.spreadsheet', 
            'parents': [folder_id]
        }
        spreadsheet_id = drive_service.files().create(body=file_metadata, fields='id').execute().get('id')
        return spreadsheet_id, folder_id
        
    return files[0]['id'], folder_id

def get_previous_month_variable_pay(drive_service: Any, sheets_service: Any, folder_id: str, current_month_dt: datetime) -> float:
    """
    Retrieves the generated variable income from the preceding month's spreadsheet.

    Args:
        drive_service: Authenticated Google Drive API instance.
        sheets_service: Authenticated Google Sheets API instance.
        folder_id: Parent Google Drive folder ID.
        current_month_dt: Reference date for the current processing context.

    Returns:
        float: Variable pay accrued in the previous month.
    """
    previous_month_dt = (current_month_dt.replace(day=1) - timedelta(days=1))
    previous_month_name = previous_month_dt.strftime("%B %Y")
    
    file_query = f"name = '{previous_month_name}' and '{folder_id}' in parents and trashed = false"
    files = drive_service.files().list(q=file_query).execute().get('files', [])
    
    if not files: 
        return 0.0 
    
    try:
        spreadsheet_id = files[0]['id']
        sheet_data = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range="A:J"
        ).execute().get('values', [])
        
        # Scan backwards to locate the totals footer
        for row in reversed(sheet_data): 
            if len(row) >= 9 and "VARIABLES (Generated):" in row:
                return float(str(row[8]).replace('€', '').replace(',', '.').strip())
    except Exception:
        return 0.0
        
    return 0.0

def update_spreadsheet_data(sheets_service: Any, drive_service: Any, folder_id: str, spreadsheet_id: str, shifts: List[Dict[str, Any]], reference_dt: datetime) -> None:
    """
    Processes financial logic, preserves historical manual entries, and 
    pushes the final formatted grid to Google Sheets.

    Args:
        sheets_service: Authenticated Google Sheets API instance.
        drive_service: Authenticated Google Drive API instance.
        folder_id: Parent Google Drive folder ID.
        spreadsheet_id: Target Google Sheet ID.
        shifts: Current payload of shifts for this specific month.
        reference_dt: Anchor date to fetch previous month data.
    """
    try:
        existing_data = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range="A2:J").execute().get('values', [])
    except Exception:
        existing_data = []

    # Preserve historical data locally
    daily_records = {}
    for row in existing_data:
        if not row or len(row) < 3: 
            continue
        date_key = row[0]
        if "/" in date_key and len(date_key) == 10: 
            while len(row) < 10: 
                row.append(0) 
            if date_key not in daily_records: 
                daily_records[date_key] = []
            daily_records[date_key].append(row)

    holidays = ["01/01", "06/01", "28/02", "02/04", "03/04", "01/05", "19/08", "08/09", "12/10", "01/11", "06/12", "08/12", "25/12"]
    shifts_by_date = {}
    for shift in shifts:
        if shift['date'] not in shifts_by_date: 
            shifts_by_date[shift['date']] = []
        shifts_by_date[shift['date']].append(shift)

    # Process new payload
    for date_key, shift_list in shifts_by_date.items():
        is_split_shift = len([x for x in shift_list if not x['is_day_off']]) > 1
        dt_obj = datetime.strptime(date_key, "%d/%m/%Y")
        day_names = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
        day_label = day_names[dt_obj.weekday()]
        
        transport_bonus_applied = False
        daily_rows = []
        
        for index, shift in enumerate(shift_list):
            total_hrs, night_hrs, holiday_hrs, split_bonus, transport_bonus, total_variable = 0, 0, 0, 0, 0, 0
            
            if not shift['is_day_off'] and shift['start']:
                if not transport_bonus_applied: 
                    transport_bonus = DAILY_TRANSPORT_BONUS
                    transport_bonus_applied = True
                    
                start_dt = datetime.strptime(f"{date_key} {shift['start']}", "%d/%m/%Y %H:%M")
                end_dt = datetime.strptime(f"{date_key} {shift['end']}", "%d/%m/%Y %H:%M")
                if end_dt < start_dt: 
                    end_dt += timedelta(days=1)
                    
                total_hrs = (end_dt - start_dt).total_seconds() / 3600
                night_hrs = calculate_night_hours(start_dt, end_dt)
                
                if dt_obj.weekday() == 6 or date_key[:5] in holidays: 
                    holiday_hrs = total_hrs
                    
                split_bonus = SPLIT_SHIFT_BONUS if (is_split_shift and index == 0) else 0
                total_variable = (night_hrs * NIGHT_SHIFT_RATE) + (holiday_hrs * HOLIDAY_RATE) + split_bonus + transport_bonus
            
            # Rescue manually inputted overtime from historical state
            manual_overtime = 0
            if date_key in daily_records and len(daily_records[date_key]) > index:
                try: 
                    manual_overtime = float(str(daily_records[date_key][index][8]).replace(',', '.'))
                except ValueError: 
                    pass

            daily_rows.append([
                date_key, day_label, shift['title'], total_hrs, night_hrs, 
                holiday_hrs, split_bonus, transport_bonus, manual_overtime, total_variable
            ])
        
        daily_records[date_key] = daily_rows 

    # Reconstruct the grid
    sorted_dates = sorted(daily_records.keys(), key=lambda x: datetime.strptime(x, "%d/%m/%Y"))
    final_grid = []
    accumulated_variable, total_month_hrs, days_worked = 0, 0, 0
    weekly_hours_summary = {}

    for date_key in sorted_dates:
        dt_obj = datetime.strptime(date_key, "%d/%m/%Y")
        week_num = dt_obj.isocalendar()[1]
        
        if week_num not in weekly_hours_summary: 
            weekly_hours_summary[week_num] = 0

        for row in daily_records[date_key]:
            final_grid.append(row)
            try:
                total_month_hrs += float(row[3])
                accumulated_variable += float(row[9])
                weekly_hours_summary[week_num] += float(row[3])
                if float(row[7]) > 0: 
                    days_worked += 1 
            except ValueError: 
                pass

    # Build financial summary
    base_salary = total_month_hrs * BASE_HOURLY_RATE
    previous_month_variable = get_previous_month_variable_pay(drive_service, sheets_service, folder_id, reference_dt)
    estimated_net_payroll = base_salary + previous_month_variable

    final_grid.extend([
        [], 
        ["", "", "--- ACTUAL PAYROLL THIS MONTH ---", "", "", "", "", "BASE (Current Month Hrs):", round(base_salary, 2), "€"],
        ["", "", "Days Worked:", days_worked, "", "", "", "VARIABLES (From Last Month):", round(previous_month_variable, 2), "€"],
        ["", "", "Total Hours:", round(total_month_hrs, 2), "", "", "", "ESTIMATED GROSS PAYROLL:", round(estimated_net_payroll, 2), "€"],
        [],
        ["", "", "--- VARIABLES GENERATED FOR NEXT MONTH ---", "", "", "", "", "VARIABLES (Generated):", round(accumulated_variable, 2), "€"]
    ])

    headers = [["DATE", "DAY", "SHIFT", "TOTAL HRS", "NIGHT HRS", "HOLIDAY HRS", "SPLIT BONUS (€)", "TRANSPORT BONUS (€)", "MANUAL EXTRA (HRS)", "TOTAL VARIABLE (€)"]]
    
    # Write Main Matrix
    sheets_service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range="A:J").execute()
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range="A1", valueInputOption="USER_ENTERED", body={"values": headers + final_grid}
    ).execute()

    # Write Weekly Summary Matrix
    weekly_grid = [["WEEK", "TOTAL HOURS"]]
    for week, hours in weekly_hours_summary.items():
        weekly_grid.append([f"Week {week}", round(hours, 2)])
    
    sheets_service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range="L1:M50").execute()
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range="L1", valueInputOption="USER_ENTERED", body={"values": weekly_grid}
    ).execute()

def main() -> None:
    """
    Main execution pipeline. Coordinates extraction, authentication, 
    and data pushing across different Google Cloud services.
    """
    extracted_shifts = extract_shifts_from_portal()
    if extracted_shifts:
        calendar_service, sheets_service, drive_service = authenticate_google_services()
        sync_to_google_calendar(calendar_service, extracted_shifts)
        
        shifts_by_month = {}
        for shift in extracted_shifts:
            dt_ref = datetime.strptime(shift['date'], "%d/%m/%Y")
            month_label = dt_ref.strftime("%B %Y")
            if month_label not in shifts_by_month: 
                shifts_by_month[month_label] = []
            shifts_by_month[month_label].append(shift)
            
        for month_name, shift_list in shifts_by_month.items():
            spreadsheet_id, folder_id = get_or_create_monthly_spreadsheet(drive_service, sheets_service, month_name)
            reference_date = datetime.strptime(shift_list[0]['date'], "%d/%m/%Y")
            update_spreadsheet_data(sheets_service, drive_service, folder_id, spreadsheet_id, shift_list, reference_date)

if __name__ == '__main__':
    main()
