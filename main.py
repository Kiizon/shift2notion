import pandas as pd
import os.path
import os, base64, re
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from datetime import datetime, timedelta


SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/calendar"
]
# CONFIGURATION: Change this to your name as it appears in the Excel schedule
name = "Kish"
days_of_week = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

def build_gmail(creds):
    """
    Builds the Gmail service
    """
    return build("gmail", "v1", credentials=creds)

def gmail_search_ids(gmail, query, max_results=10):
    resp = gmail.users().messages().list(userId="me", q=query, maxResults=max_results).execute()
    return [m["id"] for m in resp.get("messages", [])]

def _iter_parts(payload):
    if not payload:
        return
    if payload.get("parts"):
        for p in payload["parts"]:
            yield from _iter_parts(p)
    else:
        yield payload
def gmail_download_first_excel(gmail, query, save_dir="./downloads"):
    """
    Returns the saved file path (str) or None if not found.
    """
    os.makedirs(save_dir, exist_ok=True)
    ids = gmail_search_ids(gmail, query, max_results=20)
    for mid in ids or []:
        msg = gmail.users().messages().get(userId="me", id=mid, format="full").execute()
        payload = msg.get("payload", {})

        headers = {h["name"].lower(): h["value"] for h in payload.get("headers", [])}
        sender = headers.get("from", "")

        for part in _iter_parts(payload):
            filename = part.get("filename") or ""
            if not filename:
                continue
            if not re.search(r"\.(xlsx|xls)$", filename, re.I):
                continue

            body = part.get("body", {})
            data = body.get("data")
            attachment_id = body.get("attachmentId")

            if data: 
                raw = base64.urlsafe_b64decode(data)
            elif attachment_id:
                att = gmail.users().messages().attachments().get(
                    userId="me", messageId=mid, id=attachment_id
                ).execute()
                raw = base64.urlsafe_b64decode(att["data"])
            else:
                continue

            # safe filename + timestamp
            base, ext = os.path.splitext(filename)
            safe = "".join(c for c in base if c.isalnum() or c in ("-","_"))[:80]
            out_path = os.path.join(save_dir, f"{safe}{ext}")
            with open(out_path, "wb") as f:
                f.write(raw)
            return out_path
    return None
def get_shifts(file_path: str) -> list:
    """
    Find the name in the Excel file and return the row number
    """
    try:
        shifts = []
        df = pd.read_excel(file_path, header=None)
        
        for idx, value in enumerate(df.iloc[:, 1]):
            if value == name:
                for i, day in enumerate(days_of_week):
                    am_col_idx = 2 + (i * 2)  # AM shift column index (even)
                    pm_col_idx = 3 + (i * 2)  # PM shift column index (odd)
        
                    am_shift = df.iloc[idx, am_col_idx]  # AM shift value
                    pm_shift = df.iloc[idx, pm_col_idx]  # PM shift value
                    
                    if pd.notna(am_shift):
                        shifts.append(f"{day} {am_shift}")
                    if pd.notna(pm_shift):
                        shifts.append(f"{day} {pm_shift}")
        formatted_shifts = parse_shifts(shifts)
        return formatted_shifts
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def parse_shifts(shifts: list) -> list[tuple[str, str, str]]:
    """
    Accepts a string of shifts and returns the start and end times

    ['tuesday 6-CL', 'wednesday 5:30-CL', 'thursday 6-CL']

    """
    formatted_shifts = []
    if not shifts:
        print("No shifts found")
        return None, None
    

    for shift in shifts:
        parts = shift.split()
        if len(parts) != 2:
            return None, None
    
        day = parts[0].strip()  # "tuesday"
        shift_time = parts[1].strip()  # "6-CL" or "5:30-CL"
    
        if "-CL" in shift_time:
            # Extract the time part before "-CL"
            time_part = shift_time.split("-CL")[0]
            
            # Handle format like "5:30" - convert to 24-hour format
            if ":" in time_part:
                hour, minute = time_part.split(":")
                hour_24 = int(hour) + 12  # Convert to 24-hour format (assuming PM)
                start_time = f"{hour_24}:{minute}"
                end_time = "23:59"
                formatted_shifts.append((day, start_time, end_time))

    return formatted_shifts

def build_events(formatted_shifts):
    """
    Builds events from the formatted shifts
    """

    events = []
    for shift in formatted_shifts:
        day, start, end = shift

        base_date = datetime.now()

        shift_date = get_actual_date(base_date,day) # get the actual date of the shift

        start_time = datetime.strptime(start, "%H:%M").time()
        end_time = datetime.strptime(end, "%H:%M").time()

        start_dt = datetime.combine(shift_date.date(), start_time).isoformat()
        end_dt = datetime.combine(shift_date.date(), end_time).isoformat()
        event = {
            "summary": "Work Shift",
            "description": "Work Shift",
            "colorId": 1,
            "start": {
                "dateTime": start_dt + "-04:00"
            },
            "end": {
                "dateTime": end_dt + "-04:00"
            }
        }
        events.append(event)
    return events
    

def get_actual_date(base_date, day):
    """
    Get the actual date of the shift
    """
    days_of_week = {
        "monday": 0,
        "tuesday": 1,
        "wednesday": 2,
        "thursday": 3,
        "friday": 4,
        "saturday": 5,
        "sunday": 6
    }
    target = days_of_week[day.lower()]
    base_weekday = base_date.weekday()
    # Find offset to next occurrence
    days_ahead = (target - base_weekday) % 7
    return base_date + timedelta(days=days_ahead)


def main():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    try:
        gmail = build_gmail(creds)
        calendar = build("calendar", "v3", credentials=creds)

        # CONFIGURATION: Change this to your employer's email address
        employer_email = "your.employer@company.com"
        query = f'from:{employer_email} newer_than:6d (filename:xls OR filename:xlsx)'


        file_path = gmail_download_first_excel(gmail, query, save_dir="./downloads")
        print(file_path)
        events = build_events(get_shifts(file_path))
        for event in events:
            event = calendar.events().insert(calendarId="primary", body=event).execute()
            print("--------------------------------Adding event to calendar--------------------------------")
            print("Event created: ", event.get("htmlLink"))

    except HttpError as error:
        print(f"An error occurred: {error}")



if __name__ == "__main__":
    file_path = "pm test.xls"
    main()