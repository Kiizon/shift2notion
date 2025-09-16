import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import pandas as pd

SCOPES = ["https://www.googleapis.com/auth/calendar"] # read and write access
name = "Kish"
days_of_week = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

def main():
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)

  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("calendar", "v3", credentials=creds)

    event = {
        "summary": "Work Shift",
        "description": "Work Shift",
        "colorId": 1,
        "start": {
            "dateTime": "2025-09-15T10:00:00-04:00"
        },
        "end": {
            "dateTime": "2025-09-15T11:00:00-04:00"
        }

    }
    
    event = service.events().insert(calendarId="primary", body=event).execute()

    print("Event created: ", event.get("htmlLink"))

  except HttpError as error:
    print(f"An error occurred: {error}")


def build_events(file_path):
    """
    Find the name in the Excel file and return the row number
    """
    try:
        df = pd.read_excel(file_path, header=None)
        
        for idx, value in enumerate(df.iloc[:, 1]):
            if value == name:
                for i, day in enumerate(days_of_week):
                    am_col_idx = 2 + (i * 2)  # AM shift column index (even)
                    pm_col_idx = 3 + (i * 2)  # PM shift column index (odd)
        
                    am_shift = df.iloc[idx, am_col_idx]  # AM shift value
                    pm_shift = df.iloc[idx, pm_col_idx]  # PM shift value
                    
                    if pd.notna(am_shift):
                        print(f"{day} AM: {am_shift}")
                    if pd.notna(pm_shift):
                        print(f"{day} PM: {pm_shift}")
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


if __name__ == "__main__":
  main()
  file_path = "pm test.xlsx"