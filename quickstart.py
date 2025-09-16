import datetime
import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from main import build_events

SCOPES = ["https://www.googleapis.com/auth/calendar"] # read and write access

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

    SHIFTS = build_events(file_path)

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


if __name__ == "__main__":
  main()
  file_path = "schedule.xls"
  name = build_events(file_path)