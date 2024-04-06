# sling_to_gcal v0.1 - jonahh2160 4/6/2024

# Import packages
import os, os.path, datetime, openpyxl
from stg_functions import process_event
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Scope includes reading and writing to calendar
SCOPES = ["https://www.googleapis.com/auth/calendar"]

# From Google's Python quickstart guide; handles authentication
creds = None
if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

# Change the shift file from .xls to .xlsx (openpyxl only handles .xlsx)
if os.path.exists('shifts-export.xls'):
    os.rename('shifts-export.xls', 'shifts-export.xlsx')

# Try catch block to contain rest of the script's logic
try:
    # Open the shifts file and get its only worksheet
    wb = openpyxl.load_workbook(filename='shifts-export.xlsx')
    ws = wb.active

    # Set up Google Calendar
    service = build("calendar", "v3", credentials=creds)

    # Loop through every cell, excluding the first row
    for row in range(2, ws.max_row+1):
        for col in range(1, ws.max_column+1):
            # Get the value in the current cell if it has an event
            if ws.cell(row=row, column=col).value is not None:
                cell_value = ws.cell(row=row, column=col).value
                
                print()
                print(f"-=-=-=-=- CELL ({row}.{col}) -=-=-=-=-")
                print(cell_value)   
                print()
                
                # TODO: Get the date, the time_range, the description, and employee
                # TODO: If there are multiple events for the same day, pass the same day variable
                # Process events
                process_event(date, time_range, description, employee)
                
except FileNotFoundError:
    print("File not found! Exiting...")
except HttpError as e:
    print("An error occurred:", e)
except Exception as e:
    print("An error occurred:", e)
