# sling_to_gcal v0.3 - jonahh2160 4/6/2024

# Import packages
import os, os.path, datetime, openpyxl
from stg_conversions import *
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
    print("Opened token.json")
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      print("No valid credentials found, letting user log in...")
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())
    print("Saved credentials")

# Change the shift file from .xls to .xlsx (openpyxl only handles .xlsx)
if os.path.exists('shifts-export.xls'):
    os.rename('shifts-export.xls', 'shifts-export.xlsx')
    print("Changed shifts-export's file extention")
    
# Function to sync any given shift from the .xlsx file
def process_event(date, time_range, description, employee):
    # Process date
    year = datetime.datetime.now().year
    month, day = convert_date(date)
    
    # Process time_range
    start_hr, start_min, end_hr, end_min = convert_time(time_range)
    
    # Process description
    position, location = convert_desc(description)
    
    # TODO: Check if event already exists
    # If no event exists, then create one
    event = {
        "summary": position,
        "location": location,
        "description": f"Employee: {employee}",
        "start": {
            "dateTime": f"{year}-{month}-{day}T{start_hr}:{start_min}:00"
        },
        "end": {
            "dateTime": f"{year}-{month}-{day}T{end_hr}:{end_min}:00"
        },
        "reminders": {
            "useDefault": True,
        }
    }
    
    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Created an event: %s' % (event.get('htmlLink')))

# Try catch block to contain rest of the script's logic
try:
    # Open the shifts file and get its only worksheet
    wb = openpyxl.load_workbook(filename='shifts-export.xlsx')
    ws = wb.active
    print("Opened shifts-export.xlsx")

    # Set up Google Calendar
    service = build("calendar", "v3", credentials=creds)
    print("Set up the connection to Calendar")

    print("Processing shifts...")
    print()

    # Loop through every cell, excluding the first row
    for row in range(2, ws.max_row+1):
        for col in range(1, ws.max_column+1):
            # Get the value in the current cell if it has an event
            if ws.cell(row=row, column=col).value is not None:
                cell_value = ws.cell(row=row, column=col).value

                # First line is always the date
                date = cell_value.split('\n')[0]
                
                # Shifts are seperated by double newlines
                shifts = cell_value.split('\n\n')[1:]
                
                # Process each shift in the cell
                for shift in shifts:
                    # Split the string up line by line
                    lines = shift.split('\n')
                    
                    time_range = lines[0]
                    description = lines[1]
                    employee = lines[2]
                    
                    """
                    print()
                    print("-=-=-=-=-=-=-=-=-=-=-=-=-=-=-")
                    print(shift)
                    print()
                    print("Date:", date, convert_date(date))
                    print("Time:", time_range, convert_time(time_range))
                    print("Description:", description, convert_desc(description))
                    print("Employee:", employee)
                    print()
                    """
                    
                    # process_event(date, time_range, description, employee)
    
    print()
    input("File finished. Press <ENTER> to exit.")
                
except FileNotFoundError:
    print("File not found! Was the shifts-export file removed?")
    print("Exiting...")
except HttpError as e:
    print("An error occurred:", e)
    print("Exiting...")
except Exception as e:
    print("An error occurred:", e)
    print("Exiting...")
