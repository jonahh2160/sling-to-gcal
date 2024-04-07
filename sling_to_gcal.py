# sling_to_gcal v0.4 - jonahh2160 4/7/2024

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
    
# Function to create all shifts from the .xlsx file
def create_event(date, time_range, description, employee):
    # Convert date
    month, day = convert_date(date)
    
    # Convert time_range
    start_hr, start_min, end_hr, end_min = convert_time(time_range)
    
    # Convert description
    position, location = convert_desc(description)
    
    # Set up corresponding event dict
    event = {
        "summary": position,
        "location": location,
        "description": f"Employee: {employee}",
        "start": {
            "dateTime": f"{year}-{month}-{day}T{start_hr}:{start_min}:00"
        },
        "end": {
            "dateTime": f"{year}-{month}-{day}T{end_hr}:{end_min}:00"
        }
    }
    
    # Add the event to the list of shift events
    shift_events.append(event)

# Try catch block to contain rest of the script's logic
try:
    # Open the shifts file and get its only worksheet
    wb = openpyxl.load_workbook(filename='shifts-export.xlsx')
    ws = wb.active
    print("Opened shifts-export.xlsx")

    # Set up Google Calendar
    service = build("calendar", "v3", credentials=creds)
    print("Set up the connection to Calendar")
    
    # Check the year
    print("Grabbing the current year")
    year = datetime.datetime.now().year

    # Find the ranges of the shifts to check
    print("Checking range of shifts")
    # Find the max date by checking the last row right to left
    for col in range(ws.max_column+1, 1, -1):
        if ws.cell(row=ws.max_row, column=col).value is not None:
            max_date = ws.cell(row=ws.max_row, column=col).value
            max_col = col
            break
    # Parse and convert to datetime
    max_date = max_date.split('\n')[0]
    max_date = f"{year} {convert_date(max_date)}"
    max_date = max_date.replace(",", "").replace("(", "").replace(")", "").replace("'", "")
    max_date = datetime.datetime.strptime(max_date, "%Y %m %d")
    # Then find the min date by offset
    day_offset = (7 * (ws.max_row - 2)) + max_col - 1
    min_date = (max_date - datetime.timedelta(days = day_offset)).isoformat()
    max_date = (max_date.replace(hour=23, minute=59, second=59)).isoformat() 
    
    # Use the acquired ranges to get a list of the current shifts
    print("Getting existing shifts from Calendar")
    # TODO: Find a way to decide what the keyword and the timezone (-07:00) should be
    keyword = "Clerk"
    events_result = service.events().list(calendarId='primary', q="Clerk", timeMin=f"{min_date}-07:00", timeMax=f"{max_date}-07:00").execute()
    current_events = events_result.get('items', [])
    shift_events = []

    # Loop through every cell, excluding the first row
    print("Processing files...")
    print()
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
                    
                    # TODO: Remember to look into offsetting it by ten minutes
                    create_event(date, time_range, description, employee)
    
    # Compare shift_events and current_events
    # TODO: Is it possible to move this block back into create_event function?
    # TODO: If events exist in current events but not in shift events, delete
    # TODO: If an event exists in both, update
    # TODO: Else, create a new event and print('Created an event: %s' % (event.get('htmlLink')))
    
    print()
    input("File finished. Press <ENTER> to exit.")
                
# Error handling
except FileNotFoundError:
    print()
    print("File not found! Was the shifts-export file removed? Exiting...")
except HttpError as e:
    print()
    print("An error occurred:", e, end="\n\n")
    print("Exiting...")
except Exception as e:
    print()
    print("An error occurred:", e, end="\n\n")
    print("Exiting...")
