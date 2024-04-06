# Contains functions for the main script to use - jonahh2160 4/6/2024 

# Import packages
import os, os.path, datetime, openpyxl; from stg_functions import *
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Convert the time line in a shift into variables for process_event
def convert_time(time_range):
    # Parse the string and set variables based on whitespace
    time_range = time_range.replace(":", "").replace("-", "")
    start_time, start_ampm, end_time, end_ampm = time_range.split()

    # Add a 0 if the hour is only one digit to get a 0000 format
    if len(start_time) == 3:
        start_time = "0" + start_time
    if len(end_time) == 3:
        end_time = "0" + end_time

    # Slice the time strings by hour and minute
    start_hr = start_time[:2]
    end_hr = end_time[:2]
    start_min = start_time[2:]
    end_min = end_time[2:]
        
    # Convert each time from 12-hour to military
    if start_ampm == "PM" and start_hr != "12":
        start_hr = str(int(start_hr) + 12)
    elif start_ampm == "AM" and start_hr == "12":
        start_hr = "00"
    if end_ampm == "PM" and end_hr != "12":
        end_hr = str(int(end_hr) + 12)
    elif end_ampm == "AM" and end_hr == "12":
        end_hr = "00"
        
    return start_hr, start_min, end_hr, end_min

# Function to sync any given shift from the .xlsx file
def process_event(date, time_range, description, employee):
    # TODO: Process date
    
    year = datetime.datetime.now().year
    # TODO: Process time_range
    
    # TODO: Process description
    
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
    print('Event created: %s' % (event.get('htmlLink')))
