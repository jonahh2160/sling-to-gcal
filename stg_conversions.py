# Contains functions for the main script to use - jonahh2160 4/6/2024 

# Import packages
import datetime

# Convert the date line in a shift into variables for datetime
def convert_date(date):
    # Split the string based on whitespace
    day, month = date.split()
    
    # Convert month into datetime format
    month = datetime.datetime.strptime(month, "%b").strftime("%m")

    # Convert day into datetime format
    day = datetime.datetime.strptime(day, "%d").strftime("%d")
    
    return month, day

# Convert the time line in a shift into variables for datetime
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

# Convert the description line in a shift into variables for process_event
def convert_desc(description):
    # Splint the string based on the bullet
    position, location = description.split("•")
    
    # Get rid of leading or trailing whitespace
    position = position.strip()
    location = location.strip()
    
    return position, location
