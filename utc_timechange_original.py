import pandas as pd
import pytz
from datetime import datetime, time

def convert_utc_to_client_time(utc_time, client_time_zone):
    """
    Convert UTC time to client time zone.
    """
    # Check if the time is a datetime.time object and convert it to a full datetime
    if isinstance(utc_time, time):
        # Assign a default date for conversion
        utc_time = datetime.combine(datetime(2000, 1, 1), utc_time)
    elif isinstance(utc_time, str) and len(utc_time) == 8:  # Format 'HH:MM:SS'
        utc_time = datetime.strptime(utc_time, '%H:%M:%S')
        utc_time = utc_time.replace(year=2000, month=1, day=1)
    elif isinstance(utc_time, str):
        utc_time = datetime.strptime(utc_time, '%Y-%m-%d %H:%M:%S')

    # Localize the time to UTC
    utc_time = pytz.utc.localize(utc_time)

    # Define client time zones
    time_zones = {
        'PST': 'America/Los_Angeles',
        'CST': 'America/Chicago',
        'EST': 'America/New_York'
    }

    # Convert to client time zone
    client_tz = pytz.timezone(time_zones[client_time_zone])
    client_time = utc_time.astimezone(client_tz)
    
    # Return the time as string without date part if it was initially in 'HH:MM:SS' format
    if isinstance(utc_time, time) or (isinstance(utc_time, str) and len(utc_time) == 8):
        return client_time.strftime('%H:%M:%S')
    else:
        return client_time.strftime('%H:%M:%S')

# Read the Excel file
df = pd.read_excel('file.xlsx')

# Create a new column for client time
client_times = []

# Loop through each row in the DataFrame
for index, row in df.iterrows():
    client_time = convert_utc_to_client_time(row['UTC_time'], row['client_time_zone'])
    client_times.append(client_time)

# Add the new column to the DataFrame
df['client_time'] = client_times

# Save the updated data to a new Excel file
df.to_excel('file_with_client_time.xlsx', index=False)

print("Conversion complete. The updated file is saved as 'file_with_client_time.xlsx'.")
