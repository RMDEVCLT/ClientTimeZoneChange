import pandas as pd  # Importing the pandas library for handling Excel files
import pytz  # Importing pytz for time zone conversions
from datetime import datetime, time  # Importing datetime and time for handling time objects

def convert_utc_to_client_time(utc_time, client_time_zone):
    """
    Convert UTC time to the client's time zone.
    """
    # If utc_time is a time object, convert it to a datetime object with today's date
    if isinstance(utc_time, time):
        utc_time = datetime.combine(datetime.today(), utc_time)
    # If utc_time is a string, check its format and convert accordingly
    elif isinstance(utc_time, str):
        if len(utc_time) == 8:  # Format 'HH:MM:SS'
            utc_time = datetime.strptime(utc_time, '%H:%M:%S')
        else:
            utc_time = datetime.strptime(utc_time, '%H:%M:%S')

    # Localize the time to UTC
    utc_time = pytz.utc.localize(utc_time)

    # Dictionary to map time zone names to pytz time zones
    time_zones = {
        'PST': 'America/Los_Angeles',
        'CST': 'America/Chicago',
        'EST': 'America/New_York'
    }

    # Convert the UTC time to the client's time zone
    client_tz = pytz.timezone(time_zones[client_time_zone])
    client_time = utc_time.astimezone(client_tz)
    
    # Return the time as a string in 'HH:MM:SS' format
    return client_time.strftime('%H:%M:%S')

# Read the Excel file into a DataFrame
df = pd.read_excel('file.xlsx')

# Apply the time conversion function to each row and create a new 'client_time' column
df['client_time'] = df.apply(lambda row: convert_utc_to_client_time(row['UTC_time'], row['client_time_zone']), axis=1)

# Save the updated DataFrame to a new Excel file
df.to_excel('file_with_client_time.xlsx', index=False)

print("Conversion complete. The updated file is saved as 'file_with_client_time.xlsx'.")
