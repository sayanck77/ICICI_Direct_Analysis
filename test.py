from datetime import datetime, timedelta
import pytz

# Define the UTC timezone
utc = pytz.timezone('UTC')

start_date_str = "2013-01-01T09:20:00.000Z"
end_date_str = "2025-09-29T15:29:00.000Z"

# Convert end_date string to datetime object and localize to UTC
end_date_utc = datetime.fromisoformat(end_date_str.replace('Z', '+00:00')).astimezone(utc)

# Initialize the start_date as a string
current_start_date_str = start_date_str

while True:
    # Convert the current start date string to a datetime object localized to UTC
    start_date_utc = datetime.fromisoformat(current_start_date_str.replace('Z', '+00:00')).astimezone(utc)

    # Check if the current start date is past the end date
    if start_date_utc > end_date_utc:
        break

    # Format the UTC datetime object back into a string in ISO format with 'Z' for the start date
    start_date_formatted = start_date_utc.isoformat().replace('+00:00', 'Z')
    print(f"Start date (UTC): {start_date_formatted}")

    # Calculate the end date for the current start date (already in UTC)
    # Set time to 15:29 in UTC and ensure microseconds are included
    end_date_current_utc = start_date_utc.replace(hour=15, minute=29, second=0, microsecond=0)
    end_date_current_formatted = end_date_current_utc.isoformat().replace('+00:00', 'Z')

    print(f"End date (UTC): {end_date_current_formatted}")
    print("\n")

    # Advance the start_date to the next day (in UTC) as a datetime object
    next_day_start_date_utc = start_date_utc + timedelta(days=1)

    # Convert the next day's start date datetime object back to a string for the next loop iteration
    current_start_date_str = next_day_start_date_utc.isoformat().replace('+00:00', 'Z')