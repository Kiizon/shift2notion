#!/usr/bin/env python3

import pandas as pd

name = "Kish"


close_time_map = {
    "monday": "24:00",
    "tuesday": "1:30",
    "wednesday": "1:00",
    "thursday": "1:00",
    "friday": "1:30",
    "saturday": "1:30",
    "sunday": "1:00" 
}

def get_shift_day(idx):
    """
    Accepts the index of the row and returns the date of the shift
    """
    days = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]
    day_idx = idx // 2
    return days[day_idx]

def get_shift_end_time(idx):
    """
    Accepts the index of the row and returns the end time of the shift
    """
    day = get_shift_day(idx)
    return close_time_map[day]


def build_events(file_path):
    """
    Find the name in the Excel file and return the row number
    """
    try:
        df = pd.read_excel(file_path, header=None)
        
        for idx, value in enumerate(df.iloc[:, 1]):
            if value == name:
                print(f"Row {idx}: {value}")
                for idx2, value2 in enumerate(df.iloc[idx, 2:16]):
                    print(f"Column {idx2}: {value2}")
                    if 
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

    for 
    event  = {
        "start_time": None,
        "end_time": None
    }
    return event

def 
    service.events().insert(calendarId='primary', body=event).execute()


    # format weekday columns
    # times will have -CL, each day closing will end different due to different closing times
    # Monday will have 12:00 am end
    # Tuesday will have 1:30 am end
    # Wednesday will have 12:30am end
    # Thursday will have 12:30am end
    # Friday will have 1:30am end
    # Saturday will have 1:30am end
    # Sunday will have 1:00 am end
    
    ## each day occupies 2 columns from column 2-16 (index 0-13)

    ## approach
    # 1. find the row number of the name
    # 2. iterate through the columns of the row
    # 3. map column index to day, time
    # 4. extract the time from the column
    # 
    
    

if __name__ == "__main__":
    file_path = "schedule.xls"
    name = find_name(file_path)