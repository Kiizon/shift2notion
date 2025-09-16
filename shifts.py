import pandas as pd
from datetime import datetime, timedelta

name = "Kish"
days_of_week = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]

def get_shifts(file_path) -> list:
    """
    Find the name in the Excel file and return the row number
    """
    try:
        shifts = []
        df = pd.read_excel(file_path, header=None)
        
        for idx, value in enumerate(df.iloc[:, 1]):
            if value == name:
                for i, day in enumerate(days_of_week):
                    am_col_idx = 2 + (i * 2)  # AM shift column index (even)
                    pm_col_idx = 3 + (i * 2)  # PM shift column index (odd)
        
                    am_shift = df.iloc[idx, am_col_idx]  # AM shift value
                    pm_shift = df.iloc[idx, pm_col_idx]  # PM shift value
                    
                    if pd.notna(am_shift):
                        shifts.append(f"{day} {am_shift}")
                    if pd.notna(pm_shift):
                        shifts.append(f"{day} {pm_shift}")
        print(shifts)
        formatted_shifts = parse_shifts(shifts)
        return formatted_shifts
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

# def parse_shifts(shifts):
#     """
#     Accepts a string of shifts and returns the start and end times

#     tuesday PM: 6-CL -> (18:00, 23:59)
#     wednesday PM: 5-CL -> (17:00, 23:59)
#     thursday PM: 6-CL -> (18:00, 23:59)

#     """
#     if not shifts or pd.isna(shifts):
#         return None, None
#     shifts = str(shifts).split() 
#     print(shifts)

def parse_shifts(shifts):
    """
    Accepts a string of shifts and returns the start and end times

    ['tuesday 6-CL', 'wednesday 5-CL', 'thursday 6-CL']

    """
    formatted_shifts = []
    if not shifts:
        print("No shifts found")
        return None, None
    

    for shift in shifts:
        parts = shift.split()
        if len(parts) != 2:
            return None, None
    
        day = parts[0].strip()  # "tuesday"
        shift_time = parts[1].strip()  # "6-CL"
    
        if "-CL" in shift_time:
            start_hour = int(shift_time.split("-CL")[0])
            start_time = f"{start_hour}:00"
            end_time = "23:59"
            formatted_shifts.append((day,start_time,end_time))

    return formatted_shifts

def build_events(formatted_shifts):
    """
    Builds events from the formatted shifts
    """
    for shift in formatted_shifts:
        day, start, end = shift

        base_date = datetime.now()

        shift_date = get_actual_date(base_date,day) # get the actual date of the shift

        start_time = datetime.strptime(start, "%H:%M").time()
        end_time = datetime.strptime(end, "%H:%M").time()

        start_dt = datetime.combine(shift_date.date(), start_time).isoformat()
        end_dt = datetime.combine(shift_date.date(), end_time).isoformat()
        event = {
            "summary": "Work Shift",
            "description": "Work Shift",
            "colorId": 1,
            "start": {
                "dateTime": start_dt
            },
            "end": {
                "dateTime": end_dt
            }
        }
        print(event)

def get_actual_date(base_date, day):
    """
    Get the actual date of the shift
    """
    days_of_week = {
        "monday": 0,
        "tuesday": 1,
        "wednesday": 2,
        "thursday": 3,
        "friday": 4,
        "saturday": 5,
        "sunday": 6
    }
    target = days_of_week[day.lower()]
    base_weekday = base_date.weekday()
    # Find offset to next occurrence
    days_ahead = (target - base_weekday) % 7
    return base_date + timedelta(days=days_ahead)




if __name__ == "__main__":
    file_path = "pm test.xls"
    build_events(get_shifts(file_path))