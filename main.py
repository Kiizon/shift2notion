#!/usr/bin/env python3

import pandas as pd

name = "Kish"

days_of_week = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]


def build_events(file_path):
    """
    Find the name in the Excel file and return the row number
    """
    try:
        df = pd.read_excel(file_path, header=None)
        
        for idx, value in enumerate(df.iloc[:, 1]):
            if value == name:
                for i, day in enumerate(days_of_week):
                    am_col_idx = 2 + (i * 2)  # AM shift column index (even)
                    pm_col_idx = 3 + (i * 2)  # PM shift column index (odd)
        
                    am_shift = df.iloc[idx, am_col_idx]  # AM shift value
                    pm_shift = df.iloc[idx, pm_col_idx]  # PM shift value
                    
                    # Print out the day and shifts if available
                    if pd.notna(am_shift):
                        print(f"{day} AM: {am_shift}")
                    if pd.notna(pm_shift):
                        print(f"{day} PM: {pm_shift}")
    
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None



    

if __name__ == "__main__":
    file_path = "schedule.xls"
    name = build_events(file_path)