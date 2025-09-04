#!/usr/bin/env python3

import pandas as pd

name = "Kish"


def find_name(file_path):
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
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

if __name__ == "__main__":
    file_path = "schedule.xls"
    name = find_name(file_path)