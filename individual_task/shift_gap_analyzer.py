import pandas as pd
import tkinter as tk
from tkinter import filedialog


root = tk.Tk()
root.withdraw()  

# Use the file dialog to select the Excel file
file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])

# Check if the user canceled the file selection
if not file_path:
    print("File selection canceled.")
else:
    # Read the Excel file into a DataFrame
    try:
        df = pd.read_excel(file_path)
        
        # Convert the "Time" and "Time Out" columns to datetime objects
        df['Time'] = pd.to_datetime(df['Time'], format='%m-%d-%Y %I:%M %p')
        df['Time Out'] = pd.to_datetime(df['Time Out'], format='%m-%d-%Y %I:%M %p')
        
        # Calculate the time difference between shifts
        df['Time Difference'] = (df['Time'] - df['Time Out'].shift()).dt.total_seconds() / 3600
        
        # Filter the DataFrame to find employees who meet the criteria
        filtered_df = df[(df['Time Difference'] > 1) & (df['Time Difference'] < 10)]
        
         # Drop duplicate rows based on the "Employee Name" column
        unique_employees = filtered_df[['Employee Name', 'Position ID']].drop_duplicates()
        
        # Display the employees who meet the criteria (only "Employee Name" and "Position ID" columns)
        print("Employee who meet the criteria")
        print(unique_employees)
    except FileNotFoundError:
        print(f"File not found: {file_path}")

# Close the Tkinter root window (optional)
root.destroy()
