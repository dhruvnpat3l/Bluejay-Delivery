import pandas as pd
import tkinter as tk
from tkinter import filedialog

# Create a Tkinter root window (it will not be displayed)
root = tk.Tk()
root.withdraw()  # Hide the root window

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
        df['Time'] = pd.to_datetime(df['Time'], format='%Y-%m-%d %H:%M:%S')
        
        # Sort the DataFrame by the "Employee Name" and "Time" columns
        df.sort_values(by=['Employee Name', 'Time'], inplace=True)
        
        #  Calculate consecutive workdays for each employee
        consecutive_days = df.groupby('Employee Name')['Time'].diff().dt.days == 1
        
        
        # Calculate the duration of each shift in hours
        df['Shift Duration'] = (df['Time Out'] - df['Time']).dt.total_seconds() / 3600
        
        # Identify employees who have worked for 7 consecutive days
        consecutive_counts = consecutive_days.groupby(df['Employee Name']).cumsum()
        employees_with_7_consecutive_days = df[consecutive_counts == 6]['Employee Name'].unique()

        
        # Display employees who worked more than 14 hours in a single shift
        print("Employees who have worked for 7 consecutive days:")
        for employee in employees_with_7_consecutive_days:
            print(employee)
        
    except FileNotFoundError:
        print(f"File not found: {file_path}")

# Close the Tkinter root window (optional)
root.destroy()
