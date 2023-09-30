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
        
        # Calculate the duration of each shift in hours
        df['Shift Duration'] = (df['Time Out'] - df['Time']).dt.total_seconds() / 3600
        
        # Filter rows where the shift duration is greater than 14 hours
        long_shifts = df[df['Shift Duration'] > 14]
        
        # Display employees who worked more than 14 hours in a single shift
        print(long_shifts[['Employee Name', 'Time', 'Time Out', 'Shift Duration']])
        
        output_file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        
        if not output_file_path:
            print('Output file selection canceled')
        else:
            with open(output_file_path,'w') as output_file:
                for index, row in long_shifts.iterrows():
                     output_file.write(f"{row['Employee Name']}\t{row['Time']}\t{row['Time Out']}\t{row['Shift Duration']:.2f} hours\n")
            print(f"Results saved to {output_file_path}")
    except FileNotFoundError:
        print(f"File not found: {file_path}")

# Close the Tkinter root window (optional)
root.destroy()
