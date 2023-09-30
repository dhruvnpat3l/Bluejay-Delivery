# Employee Work Analysis Tool

This tool is designed to analyze employee work records provided in an Excel file. It performs the following tasks:

1. **Analyze Employee Work Records**

    - Takes an Excel file as input.
    - Iterates over the records in the file and performs the following analyses:

        a) Identifies employees who have worked for 7 consecutive days.
        b) Identifies employees who have less than 10 hours of time between shifts but greater than 1 hour.
        c) Identifies employees who have worked for more than 14 hours in a single shift.

## Project Structure

The project is organized into two folders:

1. **python_files**

    - Contains the Python script for analyzing employee work records.

2. **jupyter_notebooks**

    - Contains a Jupyter Notebook file for running the **Analysis** interactively.
    

## Usage

### Python Script

1. Open a terminal/command prompt.
2. Navigate to the directory containing the Python script.
3. Run the script using the command:
4. Enter the file in filedialog box
5. The script will analyze the data and display hte results

## Requirements

- Python 3.x
- pandas library
- tkinter library