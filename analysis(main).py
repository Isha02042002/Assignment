import pandas as pd
from datetime import datetime, timedelta

# Define a function to analyze an Excel file
def analyze_excel_file(file_path, threshold=7):
    try:
        # Load the Excel file into a DataFrame
        data = pd.read_excel(file_path)

        # Remove any leading/trailing spaces from column names
        data.columns = data.columns.str.strip()

        # Create empty sets to store information about employees
        consecutive_workers = set()
        short_break_workers = set()
        long_shift_workers = set()

        # Start analyzing the data
        print("Analyzing part 1...")

        # Iterate over each row in the DataFrame
        for idx, record in data.iterrows():
            name = record['Employee Name']
            id = record['Position ID']

            # Skip if the employee has already been processed
            if name in consecutive_workers:
                continue

            # Check if the employee has worked consecutively for a certain number of days
            if idx > 0 and name == data.at[idx - 1, 'Employee Name']:
                consecutive_days = 1
                for i in range(idx - 1, -1, -1):
                    if data.at[i, 'Employee Name'] == name:
                        consecutive_days += 1
                    else:
                        break
                if consecutive_days >= threshold:
                    print(f"Employee: {name}, Position: {id}")
                    consecutive_workers.add(name)

        print("Analyzing part 2...")

        # Create a dictionary to track breaks for each employee
        breaks = {}  

        # Iterate over each row again
        for idx, record in data.iterrows():
            name = record['Employee Name']
            id = record['Position ID']

            # Skip if the employee has already been processed
            if name not in short_break_workers:
                # Check if the employee has taken a short break
                if name in breaks:
                    last_out = breaks[name]
                    time_in = record['Time']

                    # Convert string times to datetime objects
                    if isinstance(time_in, str) and isinstance(last_out, str):
                        time_in = datetime.strptime(time_in, '%m/%d/%Y %I:%M %p')
                        last_out = datetime.strptime(last_out, '%m/%d/%Y %I:%M %p')

                        # Calculate the difference between the current time and the last out time
                        diff = (time_in - last_out).total_seconds() / 3600
                        if 1 < diff < 10:
                            print(f"Employee: {name}, Position: {id}")
                            short_break_workers.add(name)
                    else:
                        time_in = None

                # Update the last out time for the employee
                breaks[name] = record['Time Out']

        print("Analyzing part 3...")

        # Iterate over each row once more
        for idx, record in data.iterrows():
            name = record['Employee Name']
            id = record['Position ID']

            # Skip if the employee has already been processed
            if name not in long_shift_workers:
                # Check if the employee has worked a shift longer than 14 hours
                duration_str = record['Timecard Hours (as Time)']
                if pd.notna(duration_str):
                    try:
                        hours, minutes = map(int, duration_str.split(':'))
                        duration = timedelta(hours=hours, minutes=minutes)
                    except ValueError:
                        # Handle invalid duration format 
                        duration = None
                else:
                    # Handle missing values
                    duration = None

                if duration is not None and duration.total_seconds() / 3600 > 14:
                    print(f"Employee: {name}, Position: {id}")
                    long_shift_workers.add(name)

    except FileNotFoundError:
        print(f"File not found: {file_path}")
    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Run the function if the script is run directly
if __name__ == "__main__":
    file_path = 'Assignment.xlsx'
    analyze_excel_file(file_path, threshold=7)
