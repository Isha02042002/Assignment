import pandas as pd
from datetime import datetime, timedelta
import streamlit as st
# Define a function to analyze an Excel file
def analyze_excel_file(file_path, analysis_part=1, consecutive_days_threshold=7):
    try:
         # Read the Excel file and strip whitespace from column headers
        df = pd.read_excel(file_path)
        df.columns = df.columns.str.strip()
         # Initialize sets to store employees who have been printed in each analysis part
        consecutive_printed = set()
        short_break_printed = set()
        long_shift_printed = set()

         # Set up Streamlit configuration and layout
        st.set_page_config(
            page_title="Excel Analysis App",
            page_icon=":chart_with_upwards_trend:",
            layout="wide",
            initial_sidebar_state="expanded",
        )

          # Add custom CSS styles
        st.markdown("""
        <style>
    [data-testid=stSidebar] {
        background-color: #778899;
        font-size: 7rem !important;
    }

    /* Change the table header background color */
    thead > tr > th {
        background-color: #b0c4de !important;
        color: #000000 !important;
        font-family: Arial, sans-serif; /* Change the font style here */
    }
   /* Change the color of odd rows */
tbody > tr:nth-child(odd) {
    background-color: #f2f2f2; /* Change the color as needed */
    color: #28231D
}

/* Change the color of even rows */
tbody > tr:nth-child(even) {
    background-color: #dddddd; /* Change the color as needed */
                      color: #28231D
}

  
    /* Change the table border color and make it bolder */
    table, th, td {
        border: 3px solid #1c2841 !important; /* Change the border color and width here */
    }


        </style>
        """, unsafe_allow_html=True)

        # Display title and description of the application
        st.title("📊 Excel Analysis Application")
        st.markdown(
            """
            Analyze different aspects of the Excel file, and visualize the results interactively!
            """
        )
         # Add sidebar with analysis options
        st.sidebar.title("Analysis Options")
        analysis_part = st.sidebar.radio("Select Analysis Part", [1, 2, 3], key="analysis_part")

        # Perform different analyses depending on the selected part
        if analysis_part == 1:
            # Analyze consecutive days worked
            st.header("Consecutive Days Worked Analysis")
            result_data = []  

            for index, row in df.iterrows():
                employee_name = row['Employee Name']
                position_id = row['Position ID']

                if employee_name in consecutive_printed:
                    continue

                if index > 0 and employee_name == df.at[index - 1, 'Employee Name']:
                    consecutive_days = 1
                    for i in range(index - 1, -1, -1):
                        if df.at[i, 'Employee Name'] == employee_name:
                            consecutive_days += 1
                        else:
                            break

                    if consecutive_days >= consecutive_days_threshold:
                        result_data.append({'Employee': employee_name, 'Position': position_id, 'Consecutive Days Worked': consecutive_days})
                        consecutive_printed.add(employee_name)

            st.subheader("Results:")
            result_df = pd.DataFrame(result_data)
            st.table(result_df)

        elif analysis_part == 2:
             # Analyze short breaks
            st.header("Short Breaks Analysis")
            result_data = []  
            employee_breaks = {}  

            for index, row in df.iterrows():
                employee_name = row['Employee Name']
                position_id = row['Position ID']

                if employee_name in short_break_printed:
                    continue

                if employee_name in employee_breaks:
                    last_time_out = employee_breaks[employee_name]
                    time_in = row['Time']

                    if isinstance(time_in, str) and isinstance(last_time_out, str):
                        time_in = datetime.strptime(time_in, '%m/%d/%Y %I:%M %p')
                        last_time_out = datetime.strptime(last_time_out, '%m/%d/%Y %I:%M %p')

                        time_diff = (time_in - last_time_out).total_seconds() / 3600
                        if 1 < time_diff < 10:
                            result_data.append({'Employee': employee_name, 'Position': position_id, 'Analysis Result': 'Short Break'})
                            short_break_printed.add(employee_name)
                    else:
                        time_in = None

                employee_breaks[employee_name] = row['Time Out']

            st.subheader("Results:")
            if not result_data:
                st.write("No employees found with short breaks.")
            else:
                result_df = pd.DataFrame(result_data)
                st.table(result_df)

        elif analysis_part == 3:
         # Analyze long shifts

            st.header("Long Shifts Analysis")
            result_data = []  

            for index, row in df.iterrows():
                employee_name = row['Employee Name']
                position_id = row['Position ID']

                if employee_name in long_shift_printed:
                    continue

                duration_str = row['Timecard Hours (as Time)']
                if pd.notna(duration_str):
                    try:
                        hours, minutes = map(int, duration_str.split(':'))
                        duration = timedelta(hours=hours, minutes=minutes)
                    except ValueError:
                        duration = None
                else:
                    duration = None

                if duration is not None and duration.total_seconds() / 3600 > 14:
                    result_data.append({'Employee': employee_name, 'Position': position_id, 'Analysis Result': 'Long Shift'})
                    long_shift_printed.add(employee_name)

            st.subheader("Results:")
            if not result_data:
                st.write("No employees found with long shifts.")
            else:
                result_df = pd.DataFrame(result_data)
                st.table(result_df)

      # Handle file not found error
    except FileNotFoundError:
        st.error(f"File not found: {file_path}")
        # Handle other exceptions
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
# Run the main function
if __name__ == "__main__":
    file_path = 'Assignment.xlsx'
    analyze_excel_file(file_path)
