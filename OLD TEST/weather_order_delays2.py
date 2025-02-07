import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import subprocess
import sys

# Ensure required dependencies are installed
def install_package(package):
    try:
        __import__(package)
    except ImportError:
        print(f"Installing {package}...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Install required packages
install_package("xlsxwriter")

# Function to select a file
def select_file(title):
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(title=title)
    return file_path

# Prompt user to select files
print("Select the EZ-Open CSV file")
ez_open_file_path = select_file("Select the EZ-Open CSV File")
print("Select the Weather Alert Excel file")
weather_alert_file_path = select_file("Select the Weather Alert Excel File")

# Load the data
ez_open_df = pd.read_csv(ez_open_file_path)
weather_alert_df = pd.read_excel(weather_alert_file_path, sheet_name=0)

# Function to match counties in the WHERE column
def match_county_in_where(county, where_text):
    """Check if the county exists within the WHERE text."""
    if pd.isna(where_text) or pd.isna(county):
        return False
    return county.lower() in where_text.lower()

# Extract relevant weather alert details for matching rows
def get_matching_alert_details(county, where_texts, columns, row_index):
    """Get the alert details for a matching county."""
    for i, where in enumerate(where_texts):
        if match_county_in_where(county, where):
            return {col: weather_alert_df.at[i, col] for col in columns}
    return {col: None for col in columns}

# Define relevant weather alert columns
alert_columns = ['Title', 'WHAT', 'WHERE', 'WHEN', 'IMPACTS']

# Apply matching logic and extract weather alert details
weather_alert_details = [
    get_matching_alert_details(row['County'], weather_alert_df['WHERE'], alert_columns, idx)
    for idx, row in ez_open_df.iterrows()
]

# Convert weather alert details to DataFrame
alert_details_df = pd.DataFrame(weather_alert_details)

# Merge weather alert details with EZ-Open data
merged_with_alerts_df = pd.concat([ez_open_df, alert_details_df], axis=1)

# Filter rows where at least one alert column is not NaN (indicating a match)
merged_with_alerts_df = merged_with_alerts_df.dropna(subset=['Title'])

# Select and rename relevant columns from EZ-Open data
columns_to_add_corrected = [
    'Job Id', 'Service', 'Street Addr', 'City', 'State',
    'County', 'Due', 'Rep Due', 'Client'
]

# Merge EZ-Open data with matched weather alerts
merged_with_all_columns_corrected = pd.merge(
    ez_open_df[columns_to_add_corrected],
    merged_with_alerts_df[['Job Id', 'Title', 'WHAT', 'WHERE', 'WHEN', 'IMPACTS']],
    on='Job Id',
    how='inner'
)

# Final column order
columns_order_corrected = [
    'Job Id', 'Service', 'Street Addr', 'City', 'State', 'County',
    'Due', 'Rep Due', 'Client', 'Title', 'WHAT', 'WHERE', 'WHEN', 'IMPACTS'
]

# Format data in the specified column order
formatted_df_corrected_final = merged_with_all_columns_corrected[columns_order_corrected]

# Group data by client and write to Excel
from datetime import datetime
counter = 1
output_file_path_corrected_final = f"Weather_Delayed_Orders_{datetime.now().strftime('%m-%d-%y')}.xlsx"
while os.path.exists(output_file_path_corrected_final):
    output_file_path_corrected_final = f"Weather_Delayed_Orders_{datetime.now().strftime('%m-%d-%y')}({counter}).xlsx"
    counter += 1
import os

# Ensure os is imported before using it

# Remove the file if it already exists to avoid permission errors
if os.path.exists(output_file_path_corrected_final):
    os.remove(output_file_path_corrected_final)

with pd.ExcelWriter(output_file_path_corrected_final, engine='xlsxwriter') as writer:
    # Add Alert Summary tab as the first tab
    alert_summary_sheet_name = f"Alert Summary {datetime.now().strftime('%m-%d-%Y')}"
    weather_alert_df.to_excel(writer, sheet_name=alert_summary_sheet_name, index=False)
    worksheet = writer.sheets[alert_summary_sheet_name]
    for col_num, col in enumerate(weather_alert_df.columns):
        worksheet.set_column(col_num, col_num, 15)  # Set column width to 15
    if not formatted_df_corrected_final.empty:
        grouped_corrected_final = formatted_df_corrected_final.groupby('Client')
        for client, data in grouped_corrected_final:
            sheet_name = client[:31]  # Ensure sheet names <=31 characters
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            for col_num, col in enumerate(data.columns):
                if col in ['State', 'Title']:
                    worksheet.set_column(col_num, col_num, None)  # Autofit 'State' and 'Title'
                else:
                    worksheet.set_column(col_num, col_num, 15)  # Set column width to 15
    else:
        # Write placeholder if no data matched
        pd.DataFrame({"Message": ["No matching records found"]}).to_excel(writer, sheet_name="No_Matches", index=False)

# Notify user of output file
print(f"Processing complete. Output saved to {output_file_path_corrected_final}")
