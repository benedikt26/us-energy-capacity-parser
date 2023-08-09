import os
from datetime import datetime
import pandas as pd
import openpyxl

# Specify the working directory containing the folders
folder_directory = R"/Users/lechl/Library/CloudStorage/OneDrive-TUM/Hiwi/Jeana/Energy Capacities/Data"


# Specify the result file path
result_file_xlsx = R"/Users/lechl/Library/CloudStorage/OneDrive-TUM/Hiwi/Jeana/Energy Capacities/Output/energy-capacities.xlsx"
result_file_pkl = R"/Users/lechl/Library/CloudStorage/OneDrive-TUM/Hiwi/Jeana/Energy Capacities/Output/energy-capacities.pkl"

# Create an empty DataFrame for the result
result_df = pd.DataFrame()

# Initialize counters for successful and unsuccessful runs
successful_runs = 0
unsuccessful_runs = 0
error_files = []

# Iterate over folders in the directory
for folder in os.listdir(folder_directory):
    folder_path = os.path.join(folder_directory, folder)
    
    # Check if the folder is valid
    if os.path.isdir(folder_path):
        file_path = os.path.join(folder_path, "Table_6_02_B.xlsx")
        
        try:
            # Extract the month and year from the folder name
            month = folder[:-4]
            year = folder[-4:]
            datecheck = pd.to_datetime(f"{month} {year}", format="%B %Y")

            # Check whether it is a file of old format or of new format according to the folder name
            if datecheck <= pd.to_datetime("October 2015"):
                excel_data = pd.read_excel(file_path, header=None)
            else:
                excel_data = pd.read_excel(file_path)
            
            # Convert the date value to the desired format
            date_value = str(excel_data.iloc[2, 1])
            date_value = datetime.strptime(date_value, "%B %Y").strftime("%Y-%m")
            
            # Extract wind energy capacity values for all states (rows 5-64, excluding specific rows)
            state_wind_capacities = {}
            for row_idx in range(4, 63):
                if row_idx not in [10, 14, 20, 28, 38, 43, 48, 57, 61]:
                    state = excel_data.iloc[row_idx, 0]
                    capacity = excel_data.iloc[row_idx, 1]
                    state_wind_capacities[state] = capacity
            
            # Create a DataFrame from the extracted data for the current month/year
            extracted_data = pd.DataFrame(state_wind_capacities, index=[date_value])
            extracted_data.insert(0, "Date", date_value)  # Insert the "Date" column as the first column
            
            # Concatenate the extracted data with the result DataFrame
            result_df = pd.concat([result_df, extracted_data], ignore_index=False)
            
            # Increment the successful runs counter
            successful_runs += 1
        
        except Exception as e:
            # Log the error and increment the unsuccessful runs counter
            unsuccessful_runs += 1
            error_files.append(file_path)
            print(f"Error processing file: {file_path}")
            print(f"Error message: {str(e)}")
    
# Sort the result DataFrame based on the "Date" column in ascending order
result_df = result_df.sort_values(by="Date")

# Save the result DataFrame to the specified file
result_df.to_excel(result_file_xlsx, index=False)
result_df.to_pickle(result_file_pkl)

# Display the statistics
print("Processing complete.")
print(f"Successful runs: {successful_runs}")
print(f"Unsuccessful runs: {unsuccessful_runs}")
print("Error files:", error_files)
