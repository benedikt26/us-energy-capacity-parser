import os
from datetime import datetime
import pandas as pd

# Specify the working directory containing the folders
folder_directory = R"C:/Users/lechl/OneDrive - TUM/Hiwi/Jeana/Energy Capacities/Data"

# Specify the result file path
result_file = R"C:/Users/lechl/OneDrive - TUM/Hiwi/Jeana/Energy Capacities/Output/energy-capacities.xlsx"

# Create an empty DataFrame for the result
result_df = pd.DataFrame(columns=["Date", "State", "WindCapacity"])

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
            
            date_value = str(excel_data.iloc[2, 1])
            state_value_1 = excel_data.iloc[60, 0]
            wind_capacity_1 = excel_data.iloc[60, 1]
            state_value_2 = excel_data.iloc[52, 0]
            wind_capacity_2 = excel_data.iloc[52, 1]
            
            # Convert the date value to the desired format
            date_value = datetime.strptime(date_value, "%B %Y").strftime("%Y-%m")
            
            # Create a DataFrame from the extracted data
            extracted_data_1 = pd.DataFrame({"Date": [date_value],
                                        "State": [state_value_1],
                                        "WindCapacity": [wind_capacity_1]})
            extracted_data_2 = pd.DataFrame({"Date": [date_value],
                                        "State": [state_value_2],
                                        "WindCapacity": [wind_capacity_2]})

            # Concatenate the extracted data with the result DataFrame
            result_df = pd.concat([result_df, extracted_data_1], ignore_index=True)
            result_df = pd.concat([result_df, extracted_data_2], ignore_index=True)
            
            # Increment the successful runs counter
            successful_runs += 1
        
        except Exception as e:
            # Log the error and increment the unsuccessful runs counter
            unsuccessful_runs += 1
            error_files.append(file_path)
            print(f"Error processing file: {file_path}")
            print(f"Error message: {str(e)}")
    
# Save the result DataFrame to the specified file
result_df.to_excel(result_file, index=False)

# Display the statistics
print("Processing complete.")
print(f"Successful runs: {successful_runs}")
print(f"Unsuccessful runs: {unsuccessful_runs}")
print("Error files:", error_files)
