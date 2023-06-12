import requests
import time
import os

# Specify the working directory to save the downloaded zip files
working_directory = R"C:/Users/lechl/OneDrive - TUM/Hiwi/Jeana/Energy Capacities/Raw Data"

# Initialize counters for successful and unsuccessful runs
successful_runs = 0
unsuccessful_runs = 0
error_urls = []

# Define the start and end years for the range
start_year = 2000
end_year = 2023

# Define the months and corresponding URL fragments
months = [
    "january", "february", "march", "april", "may", "june", 
    "july", "august", "september", "october", "november", "december"
]

# Iterate over years and months
for year in range(start_year, end_year + 1):
    for month in months:
        # Construct the URL for the specific month and year
        url = f"https://www.eia.gov/electricity/monthly/archive/{month}{year}.zip"
        
        try:
            # Send a request to download the zip file
            response = requests.get(url)
            
            # Check if the request was successful
            if response.status_code == 200:
                # Extract the file name from the URL
                file_name = url.split("/")[-1]
                
                # Save the zip file in the working directory
                file_path = os.path.join(working_directory, file_name)
                with open(file_path, "wb") as file:
                    file.write(response.content)
                
                # Increment the successful runs counter
                successful_runs += 1
                print(file_name + " successfully downloaded")
                
                # Pause for a few seconds to avoid overloading the website
                time.sleep(3)
            else:
                # Log the unsuccessful run and increment the counter
                unsuccessful_runs += 1
                error_urls.append(url)
                print(f"Error downloading file: {url}. Status code: {response.status_code}")
        
        except Exception as e:
            # Log the error and increment the counter
            unsuccessful_runs += 1
            error_urls.append(url)
            print(f"Error downloading file: {url}")
            print(f"Error message: {str(e)}")

# Display the statistics
print("Downloading complete.")
print(f"Successful runs: {successful_runs}")
print(f"Unsuccessful runs: {unsuccessful_runs}")
print("Error URLs:", error_urls)
