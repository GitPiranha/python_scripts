import pandas as pd
import os

# Set the directory where your Excel files are located
excel_files_directory = r'C:\01_Programming\python\combine_multiple_excel_files'

# Get a list of all Excel files in the directory
excel_files = [f for f in os.listdir(excel_files_directory) if f.endswith('.xlsx')]

# Create an empty DataFrame to store the combined data
combined_data = pd.DataFrame()

# Iterate through each Excel file and concatenate its data to the combined DataFrame
for file in excel_files:
    file_path = os.path.join(excel_files_directory, file)
    df = pd.read_excel(file_path)
    combined_data = pd.concat([combined_data, df], axis=1)

# Save the combined DataFrame to a new Excel file
combined_data.to_excel(r'C:\01_Programming\python\combine_multiple_excel_files\output\file.xlsx', index=False)

print("Combining Excel files completed.")
