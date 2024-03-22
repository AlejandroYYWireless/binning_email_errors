import os
import re
import pandas as pd

# Directory where your subdirectories containing .eml files are stored
parent_directory = "emails"

# Adjusted regular expressions for broader matching
scanned_bin_re = re.compile(r"Scanned Bin: ([^\s<]+)")
scanned_imei_re = re.compile(r"Scanned IMEI: ([\w-]+)")

# List to store extracted data
data = []

# Iterate over each subdirectory in the parent directory
for folder in os.listdir(parent_directory):
    folder_path = os.path.join(parent_directory, folder)
    if os.path.isdir(folder_path):  # Check if it is a directory
        print(f"Processing folder: {folder}")  # Debugging line
        for filename in os.listdir(folder_path):
            if filename.endswith(".eml"):
                file_path = os.path.join(folder_path, filename)
                with open(file_path, 'r', encoding='utf-8') as file:
                    content = file.read()
                    # First pass to find if the lines exist
                    scanned_bin_match = scanned_bin_re.search(content)
                    scanned_imei_match = scanned_imei_re.search(content)

                    if scanned_bin_match and scanned_imei_match:
                        # Print the found values for debugging
                        # print(f"File: {filename}")
                        # print(f"Scanned Bin: {scanned_bin_match.group(1)}")
                        # print(f"Scanned IMEI: {scanned_imei_match.group(1)}\n")

                        # Append the found data to our list
                        data.append({
                            "Scanned Bin": scanned_bin_match.group(1),
                            "Scanned IMEI": scanned_imei_match.group(1)
                        })

# Convert the list to a pandas DataFrame
df = pd.DataFrame(data)

# Specify your Excel file name
excel_file_name = "binning_errors.xlsx"

# Write the DataFrame to an Excel file
df.to_excel(excel_file_name, index=False, engine='openpyxl')

print(f"Data written to {excel_file_name}")
