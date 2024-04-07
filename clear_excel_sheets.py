import os
import pandas as pd

def clear_excel_sheets(exclude_file_path):
    # Get the directory containing Excel files
    directory = os.getcwd()  # Use current working directory or specify your directory path

    # Loop through each file in the directory
    for filename in os.listdir(directory):
        if filename.endswith('.xlsx'):  # Process Excel files only
            file_path = os.path.join(directory, filename)

            if file_path == exclude_file_path:
                # Skip the file specified in exclude_file_path
                continue

            try:
                # Load the Excel file
                xls = pd.ExcelFile(file_path)

                # Create a new Excel writer for the same file path
                with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                    # Loop through each sheet and clear its contents
                    for sheet_name in xls.sheet_names:
                        df = pd.read_excel(file_path, sheet_name=sheet_name)  # Read existing data
                        df.iloc[0:0].to_excel(writer, sheet_name=sheet_name, index=False)  # Write empty DataFrame

                print(f"Cleared sheets in '{filename}'.")

            except Exception as e:
                print(f"Error occurred while processing '{filename}': {e}")

# Specify the path of the Excel file you want to exclude from clearing
exclude_file_path = "NAV.xlsx"

# Call the function to clear Excel sheets, excluding the specified file
clear_excel_sheets(exclude_file_path)
