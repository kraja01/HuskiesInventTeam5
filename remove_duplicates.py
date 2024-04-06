import pandas as pd
import os

def get_script_directory():
    # Get the directory of the script
    script_dir = os.path.dirname(os.path.realpath(__file__))
    return script_dir

def remove_duplicate_rows_from_excel(directory_path):
    # List to store unique rows (as tuples of values)
    unique_rows = []

    # Loop through each file in the directory
    for file_name in os.listdir(directory_path):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(directory_path, file_name)
            try:
                # Read the Excel file
                df = pd.read_excel(file_path)

                # Convert each row to a tuple of values (to compare entire rows)
                df_rows = [tuple(row) for row in df.values]

                # Filter out duplicate rows
                unique_rows = list(set(df_rows))  # Use set to automatically remove duplicates

                # Create a DataFrame with unique rows
                df_unique = pd.DataFrame(unique_rows, columns=df.columns)

                # Save the DataFrame with unique rows back to the same file
                df_unique.to_excel(file_path, index=False)

                print(f"Duplicates removed from {file_name}")

            except Exception as e:
                print(f"Error occurred while processing {file_name}: {e}")

if __name__ == "__main__":
    # Get the directory of the script
    script_dir = get_script_directory()

    # Remove duplicate rows from Excel files located in the script directory
    remove_duplicate_rows_from_excel(script_dir)

    print("Duplicate rows have been removed from Excel files in the script directory.")
