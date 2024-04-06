import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os
from remove_duplicates import remove_duplicate_rows_from_excel

def get_script_directory():
    # Get the directory of the script
    script_dir = os.path.dirname(os.path.realpath(__file__))
    return script_dir

# Remove duplicate rows from Excel files located in the script directory
remove_duplicate_rows_from_excel(get_script_directory())

def search_value_in_files(input_value):
    found_files = []

    # Get the directory of the script
    script_dir = get_script_directory()

    # Loop through each file in the script directory
    for file_name in os.listdir(script_dir):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(script_dir, file_name)
            try:
                # Read the Excel file
                df = pd.read_excel(file_path)

                # Specify the column name to search in
                column_name = 'RFIDNumber'

                # Check if the input_value is in the specified column
                if input_value in df[column_name].values:
                    found_files.append(file_name)

            except Exception as e:
                messagebox.showwarning("Error", f"Error occurred while processing {file_name}: {e}")

    return found_files

def on_search():
    input_value = value_entry.get().strip()

    if not input_value:
        messagebox.showwarning("Incomplete Input", "Please provide a value to search.")
        return

    # Perform the search in the script directory
    found_files = search_value_in_files(input_value)

    if found_files:
        messagebox.showinfo("Search Results", f"The value '{input_value}' was found in:\n" + "\n".join(found_files))
    else:
        messagebox.showinfo("Search Results", f"The value '{input_value}' was not found in any of the files.")

# Create main window
root = tk.Tk()
root.title("Excel Value Search")

# Create GUI widgets
value_label = tk.Label(root, text="Enter Value:")
value_label.pack(pady=10)

value_entry = tk.Entry(root, width=30)
value_entry.pack(pady=5)

search_button = tk.Button(root, text="Search", command=on_search)
search_button.pack(pady=10)

# Start GUI main loop
root.mainloop()
