import tkinter as tk
from tkinter import messagebox
import pandas as pd
import os

def get_script_directory():
    # Get the directory of the script
    script_dir = os.path.dirname(os.path.realpath(__file__))
    return script_dir

# Load Excel data into a DataFrame
try:
    df = pd.read_excel('RDIF Excel Spread.xlsx')  # Load data from RDIF Excel Spread.xlsx
except FileNotFoundError:
    messagebox.showerror("Error", "Excel file 'RDIF Excel Spread.xlsx' not found.")
    exit()

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
                df_file = pd.read_excel(file_path)

                # Specify the column name to search in
                column_name = 'RFIDNumber'

                # Check if the input_value is in the specified column
                if input_value in df_file[column_name].values:
                    found_files.append(file_name)

            except Exception as e:
                messagebox.showwarning("Error", f"Error occurred while processing {file_name}: {e}")

    return found_files

def on_search():
    dock_num = dock_entry.get().strip()
    dock_line_num = dock_line_entry.get().strip()
    package_line_num = package_line_entry.get().strip()

    if not (dock_num and dock_line_num and package_line_num):
        messagebox.showwarning("Incomplete Input", "Please provide values for Dock#, Dock Line#, and Package Line#.")
        return

    try:
        dock_line_num = int(dock_line_num)
        package_line_num = int(package_line_num)
    except ValueError:
        messagebox.showerror("Error", "Invalid input. DockLine# and PackageLine# must be numeric.")
        return

    # Perform lookup in the DataFrame to retrieve RFIDNumber
    try:
        # Filter DataFrame based on user inputs
        filtered_df = df[(df['Dock#'] == dock_num) &
                         (df['DockLine#'] == dock_line_num) &
                         (df['PackageLine#'] == package_line_num)]

        if not filtered_df.empty:
            # Retrieve the first matching RFIDNumber
            rfid_number = filtered_df.iloc[0]['RFIDNumber']

            # Perform the search across Excel files using the retrieved RFIDNumber
            found_files = search_value_in_files(rfid_number)

            if found_files:
                messagebox.showinfo("Search Results", f"The RFIDNumber '{rfid_number}' was found in:\n" + "\n".join(found_files))
            else:
                messagebox.showinfo("Search Results", f"The RFIDNumber '{rfid_number}' was not found in any Excel files.")

        else:
            messagebox.showinfo("Search Results", "No matching record found for the specified Dock#, Dock Line#, and Package Line#.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create main window
root = tk.Tk()
root.title("Excel Value Search")

# Create GUI widgets
tk.Label(root, text="Dock Number:").pack()
dock_entry = tk.Entry(root)
dock_entry.pack()

tk.Label(root, text="Dock Line Number:").pack()
dock_line_entry = tk.Entry(root)
dock_line_entry.pack()

tk.Label(root, text="Package Line Number:").pack()
package_line_entry = tk.Entry(root)
package_line_entry.pack()

search_button = tk.Button(root, text="Search", command=on_search)
search_button.pack(pady=10)

# Start GUI main loop
root.mainloop()
