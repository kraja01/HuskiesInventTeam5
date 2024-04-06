import pandas as pd
import tkinter as tk
from tkinter import messagebox

# Load Excel data into a DataFrame
try:
    df = pd.read_excel('NAV.xlsx')  # Load all sheets from the Excel file
except FileNotFoundError:
    messagebox.showerror("Error", "Excel file 'NAV.xlsx' not found.")
    exit()

# Function to search RFID number based on inputs
def search_rfid():
    dock_num = dock_entry.get().strip()
    dock_line_num = dock_line_entry.get().strip()
    package_line_num = package_line_entry.get().strip()

    try:
        dock_line_num = int(dock_line_num)
        package_line_num = int(package_line_num)
    except ValueError:
        messagebox.showerror("Error", "Invalid input. DockLine# and PackageLine# must be numeric.")
        return

    # Perform lookup using DataFrame filtering
    try:
        # Filter DataFrame based on user inputs
        filtered_df = df[(df['Dock#'] == dock_num) & 
                         (df['DockLine#'] == dock_line_num) & 
                         (df['PackageLine#'] == package_line_num)]

        # Check if any rows match the criteria
        if not filtered_df.empty:
            # Retrieve the RFIDNumber from the first matching row
            rfid_number = filtered_df.iloc[0]['RFIDNumber']
            messagebox.showinfo("RFID Number", f"The corresponding RFID number is: {rfid_number}")
        else:
            messagebox.showwarning("No Match", "No matching record found.")

    except KeyError:
        messagebox.showerror("Error", "Column names not found in DataFrame.")

# Create the main application window
root = tk.Tk()
root.title("RFID Lookup")

# Create input labels and entry widgets
tk.Label(root, text="Dock Number:").pack()
dock_entry = tk.Entry(root)
dock_entry.pack()

tk.Label(root, text="Dock Line Number:").pack()
dock_line_entry = tk.Entry(root)
dock_line_entry.pack()

tk.Label(root, text="Package Line Number:").pack()
package_line_entry = tk.Entry(root)
package_line_entry.pack()

# Create search button
search_button = tk.Button(root, text="Search", command=search_rfid)
search_button.pack(pady=10)

# Run the main event loop
root.mainloop()
