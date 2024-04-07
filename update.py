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

# Custom colors
bg_color = "#F0F0F0"  # Light grey background
label_color = "#333333"  # Dark grey for labels
entry_bg_color = "#FFFFFF"  # White background for entry widgets
button_bg_color = "#4CAF50"  # Green background for the search button
button_fg_color = "#FFFFFF"  # White text color for the search button

# Configure root window background color
root.configure(bg=bg_color)

# Function to create a label and entry widget with custom colors
def create_entry_with_label(text):
    frame = tk.Frame(root, bg=bg_color)
    frame.pack(pady=5)
    label = tk.Label(frame, text=text, bg=bg_color, fg=label_color)
    label.pack(side=tk.LEFT)
    entry = tk.Entry(frame, bg=entry_bg_color)
    entry.pack(side=tk.LEFT)
    return entry

# Create input labels and entry widgets with custom colors
dock_entry = create_entry_with_label("Dock Number:")
dock_line_entry = create_entry_with_label("Dock Line Number:")
package_line_entry = create_entry_with_label("Package Line Number:")

# Create search button with custom colors
search_button = tk.Button(root, text="Search", command=search_rfid, bg=button_bg_color, fg=button_fg_color)
search_button.pack(pady=10)

# Run the main event loop
root.mainloop()