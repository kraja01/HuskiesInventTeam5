import tkinter as tk
from tkinter import messagebox
import pandas as pd

def search_excel(input_value):
    try:
        df = pd.read_excel('Shipping Book with RFID.xlsx')  # Update with your Excel file path
        column_name = 'RFIDNumber'  # Update with the column name to search in
        match = df[df[column_name] == input_value]

        if not match.empty:
            messagebox.showinfo("Match Found", f"Value '{input_value}' found in Excel!")
        else:
            messagebox.showinfo("No Match", f"No matching value found for '{input_value}'.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

def on_search():
    input_value = entry.get().strip()
    if input_value:
        search_excel(input_value)
    else:
        messagebox.showwarning("Empty Input", "Please enter a value to search.")

# Create main window
root = tk.Tk()
root.title("Excel Value Search")

# Create GUI widgets
label = tk.Label(root, text="Enter Value:")
label.pack(pady=10)

entry = tk.Entry(root, width=30)
entry.pack(pady=5)

search_button = tk.Button(root, text="Search", command=on_search)
search_button.pack(pady=10)

# Start GUI main loop
root.mainloop()
