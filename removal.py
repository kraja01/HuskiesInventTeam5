import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import os

def remove_duplicates_excel(input_file, output_file):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)
    
    # Remove duplicates
    df.drop_duplicates(inplace=True)
    
    # Write the modified DataFrame to a new Excel file
    df.to_excel(output_file, index=False)
    
    messagebox.showinfo("Success", "Duplicates removed successfully!")

def browse_input_file():
    input_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if input_file_path:
        input_entry.delete(0, tk.END)
        input_entry.insert(0, input_file_path)

def browse_output_file():
    output_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if output_file_path:
        output_entry.delete(0, tk.END)
        output_entry.insert(0, output_file_path)

def remove_duplicates():
    input_file = input_entry.get()
    output_file = output_entry.get()
    if not input_file or not output_file:
        messagebox.showerror("Error", "Please select input and output files.")
        return
    
    # Get the current directory
    current_directory = os.getcwd()
    
    # Join file names with the current directory
    input_file = os.path.join(current_directory, input_file)
    output_file = os.path.join(current_directory, output_file)
    
    try:
        remove_duplicates_excel(input_file, output_file)
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")

# GUI setup
root = tk.Tk()
root.title("Excel Duplicates Remover")

input_label = tk.Label(root, text="Input Excel File:")
input_label.grid(row=0, column=0, padx=5, pady=5)

input_entry = tk.Entry(root, width=50)
input_entry.grid(row=0, column=1, padx=5, pady=5)

browse_input_button = tk.Button(root, text="Browse", command=browse_input_file)
browse_input_button.grid(row=0, column=2, padx=5, pady=5)

output_label = tk.Label(root, text="Output Excel File:")
output_label.grid(row=1, column=0, padx=5, pady=5)

output_entry = tk.Entry(root, width=50)
output_entry.grid(row=1, column=1, padx=5, pady=5)

browse_output_button = tk.Button(root, text="Browse", command=browse_output_file)
browse_output_button.grid(row=1, column=2, padx=5, pady=5)

remove_button = tk.Button(root, text="Remove Duplicates", command=remove_duplicates)
remove_button.grid(row=2, column=1, padx=5, pady=5)

root.mainloop()
