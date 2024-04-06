import pandas as pd

def search_excel(input_value):
    try:
        df = pd.read_excel('data.xlsx')  # Update with your Excel file path
        column_name = 'YourColumnName'  # Update with the column name to search in
        match = df[df[column_name] == input_value]

        if not match.empty:
            print(f"Value '{input_value}' found in Excel!")
        else:
            print(f"No matching value found for '{input_value}'.")

    except Exception as e:
        print(f"An error occurred: {str(e)}")

# Usage example
user_input = input("Enter value to search in Excel: ").strip()
if user_input:
    search_excel(user_input)
else:
    print("Please enter a value to search.")
