import pandas as pd

def remove_duplicates_excel(input_file, output_file):
    # Read the Excel file into a pandas DataFrame
    df = pd.read_excel(input_file)
    
    # Remove duplicates
    df.drop_duplicates(inplace=True)
    
    # Write the modified DataFrame to a new Excel file
    df.to_excel(output_file, index=False)
    
    print("Duplicates removed successfully!")

if __name__ == "__main__":
    input_file = "input.xlsx"  # Replace "input.xlsx" with the path to your input Excel file
    output_file = "output.xlsx"  # Replace "output.xlsx" with the desired name of the output Excel file
    
    remove_duplicates_excel(input_file, output_file)
