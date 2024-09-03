import os
import pandas as pd
from pyxlsb import open_workbook

def count_pending_in_status_column(file_name, sheet_name):
    with open_workbook(file_name) as wb:
        with wb.get_sheet(sheet_name) as sheet:
            # Read the data into a pandas DataFrame
            data = []
            for row in sheet.rows():
                data.append([item.v for item in row])

            # Convert to DataFrame
            df = pd.DataFrame(data[1:], columns=data[0])

            # Check if 'Status' column exists
            if 'Status' not in df.columns:
                raise KeyError(f"'Status' column not found in sheet '{sheet_name}' of file '{file_name}'.")

            # Count the occurrences of 'Pending' in the 'Status' column
            pending_count = df['Status'].str.contains('Pending', case=False).sum()

            # Count the total number of rows
            total_rows = len(df)
    
    return pending_count, total_rows

def process_files():
    # List all files in the current directory
    files = os.listdir()
    
    # Filter out the .xlsb files
    xlsb_files = [file for file in files if file.endswith('.xlsb')]
    
    if len(xlsb_files) < 2:
        print("There should be at least two .xlsb files in the directory.")
        return None
    
    # Display available files to the user
    print("Available .xlsb files:")
    for i, file_name in enumerate(xlsb_files, 1):
        print(f"{i}. {file_name}")
    
    # Prompt user to select the first and second files
    first_file_choice = int(input("Enter the number corresponding to the first file: "))
    second_file_choice = int(input("Enter the number corresponding to the second file: "))
    
    first_file = xlsb_files[first_file_choice - 1]
    second_file = xlsb_files[second_file_choice - 1]
    
    def get_pending_count_and_rows(file_name):
        with open_workbook(file_name) as wb:
            sheet_names = wb.sheets
            
            # Display available sheets to the user
            print(f"Available sheets in '{file_name}':")
            for i, sheet_name in enumerate(sheet_names, 1):
                print(f"{i}. {sheet_name}")
            
            # Prompt user to select a sheet
            sheet_choice = int(input("Enter the number corresponding to the sheet you want to analyze: "))
            chosen_sheet_name = sheet_names[sheet_choice - 1]
            
            # Count 'Pending' and total rows
            pending_count, total_rows = count_pending_in_status_column(file_name, chosen_sheet_name)
        
        return pending_count, total_rows
    
    # Get pending counts and row counts for both files
    first_file_pending_count, first_file_rows = get_pending_count_and_rows(first_file)
    second_file_pending_count, second_file_rows = get_pending_count_and_rows(second_file)
    
    # Calculate the difference in row count between the two sheets
    row_difference = abs(second_file_rows - first_file_rows)
    
    # Calculate the final result
    final_result = first_file_pending_count + second_file_pending_count + row_difference
    
    print(f"Pending in first file: {first_file_pending_count}")
    print(f"Pending in second file: {second_file_pending_count}")
    print(f"Difference in row count: {row_difference}")
    print(f"Final result (Pending in both files + row difference): {final_result}")

# Run the function to perform the task
process_files()
