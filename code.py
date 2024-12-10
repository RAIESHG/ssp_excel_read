import pandas as pd
import openpyxl
from PIL import Image
import io
import os
import xlrd

def get_excel_sheets(file_path):
    """Read all sheets from Excel file and return sheet names"""
    try:
        excel_file = pd.ExcelFile(file_path)
        return excel_file.sheet_names
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def extract_images(file_path, sheet_name):
    """Extract images from the specified sheet"""
    # Check file extension
    file_extension = os.path.splitext(file_path)[1].lower()
    
    if file_extension == '.xls':
        # For .xls files, use xlrd
        workbook = xlrd.open_workbook(file_path, formatting_info=True)
        sheet = workbook.sheet_by_name(sheet_name)
        images = []
        # Note: xlrd doesn't support image extraction from .xls files
        print("Warning: Image extraction is not supported for .xls files. Please convert to .xlsx format.")
        return images
    else:
        # For .xlsx files, use openpyxl
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        images = []
        
        for image in sheet._images:
            image_data = Image.open(io.BytesIO(image._data()))
            images.append(image_data)
        
        return images

def identify_tables(df):
    """Identify tables in the DataFrame and return their positions"""
    tables = []
    current_table_start = None
    empty_row_count = 0
    min_rows = 3  # Minimum consecutive rows required
    min_columns = 5  # Minimum columns with data required
    
    def is_valid_data_row(row):
        """Check if row has at least min_columns with data"""
        # Count values that are not NaN, empty string, or None
        valid_values = sum(1 for val in row if pd.notna(val) 
                         and str(val).strip() != ''
                         and val is not None)
        return valid_values >= min_columns
    
    def check_consecutive_rows(start_idx, min_rows):
        """Check if there are minimum required consecutive rows with valid data"""
        if start_idx + min_rows > len(df):
            return False
        for i in range(start_idx, start_idx + min_rows):
            if not is_valid_data_row(df.iloc[i]):
                return False
        return True
    
    # Iterate through rows to find table boundaries
    for idx in range(len(df)):
        row = df.iloc[idx]
        is_empty = not is_valid_data_row(row)
        
        if not is_empty and current_table_start is None:
            # Check if we have enough consecutive valid rows to start a table
            if check_consecutive_rows(idx, min_rows):
                current_table_start = idx
                empty_row_count = 0
        elif is_empty and current_table_start is not None:
            empty_row_count += 1
            # If we find 2 or more consecutive empty rows, consider it as table end
            if empty_row_count >= 2:
                tables.append({
                    'start': current_table_start,
                    'end': idx - empty_row_count
                })
                current_table_start = None
                empty_row_count = 0
    
    # Handle case where last table extends to end of sheet
    if current_table_start is not None:
        tables.append({
            'start': current_table_start,
            'end': len(df) - 1
        })
    
    return tables

def main():
    file_path = "ref.xls"
    
    # Get all sheet names
    sheet_names = get_excel_sheets(file_path)
    if not sheet_names:
        print("No sheets found or error reading file")
        return
    
    # Display available sheets
    print("\nAvailable sheets:")
    for idx, sheet in enumerate(sheet_names, 1):
        print(f"{idx}. {sheet}")
    
    # Ask user to select sheet
    while True:
        try:
            sheet_idx = int(input("\nSelect sheet number: ")) - 1
            if 0 <= sheet_idx < len(sheet_names):
                selected_sheet = sheet_names[sheet_idx]
                break
            else:
                print("Invalid selection. Please try again.")
        except ValueError:
            print("Please enter a valid number.")
    
    # Read the selected sheet
    df = pd.read_excel(file_path, sheet_name=selected_sheet)
    
    # Identify tables
    tables = identify_tables(df)
    
    # Display table information
    print(f"\nFound {len(tables)} tables in sheet '{selected_sheet}':")
    for idx, table in enumerate(tables, 1):
        start_row = table['start'] + 1  # Adding 1 for Excel row numbers
        end_row = table['end'] + 1
        print(f"\nTable {idx}:")
        print(f"Rows: {start_row} to {end_row}")
        print("Preview:")
        print(df.iloc[table['start']:table['start']+3])  # Show first 3 rows of each table
    
    # Ask user to select table
    while True:
        try:
            table_idx = int(input("\nSelect table number: ")) - 1
            if 0 <= table_idx < len(tables):
                selected_table = tables[table_idx]
                break
            else:
                print("Invalid selection. Please try again.")
        except ValueError:
            print("Please enter a valid number.")
    
    # Work with selected table
    table_df = df.iloc[selected_table['start']:selected_table['end']+1].reset_index(drop=True)
    
    # If the first row contains headers, use it as column names
    table_df.columns = table_df.iloc[0]
    table_df = table_df.iloc[1:].reset_index(drop=True)
    
    # Display first column values
    first_column = table_df.iloc[:, 0].dropna()
    print("\nAvailable values from first column:")
    for idx, value in enumerate(first_column, 1):
        print(f"{idx}. {value}")
    
    # Ask user to select value from first column
    while True:
        try:
            value_idx = int(input("\nSelect value number: ")) - 1
            if 0 <= value_idx < len(first_column):
                selected_row = table_df.iloc[value_idx]
                break
            else:
                print("Invalid selection. Please try again.")
        except ValueError:
            print("Please enter a valid number.")
    
    # Display all values from selected row
    print("\nSelected row values:")
    for column, value in selected_row.items():
        print(f"{column}: {value}")

if __name__ == "__main__":
    main()
