import streamlit as st
import pandas as pd
import openpyxl
import xlrd
import os
from PIL import Image
import io

# Set up file type handling
EXCEL_TYPES = {
    "xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "xls": "application/vnd.ms-excel"
}

def get_excel_sheets(file_path):
    """Read all sheets from Excel file and return sheet names"""
    try:
        # Use appropriate engine based on file type
        if file_path.endswith('.xls'):
            excel_file = pd.ExcelFile(file_path, engine='xlrd')
        else:
            excel_file = pd.ExcelFile(file_path, engine='openpyxl')
        return excel_file.sheet_names
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
        return None

def identify_tables(df):
    """Identify tables in the DataFrame and return their positions"""
    tables = []
    current_table_start = None
    empty_row_count = 0
    min_rows = 3
    min_columns = 5
    
    def is_valid_data_row(row):
        valid_values = sum(1 for val in row if pd.notna(val) 
                         and str(val).strip() != ''
                         and val is not None)
        return valid_values >= min_columns
    
    def check_consecutive_rows(start_idx, min_rows):
        if start_idx + min_rows > len(df):
            return False
        for i in range(start_idx, start_idx + min_rows):
            if not is_valid_data_row(df.iloc[i]):
                return False
        return True
    
    for idx in range(len(df)):
        row = df.iloc[idx]
        is_empty = not is_valid_data_row(row)
        
        if not is_empty and current_table_start is None:
            if check_consecutive_rows(idx, min_rows):
                current_table_start = idx
                empty_row_count = 0
        elif is_empty and current_table_start is not None:
            empty_row_count += 1
            if empty_row_count >= 2:
                tables.append({
                    'start': current_table_start,
                    'end': idx - empty_row_count
                })
                current_table_start = None
                empty_row_count = 0
    
    if current_table_start is not None:
        tables.append({
            'start': current_table_start,
            'end': len(df) - 1
        })
    
    return tables

def search_table(df, search_term):
    """Search all columns in the DataFrame for the search term"""
    if not search_term:
        return df
    
    # Convert all columns to string type for searching
    df_str = df.astype(str)
    
    # Create a mask that will be True for any cell containing the search term
    mask = df_str.apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
    
    # Return rows where any column contains the search term
    return df[mask]

def extract_images(file_path, sheet_name):
    """Extract images from the specified sheet"""
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        images = []
        
        # Method 1: Try to get images from drawings
        if hasattr(sheet, 'drawings'):
            for drawing in sheet.drawings:
                try:
                    # Handle different types of drawings
                    if hasattr(drawing, 'image'):
                        image_data = Image.open(io.BytesIO(drawing.image.ref))
                        images.append(image_data)
                    elif hasattr(drawing, '_data'):
                        image_data = Image.open(io.BytesIO(drawing._data()))
                        images.append(image_data)
                except Exception as e:
                    st.warning(f"Could not extract drawing: {str(e)}")
        
        # Method 2: Try to get images directly from _images
        if hasattr(sheet, '_images'):
            for img in sheet._images:
                try:
                    image_data = Image.open(io.BytesIO(img._data()))
                    images.append(image_data)
                except Exception as e:
                    st.warning(f"Could not extract _image: {str(e)}")
        
        # Method 3: Try to get images from shapes
        if hasattr(sheet, 'shapes'):
            for shape in sheet.shapes:
                try:
                    if hasattr(shape, 'image'):
                        image_data = Image.open(io.BytesIO(shape.image.ref))
                        images.append(image_data)
                except Exception as e:
                    st.warning(f"Could not extract shape: {str(e)}")
        
        # Remove any duplicate images
        unique_images = []
        seen = set()
        for img in images:
            img_bytes = img.tobytes()
            if img_bytes not in seen:
                seen.add(img_bytes)
                unique_images.append(img)
        
        return unique_images
    
    except Exception as e:
        st.warning(f"Could not extract images: {str(e)}")
        return []

def main():
    st.title("Excel Table Viewer")
    
    # Define the path to the Excel file in your directory
    file_path = "ref.xlsx"  # Specify the path to your Excel file here
    
    if os.path.exists(file_path):
        # Get sheet names
        sheet_names = get_excel_sheets(file_path)
        
        if sheet_names:
            # Sheet selection
            selected_sheet = st.selectbox("Select a sheet:", sheet_names)
            
            # Extract and display images with error handling
            try:
                images = extract_images(file_path, selected_sheet)
                if images:
                    st.subheader("Images in Sheet")
                    cols = st.columns(min(len(images), 3))  # Create up to 3 columns
                    for idx, img in enumerate(images):
                        col_idx = idx % 3
                        with cols[col_idx]:
                            st.image(img, caption=f"Image {idx + 1}", use_column_width=True)
                else:
                    st.info("No images found in this sheet")
            except Exception as e:
                st.error(f"Error processing images: {str(e)}")
            
            # Read the selected sheet
            df = pd.read_excel(file_path, sheet_name=selected_sheet)
            
            # Identify tables and use the first one
            tables = identify_tables(df)
            
            if tables:
                # Get first table
                table = tables[0]
                table_df = df.iloc[table['start']:table['end']+1].reset_index(drop=True)
                
                # Use first row as headers
                table_df.columns = table_df.iloc[0]
                table_df = table_df.iloc[1:].reset_index(drop=True)
                
                # Clean the data - replace NaN with empty string
                table_df = table_df.fillna('')
                
                # Search functionality with highlighting
                search_term = st.text_input("Search in table:", "")
                
                if search_term:
                    filtered_df = search_table(table_df, search_term)
                    
                    if not filtered_df.empty:
                        st.write(f"Found {len(filtered_df)} matching rows")
                        
                        # Display the filtered results
                        st.dataframe(filtered_df.style.apply(
                            lambda x: ['background-color: yellow' if search_term.lower() in str(v).lower() 
                                     else '' for v in x], axis=1
                        ))
                    else:
                        st.write("No matching records found")
                else:
                    # If no search term, show original table
                    st.dataframe(table_df)
                    st.write(f"Showing all {len(table_df)} rows")
            else:
                st.warning("No tables found in the selected sheet")
        else:
            st.warning("No sheets found in the Excel file.")
    else:
        st.warning(f"Excel file not found at {file_path}")

if __name__ == "__main__":
    main()
