import streamlit as st
import pandas as pd
import openpyxl
import xlrd
import os
from PIL import Image
import io

def find_nearest_table_above(df, match_row_idx, match_col_idx):
    """Find the nearest 'table' cell above the matched cell"""
    try:
        # Look for 'table' cell above the match in the same column
        for idx in range(match_row_idx - 1, -1, -1):  # Search upwards
            cell_value = str(df.iloc[idx, match_col_idx]).lower().strip()
            if cell_value == 'table':
                # Get the header (row immediately below 'table')
                header_row = idx + 1
                if header_row < match_row_idx:  # Make sure header is above match
                    return header_row
        
        # If not found in same column, search all columns above
        for idx in range(match_row_idx - 1, -1, -1):  # Search upwards
            row_values = df.iloc[idx].astype(str).str.lower().str.strip()
            if row_values.isin(['table']).any():
                header_row = idx + 1
                if header_row < match_row_idx:  # Make sure header is above match
                    return header_row
    except Exception as e:
        st.warning(f"Error finding table header: {str(e)}")
    return None

def get_table_data(df, header_row, match_row):
    """Get the entire table data starting from header row"""
    try:
        # Get headers from the row below 'table'
        headers = df.iloc[header_row].tolist()
        
        # Get all data rows after headers
        table_data = df.iloc[header_row + 1:]
        
        # Convert to DataFrame with proper headers
        table_df = pd.DataFrame(table_data.values, columns=headers)
        
        return table_df
    except Exception as e:
        st.warning(f"Error extracting table data: {str(e)}")
        return None

def search_all_sheets(file_path, search_term):
    """Search across all sheets in the Excel file"""
    combined_results = None
    match_positions = []
    
    try:
        excel_file = pd.ExcelFile(file_path)
        
        for sheet_name in excel_file.sheet_names:
            try:
                # Read the entire sheet
                df_full = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
                
                # Search for term
                df_str = df_full.astype(str)
                
                # Find all matching cells
                for col in df_str.columns:
                    mask = df_str[col].str.contains(search_term, case=False, na=False)
                    match_rows = mask[mask].index
                    
                    for row_idx in match_rows:
                        # Find nearest table header above this cell
                        header_row = find_nearest_table_above(df_full, row_idx, col)
                        
                        if header_row is not None:
                            # Get the entire table
                            table_df = get_table_data(df_full, header_row, row_idx)
                            
                            if table_df is not None:
                                # Add sheet name column
                                table_df.insert(0, 'Sheet Name', sheet_name)
                                
                                # Store match position
                                match_positions.append({
                                    'sheet': sheet_name,
                                    'row_idx': row_idx - (header_row + 1),
                                    'col_idx': col + 1  # +1 because we added Sheet Name column
                                })
                                
                                # Combine with previous results
                                if combined_results is None:
                                    combined_results = table_df
                                else:
                                    # Make sure columns match
                                    if set(combined_results.columns) == set(table_df.columns):
                                        combined_results = pd.concat([combined_results, table_df], ignore_index=True)

            except Exception as e:
                st.warning(f"Error processing sheet {sheet_name}: {str(e)}")
                continue
    
    except Exception as e:
        st.error(f"Error searching Excel file: {e}")
        return None, None
    
    return combined_results, match_positions

def extract_images(file_path, sheet_name):
    """Extract images from the specified sheet"""
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        images = []
        
        # Try multiple methods to get images
        # Method 1: Try drawings
        if hasattr(sheet, 'drawings'):
            for drawing in sheet.drawings:
                try:
                    if hasattr(drawing, 'image'):
                        image_data = Image.open(io.BytesIO(drawing.image.ref))
                        images.append(image_data)
                except Exception:
                    pass
        
        # Method 2: Try _images
        if hasattr(sheet, '_images'):
            for img in sheet._images:
                try:
                    image_data = Image.open(io.BytesIO(img._data()))
                    images.append(image_data)
                except Exception:
                    pass
        
        return images
    except Exception as e:
        st.warning(f"Could not extract images from {sheet_name}: {str(e)}")
        return []

def main():
    st.title("Excel Search Tool")
    
    file_path = "ref.xlsx"
    
    if os.path.exists(file_path):
        search_term = st.text_input("Search across all tables:", "")
        
        if search_term:
            results, match_positions = search_all_sheets(file_path, search_term)
            
            if results is not None and not results.empty:
                st.write(f"Found matches in {len(match_positions)} locations")
                
                # Display images first
                unique_sheets = set(m['sheet'] for m in match_positions)
                for sheet_name in unique_sheets:
                    images = extract_images(file_path, sheet_name)
                    if images:
                        st.subheader(f"Images from {sheet_name}")
                        cols = st.columns(min(len(images), 3))  # Up to 3 columns
                        for idx, img in enumerate(images):
                            col_idx = idx % 3
                            with cols[col_idx]:
                                st.image(img, caption=f"Image {idx + 1}", use_column_width=True)
                
                # Then display the search results table
                st.subheader("Search Results:")
                # Remove Sheet Name column before styling
                display_results = results.drop(columns=['Sheet Name'])
                
                # Create style function for highlighting all matches
                def highlight_matches(x):
                    df_styler = pd.DataFrame('', index=x.index, columns=x.columns)
                    for match in match_positions:
                        # Adjust column index since we removed Sheet Name column
                        col_idx = match['col_idx'] - 1  # Subtract 1 because we removed Sheet Name
                        row_idx = match['row_idx']
                        df_styler.iloc[row_idx, col_idx] = 'background-color: yellow'
                    return df_styler
                
                # Display table with highlighting
                styled_table = display_results.style.apply(highlight_matches, axis=None)
                st.dataframe(styled_table)
                
                # Show summary
                st.subheader("Matches found in:")
                sheet_summary = pd.Series([m['sheet'] for m in match_positions]).value_counts()
                for sheet, count in sheet_summary.items():
                    st.write(f"- {sheet}: {count} matches")
            else:
                st.write("No matching records found")
    else:
        st.warning(f"Excel file not found at {file_path}")

if __name__ == "__main__":
    main()
