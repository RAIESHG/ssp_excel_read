import streamlit as st
import pandas as pd
import openpyxl
import xlrd
import os
from PIL import Image
import io

def search_all_sheets(file_path, search_term):
    """Search across all sheets in the Excel file"""
    all_results = []
    
    try:
        # Read all sheets
        excel_file = pd.ExcelFile(file_path)
        
        for sheet_name in excel_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            # Clean the data - replace NaN with empty string
            df = df.fillna('')
            
            # Convert all columns to string type for searching
            df_str = df.astype(str)
            
            # Create a mask that will be True for any cell containing the search term
            mask = df_str.apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
            
            # Get matching rows
            matches = df[mask]
            
            if not matches.empty:
                # Add sheet name and row number columns
                matches.insert(0, 'Sheet Name', sheet_name)
                matches.insert(1, 'Row Number', matches.index + 1)
                all_results.append(matches)
                
                # Extract and display images from sheets with matches
                try:
                    images = extract_images(file_path, sheet_name)
                    if images:
                        st.subheader(f"Images in {sheet_name}")
                        cols = st.columns(min(len(images), 3))
                        for idx, img in enumerate(images):
                            col_idx = idx % 3
                            with cols[col_idx]:
                                st.image(img, caption=f"Image {idx + 1}", use_column_width=True)
                except Exception as e:
                    st.warning(f"Could not extract images from {sheet_name}: {str(e)}")
    
    except Exception as e:
        st.error(f"Error searching Excel file: {e}")
        return None
    
    # Combine all results
    if all_results:
        return pd.concat(all_results, ignore_index=True)
    return None

def extract_images(file_path, sheet_name):
    """Extract images from the specified sheet"""
    try:
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
        images = []
        
        # Try to get images from drawings
        if hasattr(sheet, 'drawings'):
            for drawing in sheet.drawings:
                try:
                    if hasattr(drawing, 'image'):
                        image_data = Image.open(io.BytesIO(drawing.image.ref))
                        images.append(image_data)
                except Exception:
                    pass
        
        return images
    except Exception as e:
        st.warning(f"Could not extract images: {str(e)}")
        return []

def main():
    st.title("Excel Search Tool")
    
    # Define the path to the Excel file in your directory
    file_path = "ref.xlsx"
    
    if os.path.exists(file_path):
        # Search functionality
        search_term = st.text_input("Search across all sheets:", "")
        
        if search_term:
            results = search_all_sheets(file_path, search_term)
            
            if results is not None and not results.empty:
                st.write(f"Found {len(results)} matching rows across all sheets")
                
                # Display results with highlighting
                st.dataframe(results.style.apply(
                    lambda x: ['background-color: yellow' if search_term.lower() in str(v).lower() 
                             else '' for v in x], axis=1
                ))
                
                # Show summary by sheet
                st.subheader("Results by Sheet:")
                sheet_summary = results['Sheet Name'].value_counts()
                for sheet, count in sheet_summary.items():
                    st.write(f"- {sheet}: {count} matches")
            else:
                st.write("No matching records found")
    else:
        st.warning(f"Excel file not found at {file_path}")

if __name__ == "__main__":
    main()
