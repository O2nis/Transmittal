import streamlit as st
import pandas as pd
from io import BytesIO
import re
from datetime import datetime

def clean_dates(df):
    """
    Convert all date columns to '11-May-25' format
    """
    for col in df.columns:
        # Try to convert to datetime
        try:
            df[col] = pd.to_datetime(df[col], errors='ignore')
            # If successful (datetime column), format it
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d-%b-%y').str.upper()
        except:
            pass
    return df

def update_excel(df, codes, date, transmittal, date_col, transmittal_col, code_col):
    """
    Update the Excel dataframe with the provided information
    """
    updated_rows = 0
    formatted_date = date.strftime('%d-%b-%y').upper()  # Format as 11-MAY-25
    
    for code in codes:
        # Find rows where the code matches (case insensitive, strip whitespace)
        mask = df[code_col].astype(str).str.strip().str.lower() == code.strip().lower()
        matching_rows = df[mask]
        
        if not matching_rows.empty:
            df.loc[mask, date_col] = formatted_date
            df.loc[mask, transmittal_col] = transmittal
            updated_rows += len(matching_rows)
    
    return df, updated_rows

def main():
    st.title("Excel Data Updater")
    st.write("Upload an Excel file, paste codes, and update corresponding rows with date and transmittal code.")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Read and clean dates on initial load
            df = pd.read_excel(uploaded_file)
            df = clean_dates(df)
            st.success("File successfully loaded and dates cleaned!")
            
            # Display preview
            st.subheader("File Preview")
            st.write(df.head())
            
            # Get column names
            columns = df.columns.tolist()
            
            # User inputs
            st.subheader("Update Parameters")
            
            # Column selection
            code_col = st.selectbox("Select the column containing codes to match:", columns)
            date_col = st.selectbox("Select the column to write the date to:", columns)
            transmittal_col = st.selectbox("Select the column to write the transmittal code to:", columns)
            
            # Data inputs
            date_value = st.date_input("Enter the date:")
            transmittal_value = st.text_input("Enter the transmittal code:")
            
            # Code input
            st.write("Paste codes (one per line or separated by commas):")
            codes_input = st.text_area("Codes", height=150)
            
            if st.button("Update Data"):
                if not codes_input:
                    st.warning("Please enter at least one code.")
                else:
                    # Parse codes (split by newline or comma)
                    codes = re.split(r'[\n,]', codes_input)
                    codes = [code.strip() for code in codes if code.strip()]
                    
                    # Update the dataframe
                    updated_df, updated_rows = update_excel(
                        df.copy(), codes, date_value, transmittal_value,
                        date_col, transmittal_col, code_col
                    )
                    
                    # Clean all dates again before output
                    updated_df = clean_dates(updated_df)
                    
                    if updated_rows > 0:
                        st.success(f"Successfully updated {updated_rows} rows!")
                        
                        # Show updated rows
                        st.subheader("Updated Rows Preview")
                        formatted_date = date_value.strftime('%d-%b-%y').upper()
                        mask = (updated_df[date_col] == formatted_date) & (updated_df[transmittal_col] == transmittal_value)
                        st.write(updated_df[mask].head())
                        
                        # Download button
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            updated_df.to_excel(writer, index=False)
                        output.seek(0)
                        
                        st.download_button(
                            label="Download Updated Excel File",
                            data=output,
                            file_name="updated_file.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("No matching codes found in the specified column.")
            
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

if __name__ == "__main__":
    main()
