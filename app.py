import streamlit as st
import pdfplumber
import pandas as pd
import io

# 1. Set up the Web Page
st.set_page_config(page_title="Salary Register Converter", layout="wide")
st.title("📄 PDF Salary Register to Excel Converter")
st.markdown("Upload your monthly PDF salary registers to generate a consolidated yearly Excel report.")

# 2. File Uploader Interface
uploaded_files = st.file_uploader("Upload PDF files", type="pdf", accept_multiple_files=True)

if uploaded_files:
    if st.button("Process Files"):
        all_data = []
        progress_bar = st.progress(0)
        
        # 3. Loop through uploaded files
        for idx, file in enumerate(uploaded_files):
            file_name = file.name # E.g., "drexel-mar 26 wage reg.pdf"
            
            with pdfplumber.open(file) as pdf:
                for page in pdf.pages:
                    # Extract the table grid from the page
                    table = page.extract_table()
                    
                    if table:
                        # Convert the extracted table into a Pandas DataFrame
                        # We use the first row as headers, and the rest as data
                        df = pd.DataFrame(table[1:], columns=table[0])
                        
                        # Add a column to track which month/year this data came from
                        df['Source_Month_File'] = file_name
                        all_data.append(df)
            
            # Update the progress bar
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        if all_data:
            # 4. Combine all pages and months into one large table
            final_df = pd.concat(all_data, ignore_index=True)
            
            # --- DATA CLEANING SECTION ---
            # Note: Because the Drexel PDF has stacked data (e.g., Name and Designation in one box),
            # the text is extracted with newlines (\n). You clean it like this:
            
            # Example: Clean the "Employee Name \n Designation" column
            # final_df[['Employee Name', 'Designation']] = final_df['Employee Name / Designation'].str.split('\n', expand=True)
            
            # Remove empty rows or rows that just say "Total"
            final_df.dropna(how='all', inplace=True)
            
            st.success("Files processed successfully!")
            st.dataframe(final_df.head(10)) # Show a preview on the website
            
            # 5. Convert to Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                final_df.to_excel(writer, index=False, sheet_name='Yearly_Consolidated')
            output.seek(0)
            
            # 6. Provide the Download Button
            st.download_button(
                label="📥 Download Consolidated Excel",
                data=output,
                file_name="Yearly_Salary_Report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.warning("No tabular data could be extracted from the uploaded files.")