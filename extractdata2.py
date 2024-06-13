import camelot
import os
import pandas as pd
import re

# Directory path
directory_path = r'C:\Users\horon\OneDrive\Desktop\IH Securities\Price Sheet\PDFs'

# Function to process DataFrame to match the template format
def process_dataframe(df, template_columns=None):
    # Remove completely empty rows and columns
    df = df.dropna(how='all').dropna(axis=1, how='all')
    
    # If template_columns is provided, reorder/keep only those columns
    if template_columns is not None:
        # Reindex the DataFrame columns to match the template, adding missing columns with NaNs
        df = df.reindex(columns=template_columns, fill_value='')
    
    return df

# List to store all the extracted tables
all_tables = []
template_columns = None  # We'll determine this from the first file

# Iterate over each file in the directory
for filename in os.listdir(directory_path):
    if filename.endswith(".pdf"):
        file_path = os.path.join(directory_path, filename)
        # Extract tables from the PDF
        tables = camelot.read_pdf(file_path, pages='all', flavor='stream')
        
        # Loop through the extracted tables
        for i, table in enumerate(tables):
            # Convert table to DataFrame
            df = table.df
            
            # For the first table of the first file, determine the template structure
            if template_columns is None and i == 0:
                template_columns = df.columns.tolist()
            
            # Process DataFrame to match the template format
            df = process_dataframe(df, template_columns)
            
            # Add the filename as a column
            df['Filename'] = filename
            
            # Append the DataFrame to the list
            all_tables.append(df)

# Combine all tables into a single DataFrame
combineddata = pd.concat(all_tables, ignore_index=True)

# Save the combined DataFrame to an Excel file
excel_file_path = r'C:\Users\horon\OneDrive\Desktop\IH Securities\Price Sheet\combined_data.xlsx'
combineddata.to_excel(excel_file_path, index=False)

print(f'DataFrame has been saved to {excel_file_path}')
