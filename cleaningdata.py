import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import re
from datetime import datetime

# Specify your Excel file path
file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/combined_data.xlsx'

# Load the Excel file
df = pd.read_excel(file_path)

# Move 'Filename' column to the first position
filename_col = df.pop('Filename')
df.insert(0, 'Filename', filename_col)

# Extract date from 'Filename' and create a new 'date' column with a recognized date format
df['date'] = df['Filename'].apply(lambda x: datetime.strptime(re.search(r'\d{2}\.\d{2}\.\d{4}', x).group(), '%d.%m.%Y'))

# Rename the columns after 'Filename' and 'date' to 'companies' and 'Bloomberg Ticker'
# Assuming 'Filename' is now column 0 and 'date' is column 1, 'companies' will be column 2, 'Bloomberg Ticker' column 3
df.columns = ['Filename', 'date'] + ['companies', 'Bloomberg Ticker'] + list(df.columns[4:])

# Ensure the dataframe has exactly 15 columns
if len(df.columns) > 15:
    df = df.iloc[:, :15]  # Adjust if necessary

# Removing completely empty columns might not be needed if all columns are filled as expected
df = df.dropna(how='all', axis=1)

# Save the cleaned dataframe back to an Excel file, using openpyxl for date formatting
wb = Workbook()
ws = wb.active

for r in dataframe_to_rows(df, index=False, header=True):
    ws.append(r)

# Apply date formatting for 'date' column, assuming it's the second column now (column 'B')
for cell in ws['B'][1:]:  # Skipping header row
    cell.number_format = 'YYYY-MM-DD'

# Save the modified file to a new location or overwrite the existing one if desired
clean_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/cleaned_combined_data.xlsx'
wb.save(clean_file_path)
