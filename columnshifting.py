import pandas as pd
import numpy as np

# The input file path
input_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/final_cleaned_data.xlsx'
output_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/final_adjusted_data.xlsx'

# Read the Excel sheet
df = pd.read_excel(input_file_path)

# Process each row and collect processed rows
processed_rows = []

for index, row in df.iterrows():
    row_data = row.dropna().tolist()
    # Ensure we have enough data to shift, adjust this number as necessary
    if len(row_data) >= 47:  # This number may need to be adjusted
        nans_to_append = [np.nan] * (len(df.columns) - len(row_data))
        adjusted_row = row_data + nans_to_append
        processed_rows.append(adjusted_row)
    else:
        processed_rows.append(row.tolist())

# Create a new DataFrame from the processed rows
adjusted_df = pd.DataFrame(processed_rows, columns=df.columns)

# Adjust the column names as necessary, trimming or expanding to match the actual DataFrame size
adjusted_df = adjusted_df.iloc[:, :15]  # Adjust this based on the actual number of columns you need
adjusted_df.columns = [
    'Date', 'Currency', 'Companies', 'Bloomberg Ticker', 'Opening ', 'LTP',
    'Closing', 'Price Change (%)', 'Previous Price ', 'Volume traded(shares)',
    'Value traded ', 'Shares In Issue(m n\'s)', 'Market Cap ',
    'Market Cap ', 'Price Change US YTD(%)'
]

# Save the modified DataFrame back to an Excel file
adjusted_df.to_excel(output_file_path, index=False)

print("The file has been updated with columns shifted and renamed.")
