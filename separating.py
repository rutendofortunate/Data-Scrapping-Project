import pandas as pd

# Input file path
input_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/final_cleaned_data.xlsx'

# Read the Excel sheet
df = pd.read_excel(input_file_path)

# Filter for RTGS and USD entries
rtgs_df = df[df['Currency'].str.upper() == 'RTGS']
usd_df = df[df['Currency'].str.upper() == 'USD']

# Output file paths
rtgs_output_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/RTGS_Entries.xlsx'
usd_output_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/USD_Entries.xlsx'

# Write RTGS and USD DataFrames to separate Excel files
rtgs_df.to_excel(rtgs_output_file_path, index=False)
usd_df.to_excel(usd_output_file_path, index=False)

print("RTGS and USD transactions have been separated into their own Excel files.")
