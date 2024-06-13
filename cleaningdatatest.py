import pandas as pd

# Specifying the file path
file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/cleaned_combined_data.xlsx'
output_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/cleaned_combined_data_with_dates.xlsx'

# Read the Excel sheet
df = pd.read_excel(file_path)

# Extract the date using regular expression
# The pattern \d{2}\.\d{2}\.\d{4} matches the date in DD.MM.YYYY format
df['Filename'] = df['Filename'].str.extract('(\d{2}\.\d{2}\.\d{4})')

# Rename the column
df.rename(columns={'Filename': 'Date'}, inplace=True)

# Save the modified DataFrame back to an Excel file
df.to_excel(output_file_path, index=False)

print("The file has been updated and saved with dates extracted.")
