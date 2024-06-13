import pandas as pd

# Input file path (the output from the last step)
input_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/cleaned_combined_data_with_dates.xlsx'
output_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/cleaned_combined_data_with_currency.xlsx'

# Read the Excel sheet
df = pd.read_excel(input_file_path)

# Initialize an empty list to store currency values
currencies = []

# Placeholder for the current currency mode
current_currency = None

# Iterate through each row to check for currency identifiers
for index, row in df.iterrows():
    if 'RTGSc' in str(row.values):
        current_currency = 'RTGS'
    elif 'USc' in str(row.values):
        current_currency = 'USD'
    currencies.append(current_currency)

# Insert the currency list as a new column right after the 'Date' column
# Assuming 'Date' is at index 0, we insert 'Currency' at index 1
df.insert(1, 'Currency', currencies)

# Save the modified DataFrame back to an Excel file
df.to_excel(output_file_path, index=False)

print("The file has been updated with a 'Currency' column.")
