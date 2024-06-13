import pandas as pd

# The input file path (the output from the previous operation)
input_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/cleaned_combined_data_with_currency.xlsx'
output_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/final_cleaned_data.xlsx'

# Read the Excel sheet
df = pd.read_excel(input_file_path)

# Step 2: Create a counter identifier column
# Assuming the counter names are in the 3rd column, which is at index 2
df['Counter Identifier'] = df.iloc[:, 2].str.replace(' ', '').str.lower()

# List of valid counters after cleaning (removing spaces and converting to lowercase)
valid_counters = [
    "afdis", "ariston", "art", "bridgefort", "bridgefortclassb", "bat", "border", 
    "cafca", "cbz", "cfi", "delta", "dairiboard", "ecocash", "econet***", "edgars", "fbc",
    "fidelitylife", "firstmutual", "firstmutualproperties", "gbholdings", "hippo", 
    "lafarge", "mash", "masimba", "meikles", "nampak", "nmb", "nts", "okzimbabwe", 
    "oldmutual", "ppc", "proplastics", "rtg", "seedco", "starafrica", "tanganda", 
    "truworths", "tsl", "turnall", "unifreight", "willdale", "zbfh", "zeco", "zhl", 
    "zimpapers", "hwange", "riozim"
]

# Step 3: Filter rows
df_filtered = df[df['Counter Identifier'].isin(valid_counters)]

# Save the modified DataFrame back to an Excel file
df_filtered.to_excel(output_file_path, index=False)

print("The file has been updated and saved with the filtered counters.")
