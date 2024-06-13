import pandas as pd

# The input file path
input_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/final_cleaned_data.xlsx'

# List of identifiers
identifiers = [
    "afdis", "ariston", "art", "bridgefort", "bridgefortclassb", "bat", "border", 
    "cafca", "cbz", "cfi", "delta", "dairiboard", "ecocash", "econet", "edgars", "fbc",
    "fidelitylife", "firstmutual", "firstmutualproperties", "gbholdings", "hippo", 
    "lafarge", "mash", "masimba", "meikles", "nampak", "nmb", "nts", "okzimbabwe", 
    "oldmutual", "ppc", "proplastics", "rtg", "seedco", "starafrica", "tanganda", 
    "truworths", "tsl", "turnall", "unifreight", "willdale", "zbfh", "zeco", "zhl", 
    "zimpapers", "hwange", "riozim"
]

# Read the Excel sheet
df = pd.read_excel(input_file_path)

# Output file path
output_file_path = 'C:/Users/horon/OneDrive/Desktop/IH Securities/Price Sheet/Identifiers_Separated_Data.xlsx'

# Create a Pandas Excel writer using XlsxWriter as the engine
with pd.ExcelWriter(output_file_path, engine='xlsxwriter') as writer:
    for identifier in identifiers:
        # Filter the DataFrame for the current identifier (case-insensitive match)
        filtered_df = df[df['Counter Identifier'].str.lower() == identifier.lower()]
        
        # If there are rows for the current identifier, write to a separate sheet
        if not filtered_df.empty:
            filtered_df.to_excel(writer, sheet_name=identifier.upper(), index=False)

print("Each identifier's data has been separated into its own sheet within the Excel workbook.")
