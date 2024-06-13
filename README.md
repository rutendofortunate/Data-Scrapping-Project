# Data-Scrapping-Project
On this project l downloaded daily price sheets from IH Securities in pdf form, extracted the tables into pdf to csv/json format files for further analysis.

## Steps to getting stock price data for the Zimbabwean Listed Companies.
1.	Download the price sheets from IH Securities. This process was manual because scrapping data is not ethical however will look for ways on how this can be automated.
2.	Extract the data on the PDFs to a single Excel sheet  (combined_data.xlsx). ZSE does provide daily price sheets in Excel however these are removed daily so an agreement might need to be devised on how to get these daily either from the site or by email or work with IH Securities to provide the Excel price sheets. extractdata2.py
3.	Clean the data (cleaned_combined_data.xlsx). cleaningdata.py
4.	Add dates to the cleaned file(cleaned_combined_data_with_dates.xlsx) by extracting the date from the Filename. cleaningdatatest.py
5.	Create a currency column and add a currency column. (cleaned_combined_data_with_currency.xlsx). currencysorting.py
6.	Clean and add counter identifiers (final_cleaned_data.xlsx). identifier.py
7.	Remove unneeded columns and spaces(final_adjusted_data.xlsx). columnshifting.py
8.	Separate RTGS (RTGS_Entries.xlsx) and USD (USD_Entries.xlsx) files. separating.py
9.	Or put RTGS and USD entries in their sheet respectively but one workbook. (Identifiers_Separated_Data.xlsx) newsheets.py

