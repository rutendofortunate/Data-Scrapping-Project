import json
import psycopg2
import re
from datetime import datetime

def to_numeric(value):
    """Attempt to convert a value to a numeric type; return None for non-numeric placeholders."""
    try:
        if value is None or isinstance(value, str) and (value.strip() in ['NaN', '-', 'SUSP', 'SUSPENDED', '(shares)', '(RTGS$)', 'Change (%)', 'Issue (mn\'s)', '(RTGS$ mn\'s)', '(US$ mn\'s)', 'RTGS YTD (%)', 'US$ YTD (%)'] or not value.strip()):
            return None
        return float(value.replace(',', ''))  # Ensure comma removal for thousands
    except ValueError:
        return None

def clean_text(value):
    """Return None for 'NaN' values and other known placeholders in text fields."""
    if value is None or value == 'NaN':
        return None
    return value

def extract_date_from_filename(filename):
    """Extracts a date from a filename using a regular expression."""
    match = re.search(r'\d{4}\d{2}\d{2}', filename)
    if match:
        return datetime.strptime(match.group(), '%Y%m%d').date()
    return None

# Database connection parameters
dbname = 'pricesheets'
user = 'postgres'
password = '#Capable123'
host = 'localhost'

# Connect to your PostgreSQL database
conn = psycopg2.connect(dbname=dbname, user=user, password=password, host=host)
cur = conn.cursor()

# Load and insert data from the JSON file
json_file_path = r"C:\Users\horon\OneDrive\Desktop\IH Securities\Price Sheet\Second_Tables.json"
try:
    with open(json_file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    for item in data:
        data_date = extract_date_from_filename(item['filename'])
        for row in item['table']:
            row_data = (
                clean_text(row.get('Unnamed: 0')), clean_text(row.get('Bloomberg')),
                to_numeric(row.get('Opening')), to_numeric(row.get('LTP')),
                to_numeric(row.get('Closing')), to_numeric(row.get('Price')),
                to_numeric(row.get('Previous')), to_numeric(row.get('Volume traded')),
                to_numeric(row.get('Value traded')), to_numeric(row.get('Shares In')),
                to_numeric(row.get('Market Cap')), to_numeric(row.get('Market Cap.1')),
                to_numeric(row.get('Price Change')), to_numeric(row.get('Price Change.1')),
                data_date
            )
            cur.execute("""
                INSERT INTO stock_data (
                    "Unnamed: 0", "Bloomberg", "Opening", "LTP", "Closing", 
                    "Price", "Previous", "Volume traded", "Value traded", 
                    "Shares In", "Market Cap", "Market Cap.1", "Price Change", "Price Change.1",
                    data_date
                ) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                """, row_data)
    conn.commit()
except psycopg2.Error as e:
    print(f"An error occurred while inserting data: {e}")
    conn.rollback()
except Exception as e:
    print(f"An error occurred: {e}")
finally:
    cur.close()
    conn.close()

print("Data import to PostgreSQL completed successfully.")
