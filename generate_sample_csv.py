"""
generate_sample_csv.py
----------------------
Generates a realistic messy CSV file to test the CSV Cleaner Tool.
"""

import csv
from pathlib import Path

Path("samples").mkdir(exist_ok=True)

# Intentionally messy data
rows = [
    ["  First Name ", " Last Name", "Email Address", "  Date Joined", "  Country ", "Revenue"],
    ["Alice",         "Silva",      "alice@email.com",   "15/03/2023",  "Brazil",    "1200.50"],
    ["Bob",           "Jones",      "bob@email.com",     "2023-03-20",  "USA",       "980"],
    ["  Carlos  ",    "Souza",      "carlos@email.com",  "01/04/2023",  " Brazil ",  "2400.00"],
    ["Alice",         "Silva",      "alice@email.com",   "15/03/2023",  "Brazil",    "1200.50"],  # duplicate
    ["",              "",           "",                  "",            "",          ""],          # empty row
    ["Diana",         "Lee",        "diana@email.com",   "10/05/2023",  "UK",        "3100"],
    ["  Eduardo",     "Martins",    "edu@email.com",     "22/06/2023",  "Brazil",    "750.00"],
    ["",              "",           "",                  "",            "",          ""],          # empty row
    ["Fiona",         "Clark",      "fiona@email.com",   "07/07/2023",  "Canada",    "1850"],
    ["Bob",           "Jones",      "bob@email.com",     "2023-03-20",  "USA",       "980"],      # duplicate
    ["  Grace  ",     "Kim",        "grace@email.com",   "30/08/2023",  "  Korea  ", "2200.00"],
]

output = Path("samples/messy_customers.csv")
with open(output, "w", newline="", encoding="latin-1") as f:
    writer = csv.writer(f)
    writer.writerows(rows)

print(f"✅ Sample messy CSV created: {output}")
print(f"   Issues planted:")
print(f"   • Encoding: latin-1 (not UTF-8)")
print(f"   • Column names with leading/trailing spaces")
print(f"   • 2 fully duplicate rows")
print(f"   • 2 fully empty rows")
print(f"   • Cells with extra whitespace")
print(f"   • Mixed date formats (DD/MM/YYYY and YYYY-MM-DD)")
print(f"\nNow run: python clean_csv.py samples/messy_customers.csv")
