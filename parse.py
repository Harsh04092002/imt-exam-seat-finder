import openpyxl   # library to read Excel files
import re         # library to extract text using patterns (regex)

# Step 1: Open the Excel file and get the sheet
wb = openpyxl.load_workbook("Seating_Plan_End_Term.xlsx", data_only=True)
ws = wb.active   # "active" = the first/only sheet

# Step 2: Convert the sheet into a simple list of rows
# Each row itself is a list of cell values, e.g. ['A', 'B', 'C', ...]
rows = list(ws.iter_rows(values_only=True))

print(f"Total rows in sheet: {len(rows)}")

# Step 3: Find every row that looks like a "header" row
# We use regex to pull out Subject, Date, Time, Hall from that one messy string
header_pattern = re.compile(
    r"Subject\(s\):\s*(.+?)\s{2,}Date:\s*(.+?)\s{2,}Time:\s*(.+?)\s{2,}Exam Hall:\s*(\S+)"
)

print(f"Total rows in sheet: {len(rows)}")

# Step 3: Find every "header" row and read the 5 seat rows below it
header_pattern = re.compile(
    r"Subject\(s\):\s*(.+?)\s{2,}Date:\s*(.+?)\s{2,}Time:\s*(.+?)\s{2,}Exam Hall:\s*(\S+)"
)

records = []  # this will hold one entry per student per exam

for i, row in enumerate(rows):
    cell = row[0]
    if cell and isinstance(cell, str):
        match = header_pattern.search(cell)
        if match:
            subject, date, time, hall = match.groups()

            # The column-letter row is 2 rows below the header (skip 1 blank row)
            letters_row = rows[i + 2]
            col_letters = letters_row[1:10]  # columns B..J hold 'A'..'I'

            # The 5 seat rows come right after the letters row
            for seat_row in rows[i + 3: i + 3 + 5]:
                row_number = seat_row[0]       # 1..5
                seats = seat_row[1:10]         # roll numbers for A..I
                for col_letter, roll in zip(col_letters, seats):
                    if roll:  # skip empty seats
                        records.append({
                            "roll": str(roll).strip(),
                            "subject": subject.strip(),
                            "date": date.strip(),
                            "time": time.strip(),
                            "hall": hall.strip(),
                            "seat": f"{row_number}{col_letter}",
                        })

print(f"Total seat records extracted: {len(records)}")
print("Sample record:", records[0])

# Step 5: Save everything as JSON, grouped by roll number, for the website to use
import json
from collections import defaultdict

grouped = defaultdict(list)
for r in records:
    grouped[r["roll"]].append(r)

with open("seating_data.json", "w") as f:
    json.dump(grouped, f, indent=2)

print(f"\nSaved data for {len(grouped)} students to seating_data.json")
