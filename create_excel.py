#!/usr/bin/env python3
"""Parse classes.md and generate an Excel workbook with two sheets."""

import re
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill

INPUT_FILE = "classes.md"
OUTPUT_FILE = "classes.xlsx"

# --- Parse the markdown file ---
with open(INPUT_FILE, "r", encoding="utf-8") as f:
    content = f.read()

# Find all "## Class of YYYY" sections (including those with no content)
class_pattern = re.compile(r"^## Class of (\d{4})\s*$", re.MULTILINE)
class_matches = list(class_pattern.finditer(content))

# Also find the "Years still missing full lists" section
missing_pattern = re.compile(r"Years still missing.*?\n\s*([\d,\s]+)")
missing_match = missing_pattern.search(content)
all_missing_raw = [
    y.strip() for y in missing_match.group(1).split(",")
] if missing_match else []
# Filter to only years in the 1975–2019 range
missing_years = [
    y for y in all_missing_raw if 1975 <= int(y) <= 2019
]

# Extract each class section
classes = []  # list of (year, champion_name_or_None, list_of_names)

for i, match in enumerate(class_matches):
    year = match.group(1)
    start = match.end()
    # End is start of next section or end of file
    end = class_matches[i + 1].start() if i + 1 < len(class_matches) else len(content)
    section = content[start:end]

    # Check for champion line: **Champion: Name(s)**
    champion_match = re.search(r"\*\*Champion:\s*(.+?)\*\*", section)
    champion = None
    if champion_match:
        champion = champion_match.group(1).strip()

    # Extract numbered pupil names
    names = []
    for line in section.splitlines():
        name_match = re.match(r"^\d+\.\s+(.+)$", line.strip())
        if name_match:
            name = name_match.group(1).strip()
            # Remove trailing "(late)" or similar markers from name for cleanliness
            names.append(name)

    classes.append((year, champion, names))

# --- Build Excel workbook ---
wb = Workbook()

# Sheet 1: All classes with Name, Class Year, Champion
ws1 = wb.active
ws1.title = "All Classes"

# Header row
headers = ["Name", "Class Year", "Champion"]
ws1.append(headers)
header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF")
for col, header in enumerate(headers, 1):
    cell = ws1.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font

# Data rows
for year, champion, names in classes:
    for name in names:
        is_champion = "Yes" if champion and name in champion.split("/") else "No"
        ws1.append([name, year, is_champion])

# Auto-adjust column widths
for column in ws1.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except Exception:
            pass
    ws1.column_dimensions[column_letter].width = min(max_length + 4, 50)

# Sheet 2: Missing classes/years
ws2 = wb.create_sheet(title="Missing Classes")

ws2.append(["Year", "Status"])
status_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
status_font = Font(color="9C0006")
for col, header in enumerate(["Year", "Status"], 1):
    cell = ws2.cell(row=1, column=col)
    cell.fill = header_fill
    cell.font = header_font

for year in missing_years:
    last_row = ws2.max_row
    ws2.append([year, "Missing full list"])
    # Highlight in red
    for cell in ws2[last_row + 1]:
        cell.fill = status_fill
        cell.font = status_font

# Auto-adjust column widths for sheet 2
for column in ws2.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except Exception:
            pass
    ws2.column_dimensions[column_letter].width = min(max_length + 4, 30)

wb.save(OUTPUT_FILE)
print(f"Excel file saved to {OUTPUT_FILE}")
print(f"  Sheet 1 ('All Classes'): {sum(len(names) for _, _, names in classes)} students across {len(classes)} classes")
print(f"  Sheet 2 ('Missing Classes'): {len(missing_years)} missing years")
