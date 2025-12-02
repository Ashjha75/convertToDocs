import csv
from docx import Document

doc = Document("testcases.docx")

rows = []
for table in doc.tables:
    for i, row in enumerate(table.rows):
        cells = [c.text.strip() for c in row.cells]

        # skip header row
        if i == 0:
            continue

        title = cells[1]                 # Test Case
        steps = cells[2]                 # Procedures / Steps
        expected = cells[3]              # Expected Outcome

        rows.append([title, steps, expected])

with open("testlodge_import.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(["Title", "Test Steps", "Expected Result"])
    writer.writerows(rows)
