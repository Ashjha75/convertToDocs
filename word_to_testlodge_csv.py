#!/usr/bin/env python3
"""
word_to_testlodge_csv.py
Parse MS Word tables into CSV suitable for TestLodge import.
Produces: testlodge_import.csv
Debug summary printed to stdout.
"""

# RUN COMMAND:
# python word_to_testlodge_csv.py testcases.docx

import sys
import csv
import re
from docx import Document
from pathlib import Path

HEADER_KEYWORDS = ("test case", "test case id", "test procedures", "expected outcome", "pass/fail")

def normalize_text_from_cell(cell):
    """
    Preserve paragraph breaks as \n, remove excessive whitespace and common bullets.
    """
    parts = []
    for p in cell.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        # replace common bullet characters with hyphen
        text = re.sub(r'^[\u2022\u00B7\-\*\•\•\u2023\•\u25E6]\s*', '- ', text)
        # collapse multiple spaces
        text = re.sub(r'\s+', ' ', text)
        parts.append(text)
    return "\n".join(parts).strip()

def row_is_header(cells_text):
    combined = " ".join(cells_text).lower()
    return any(k in combined for k in HEADER_KEYWORDS)

def row_is_blank(cells_text):
    # consider row blank if all cells are empty after stripping
    return all(not (c and c.strip()) for c in cells_text)

def main(docx_path):
    docx_path = Path(docx_path)
    if not docx_path.exists():
        print("File not found:", docx_path)
        return 1

    doc = Document(str(docx_path))

    out_rows = []
    total_rows_scanned = 0
    skipped_header_rows = 0
    skipped_blank_rows = 0
    bad_rows = []

    for table_index, table in enumerate(doc.tables, start=1):
        for r_index, row in enumerate(table.rows):
            total_rows_scanned += 1
            cells_text = [c.text.strip() for c in row.cells]

            # skip repeated header rows
            if row_is_header(cells_text):
                skipped_header_rows += 1
                continue

            # skip blank rows
            if row_is_blank(cells_text):
                skipped_blank_rows += 1
                continue

            # Defensive: ensure expected minimum number of cells (some tables may have merged cells)
            # Your table layout expects at least 4 columns: [ID, Title, Steps, Expected, Pass/Fail optional]
            # We'll attempt to map based on length:
            try:
                # prefer mapping when there are 5 or more cells:
                if len(cells_text) >= 5:
                    title_cell = row.cells[1]
                    steps_cell = row.cells[2]
                    expected_cell = row.cells[3]
                # if 4 cells (maybe no Pass/Fail column)
                elif len(cells_text) == 4:
                    title_cell = row.cells[1]
                    steps_cell = row.cells[2]
                    expected_cell = row.cells[3]
                # if 3 cells, assume [Title, Steps, Expected]
                elif len(cells_text) == 3:
                    title_cell = row.cells[0]
                    steps_cell = row.cells[1]
                    expected_cell = row.cells[2]
                else:
                    # fallback: try to locate likely columns by content heuristics
                    # find the cell that looks like title (shorter) and steps/expected (longer)
                    lengths = [len(t) for t in cells_text]
                    if not lengths:
                        raise ValueError("no cells")
                    idx_sorted = sorted(range(len(lengths)), key=lambda i: lengths[i])
                    # smallest as title, next as steps, largest as expected (heuristic)
                    title_cell = row.cells[idx_sorted[0]]
                    steps_cell = row.cells[idx_sorted[1]] if len(idx_sorted) > 1 else row.cells[idx_sorted[0]]
                    expected_cell = row.cells[idx_sorted[-1]]
            except Exception as ex:
                bad_rows.append((table_index, r_index, cells_text, str(ex)))
                continue

            title = normalize_text_from_cell(title_cell)
            steps = normalize_text_from_cell(steps_cell)
            expected = normalize_text_from_cell(expected_cell)

            # If title empty, treat as bad/skip
            if not title:
                bad_rows.append((table_index, r_index, cells_text, "empty title"))
                continue

            out_rows.append([title, steps, expected])

    # write CSV with proper quoting
    csv_path = Path("testlodge_import.csv")
    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
        writer.writerow(["Title", "Test Steps", "Expected Result"])
        for r in out_rows:
            writer.writerow(r)

    # Debug summary
    print("Parsed document:", docx_path)
    print("Total table rows scanned:", total_rows_scanned)
    print("Output rows (test cases written):", len(out_rows))
    print("Header rows skipped:", skipped_header_rows)
    print("Blank rows skipped:", skipped_blank_rows)
    print("Bad/Skipped rows (count):", len(bad_rows))
    if bad_rows:
        print("Sample bad rows (table_index, row_index, cells_preview, reason):")
        for sample in bad_rows[:10]:
            ti, ri, cells_preview, reason = sample
            preview = " | ".join([c[:60].replace("\n", " ") for c in cells_preview])
            print(f"  table#{ti} row#{ri} -- {preview} -- {reason}")

    print("\nCSV written to:", csv_path.resolve())
    return 0

if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python word_to_testlodge_csv.py path/to/testcases.docx")
        sys.exit(2)
    sys.exit(main(sys.argv[1]))
