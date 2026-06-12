import sys
import csv
from collections import defaultdict
from datetime import datetime
from openpyxl import load_workbook


def should_ignore(sheet_name):
    ignored_prefixes = ("Sheet", "Documentation", "Report")
    return sheet_name.startswith(ignored_prefixes)


def find_target_sheet(workbook):
    all_sheets = workbook.sheetnames
    remaining_sheets = [s for s in all_sheets if not should_ignore(s)]

    data_sheets = [s for s in remaining_sheets if s.startswith("Data")]
    if data_sheets:
        return data_sheets[0]

    if len(remaining_sheets) == 1:
        return remaining_sheets[0]

    return None


def parse_datetime(value):
    if isinstance(value, datetime):
        return value
    return datetime.strptime(str(value).strip(), "%m.%d.%Y %H:%M:%S")


def main():
    if len(sys.argv) != 2:
        print(f"Usage: python {sys.argv[0]} <xlsx_file>")
        sys.exit(1)

    file_path = sys.argv[1]

    try:
        workbook = load_workbook(filename=file_path, read_only=True, data_only=True)
    except FileNotFoundError:
        print(f"Error: File not found: {file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"Error opening workbook: {e}")
        sys.exit(1)

    sheet_name = find_target_sheet(workbook)
    if not sheet_name:
        print("Error: Could not determine the correct sheet to process.")
        sys.exit(1)

    sheet = workbook[sheet_name]

    if sheet.max_column != 3:
        print(f"Error: Sheet '{sheet_name}' does not have exactly 3 columns (found {sheet.max_column}).")
        sys.exit(1)

    rows = sheet.iter_rows(values_only=True)

    header = next(rows, None)
    if header is None:
        print(f"Error: Sheet '{sheet_name}' is empty.")
        sys.exit(1)

    second_col_header = str(header[1]).strip()
    outputfile = second_col_header.split()[0].lower()
    output_csv = f"{outputfile}.csv"

    result = defaultdict(dict)
    all_sites = set()

    for row_num, row in enumerate(rows, start=2):
        if row is None or len(row) < 3:
            continue

        dt_val, site, data_val = row[0], row[1], row[2]

        if dt_val is None or site is None or data_val is None:
            continue

        try:
            dt = parse_datetime(dt_val)
            day = dt.date()
            site = str(site).strip()
            data_num = int(float(data_val))
        except Exception as e:
            print(f"Warning: Skipping invalid row {row_num}: {row} ({e})")
            continue

        all_sites.add(site)

        current = result[day].get(site)
        if current is None or data_num > current:
            result[day][site] = data_num

    sorted_sites = sorted(all_sites)
    sorted_dates = sorted(result.keys())

    with open(output_csv, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["date"] + sorted_sites)

        for day in sorted_dates:
            out_row = [day.isoformat()]
            for site in sorted_sites:
                value = result[day].get(site, "")
                out_row.append(value)
            writer.writerow(out_row)

    print(f"Processed sheet: {sheet_name}")
    print(f"Output written to: {output_csv}")


if __name__ == "__main__":
    main()