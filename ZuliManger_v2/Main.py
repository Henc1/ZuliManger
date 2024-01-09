import csv
import openpyxl
from collections import Counter


def apply_filters(rows):
    # Filter 1: Remove rows without if there is no data in second column (addresses) 
    filtered_rows = [row for row in rows if row[1].strip() != '']

    # Filter 2: Custom filter - Copy data from column 2 (addresses) to column 1 (tag name) - if 1 is empty and 2 is not empty
    for row in filtered_rows:
        if row[0].strip() == '' and row[1].strip() != '':
            row[0] = row[1]

    # Filter 3: If the third (commment) column is empty, paste the text "Reserve"
    for row in filtered_rows:
        if row[3].strip() == '':
            row[3] = "Reserve"
#####
     # Check duplicity data
    values_col1 = Counter(row[0].strip() for row in filtered_rows)
    values_col2 = Counter(row[1].strip() for row in filtered_rows)

    # Print the duplicated values
    for value, count in values_col1.items():
        if count > 1:
            print(f"Duplicated Tags: {value}")
            
    for value, count in values_col2.items():
        if count > 1:
            print(f"Duplicated adresses: {value}")
            
    return filtered_rows
 

def csv_to_excel(input_csv, output_xlsx):
    with open(input_csv, 'r', newline='', encoding='ANSI') as csv_file:
        # Read CSV file
        csv_reader = csv.reader(csv_file)
        header = next(csv_reader)  # Assuming the first row is header
        rows = list(csv_reader)

        # Apply filters
        filtered_rows = apply_filters(rows)

        # Write to Excel file
        wb = openpyxl.Workbook()
        ws = wb.active

        # Define header cells
        header_cells = ['tag name', 'addresses', 'data type', 'comment']

        # Write custom header
        ws.append(header_cells)

        # Write original header
        ws.append(header)

        # Write filtered rows
        for row in filtered_rows:
            ws.append(row)

        # Save the Excel file
        wb.save(output_xlsx)

def main():
    # Replace 'your_input_file.csv' and 'your_output_file.xlsx' with your actual file names
    input_file = 'ZULI_ARG1_230918.SDF'
    output_file = 'output.xlsx'

    # Convert CSV to Excel with applied filters and custom headers
    csv_to_excel(input_file, output_file)

if __name__ == "__main__":
    main()
