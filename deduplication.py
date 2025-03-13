import openpyxl

def read_excel(file_path):
    """Reads sentences from an Excel file while preserving the column structure."""
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active  # Get the first sheet
    data = []

    for row in sheet.iter_rows(values_only=True):
        # Strip spaces and replace empty values with None
        cleaned_row = [cell.strip() if isinstance(cell, str) else None for cell in row]
        data.append(cleaned_row)

    return data

def remove_duplicates_and_shift_up(data):
    """Removes duplicate sentences while shifting non-empty cells upward in each column."""
    seen = set()

    # Transpose the data to process columns instead of rows
    num_cols = len(data[0])  # Get the number of columns
    columns = [[] for _ in range(num_cols)]

    # Populate column-wise data and remove duplicates
    for row in data:
        for col_idx, cell in enumerate(row):
            if cell and cell not in seen:
                seen.add(cell)
                columns[col_idx].append(cell)

    # Determine the max rows needed after shifting up
    max_rows = max(len(col) for col in columns)

    # Rebuild the new table by filling empty spaces at the bottom
    deduplicated_data = []
    for row_idx in range(max_rows):
        new_row = [
            columns[col_idx][row_idx] if row_idx < len(columns[col_idx]) else None
            for col_idx in range(num_cols)
        ]
        deduplicated_data.append(new_row)

    return deduplicated_data

def write_excel(file_path, data):
    """Writes deduplicated sentences back to an Excel file."""
    wb = openpyxl.Workbook()
    sheet = wb.active

    for row in data:
        sheet.append(row)

    wb.save(file_path)

# Example usage
input_file = "Sentence_dataset.xlsx"
output_file = "Sentence_dataset_deduplicated.xlsx"

# Read, process, and save the cleaned file
data = read_excel(input_file)
unique_data = remove_duplicates_and_shift_up(data)
write_excel(output_file, unique_data)

print(
    f"Deduplicated sentences saved with blank cells shifted up. File saved to: {output_file}"
)
