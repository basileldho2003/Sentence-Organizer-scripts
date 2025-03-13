import openpyxl

def read_excel(file_path):
    """Reads sentences from an Excel file while preserving the column structure."""
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active  # Get the first sheet
    data = []

    for row in sheet.iter_rows(values_only=True):
        data.append(
            [cell.strip() if isinstance(cell, str) else cell for cell in row]
        )  # Strip spaces from strings

    return data

def remove_duplicates_preserve_columns(data):
    """Removes duplicate sentences while maintaining column structure."""
    seen = set()
    deduplicated_data = []

    for row in data:
        deduplicated_row = []
        for cell in row:
            if cell and cell not in seen:
                seen.add(cell)
                deduplicated_row.append(cell)
            else:
                deduplicated_row.append("")  # Keep empty cell if it's a duplicate
        deduplicated_data.append(deduplicated_row)

    return deduplicated_data


def write_excel(file_path, data):
    """Writes deduplicated sentences back to an Excel file."""
    wb = openpyxl.Workbook()
    sheet = wb.active

    for row in data:
        sheet.append(row)

    wb.save(file_path)

# Example usage
input_file = "Sentence_dataset.xlsx"  # Change to "data.csv" if using CSV
output_file = "Sentence_dataset_deduplicated.xlsx"  # Change to "deduplicated_sentences.csv" for CSV

data = read_excel(input_file)
unique_data = remove_duplicates_preserve_columns(data)
write_excel(output_file, unique_data)

print(f"Deduplicated sentences retained in respective columns. Saved to {output_file}")
