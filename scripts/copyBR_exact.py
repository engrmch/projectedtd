import os
import xlwings as xw  # Import xlwings for Excel handling

# Directory containing the source file
source_directory = os.path.join(os.path.expanduser("~"), "Desktop", "ProjectedTD")

# Target file
target_file = os.path.join(source_directory, 'ntr.xlsx')

print("Searching for the source file...")

# Search for a file with "BR", "Bill Run", or "TD" in its name
source_file = None
for file in os.listdir(source_directory):
    if ("BR" in file or "Bill" in file or "TD" in file) and file.endswith(('.xlsx', '.xlsb')):
        source_file = os.path.join(source_directory, file)
        break

if source_file:
    print(f"Found source file: {source_file}")
    
    # Open Excel with xlwings (headless mode)
    app = xw.App(visible=False)
    wb_source = app.books.open(source_file)
    
    # Open target file or create if not exists
    if not os.path.exists(target_file):
        wb_target = app.books.add()
        wb_target.save(target_file)
    
    wb_target = app.books.open(target_file)

    # Copy each sheet exactly
    for sheet in wb_source.sheets:
        sheet_name = sheet.name
        print(f"Copying sheet: {sheet_name}")

        # Delete existing sheet in the target file if it exists
        try:
            wb_target.sheets[sheet_name].delete()
        except:
            pass  # Ignore if the sheet does not exist

        # Copy sheet to target file
        sheet.copy(after=wb_target.sheets[-1])

        # Access the copied sheet
        copied_sheet = wb_target.sheets[sheet_name]

        # Find 'DIST_NODE' column
        headers = copied_sheet.range("A1").expand("right").value
        if "DIST_NODE" in headers:
            col_index = headers.index("DIST_NODE") + 1  # Convert to Excel 1-based index
            col_letter = xw.utils.col_name(col_index)

            # Get last row number
            last_row = copied_sheet.range(f"A{copied_sheet.cells.last_cell.row}").end("up").row

            # Add "Exempt" column next to last column
            exempt_col = len(headers) + 1  # Next available column
            exempt_letter = xw.utils.col_name(exempt_col)

            copied_sheet.range(f"{exempt_letter}1").value = "Exempt"  # Add header
            copied_sheet.range(f"{exempt_letter}2:{exempt_letter}{last_row}").formula = \
                f"=VLOOKUP({col_letter}2,NTR_60!A:F,6,FALSE)"

            print(f"'Exempt' column added in '{sheet_name}' with VLOOKUP formula.")

    # Save & Close
    wb_target.save()
    wb_source.close()
    wb_target.close()
    app.quit()
    
    print("Sheets copied successfully. 'Exempt' column added where applicable.")
else:
    print("No matching source file found.")
