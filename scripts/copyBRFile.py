import os
import pandas as pd
from datetime import datetime
import xlwings as xw  # Import xlwings


# Directory containing the source file
source_directory = os.path.join(os.path.expanduser("~"), "Desktop", "ProjectedTD")

# Generate the target file name with current date
current_date = datetime.now().strftime('%Y%m%d')

# Original target file (unchanged)
target_file = os.path.join(source_directory, 'ntr.xlsx')

print("Copying contents of BR file...")

# Search for a file with "BR", "Bill Run", or "TD" in its name
source_file = None
for file in os.listdir(source_directory):
    if ("BR" in file or "Bill" in file or "TD" in file) and (file.endswith('.xlsx') or file.endswith('.csv') or file.endswith('.xlsb')):
        source_file = os.path.join(source_directory, file)
        break

if source_file:
    # Check if the source file is Excel (.xlsx), Binary Excel (.xlsb), or CSV
    if source_file.endswith('.xlsx'):
        sheets = pd.read_excel(source_file, sheet_name=None)  # Read all sheets
    elif source_file.endswith('.xlsb'):
        sheets = {}
        with pd.ExcelFile(source_file, engine='pyxlsb') as xlsb:
            for sheet_name in xlsb.sheet_names:
                sheets[sheet_name] = pd.read_excel(xlsb, sheet_name=sheet_name, engine='pyxlsb')  # Read all sheets
    else:  # Handle CSV files
        sheets = {'Sheet1': pd.read_csv(source_file, encoding='latin1')}  # Or use 'cp1252'
    print("Writing sheets to the target file...")

    # Write all sheets to the target file with the "Exempt" column
    with pd.ExcelWriter(target_file, mode='a', if_sheet_exists='replace') as writer:
        for sheet_name, data in sheets.items():
            # Find the 'DIST_NODE' column dynamically
            dist_node_col = None
            for col in data.columns:
                if col.lower() == 'dist_node':  # Match 'DIST_NODE' (case-insensitive)
                    dist_node_col = col
                    break

            if dist_node_col:
                dist_node_col_index = data.columns.get_loc(dist_node_col)  # Get the index of the column
                dist_node_col_letter = xw.utils.col_name(dist_node_col_index + 1)  # Get the column letter using xlwings
                
                # Apply the formula to each row in the 'Exempt' column
                data['Exempt'] = data.apply(
                    lambda cell: f"=VLOOKUP({dist_node_col_letter}{cell.name + 2},NTR_60!A:F,6,FALSE)", axis=1
                )  # Adjust cell name for Excel's 1-based indexing

            else:
                print(f"'{sheet_name}' does not contain a 'DIST_NODE' column.")

            # Write the updated sheet to the target file
            data.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"Copied sheet '{sheet_name}' from '{source_file}' to '{target_file}'.")

    print("All sheets have been copied successfully.")
else:
    print("No file with 'BR' or 'TD' in its name found.")



