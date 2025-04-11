import shutil
import datetime
import os

# Define the source file and the new file name
source_file = os.path.join(os.path.expanduser("~"), "Desktop", "ProjectedTD", "ntr.xlsx")
current_date = datetime.datetime.now().strftime("%Y%m%d")
destination_file = os.path.join(
    os.path.expanduser("~"), "Desktop", "ProjectedTD", f"TD_{current_date}.xlsx"
)

# Duplicate and rename the file
try:
    shutil.copy(source_file, destination_file)
    print(f"File duplicated and renamed to {destination_file}")
except FileNotFoundError:
    print(f"Source file not found: {source_file}")
except Exception as e:
    print(f"An error occurred: {e}")
