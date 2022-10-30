from pathlib import Path  # core library

from openpyxl import load_workbook  # pip install openpyxl


# Define the folder where the files are located
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "INPUT"
output_dir = current_dir / "OUTPUT"
output_dir.mkdir(parents=True, exist_ok=True)

# List all the excel files in the folder
xl_files = list(input_dir.rglob("*.xls*"))

# Loop through all the files, read only the values and save them
for xl_file in xl_files:
    wb = load_workbook(xl_file, data_only=True)
    wb.save(output_dir / f"{xl_file.stem}_valuecopy.{xl_file.suffix}")