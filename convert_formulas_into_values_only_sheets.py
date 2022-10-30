from pathlib import Path  # core library

import xlwings as xw  # pip install xlwings


# Define the folder where the files are located
current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
input_dir = current_dir / "INPUT"
output_dir = current_dir / "OUTPUT"
output_dir.mkdir(parents=True, exist_ok=True)

# List all the excel files in the folder
xl_files = list(input_dir.rglob("*.xls*"))

# List all the sheets that should be converted
sheets = ["Sales"]

with xw.App(visible=False) as app:
    for xl_file in xl_files:
        wb = app.books.open(xl_file)
        for sheet in sheets:
            wb.sheets[sheet].used_range.value = wb.sheets[sheet].used_range.value
        wb.save(output_dir / f"{xl_file.stem}_valuecopy_only_sht.{xl_file.suffix}")