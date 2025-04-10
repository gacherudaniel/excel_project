import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from copy import copy
from openpyxl.cell.cell import MergedCell

# === CONFIGURATION ===
date_str = "14th March"
target_date = pd.to_datetime("2024-03-14")

# === Load Source Data ===
source_file = "C:\\Users\\Training 3\\DataScience\\Excel_project\\03 March  25 Gross Gen.xlsx"
summary_df = pd.read_excel(source_file, sheet_name="Summary", header=None)

# Extract data section and clean it
data_df = summary_df.iloc[9:].copy()

# Step 2: Reset index and rename columns manually
data_df = data_df.reset_index(drop=True)
data_df.columns = ["DATE", "DG SET .1", "DG SET .2", "DG SET .3", "DG SET .4", "DG SET .5", "STG", "DAILY TOTAL MWH", "PLANT GROSS","ENG GROSS","STG"]

data_df = data_df[pd.to_datetime(data_df["DATE"], errors='coerce').notna()]
data_df["DATE"] = pd.to_datetime(data_df["DATE"])

# Find the MWH value for the date
daily_row = data_df[data_df["DATE"] == target_date]
mwh_value = daily_row["DAILY TOTAL MWH"].values[0] if not daily_row.empty else None


# === Write to Daily Report ===
def process_excel(report_file, date_str, mwh_value):
    if mwh_value is not None:
        report_file = "C:\\Users\\Training 3\\DataScience\\Excel_project\\Daily production report March 2025.xlsx"
        wb = load_workbook(report_file)
        
        # Check if a sheet with the given date_str exists
        if date_str in wb.sheetnames:
            # If the sheet exists, update the value in cell B8
            ws = wb[date_str]
            ws["B8"] = mwh_value
            wb.save("C:\\Users\\Training 3\\DataScience\\Excel_project\\Daily production report March 2025.xlsx")
            print(f"Inserted {mwh_value} into sheet '{date_str}' cell B8.")
        
        # If the sheet doesn't exist, create a new sheet with the date_str name
        else:
            # Find an existing sheet to duplicate (use the first sheet here as an example)
            sheet_to_duplicate = wb.worksheets[2]  # You can adjust this if you want a specific sheet
            
            # Create a new sheet by copying the content of the original sheet
            new_sheet = wb.copy_worksheet(sheet_to_duplicate)
            
            # Rename the new sheet to the current date_str
            new_sheet.title = date_str

            for row in sheet_to_duplicate.iter_rows():
                for cell in row:
                # Skip MergedCells that are not the actual top-left anchor
                    if isinstance(cell, MergedCell):
                        continue

                    new_cell = new_sheet.cell(row=cell.row, column=cell.column)
                                            
                    # Copy styles
                    if cell.has_style:
                        new_cell.font = copy(cell.font)
                        new_cell.border = copy(cell.border)
                        new_cell.fill = copy(cell.fill)
                        new_cell.number_format = copy(cell.number_format)
                        new_cell.protection = copy(cell.protection)
                        new_cell.alignment = copy(cell.alignment)

                
                    # Copy only formulas or static labels (no user-filled values)
                    if cell.data_type == 'f':
                        new_cell.value = f"={cell.value}"
                    elif isinstance(cell.value, str) and cell.value.strip() != "":
                        new_cell.value = cell.value  # Copy headers/static text
                    else:
                        new_cell.value = None  # Clear user-entered numbers or blanks

            # Copy merged cell ranges
            for merged_range in sheet_to_duplicate.merged_cells.ranges:
                new_sheet.merge_cells(str(merged_range))
            # Copy column widths
            for col in sheet_to_duplicate.column_dimensions:
                new_sheet.column_dimensions[col].width = sheet_to_duplicate.column_dimensions[col].width

            # Copy row heights
            for row_dim in sheet_to_duplicate.row_dimensions:
                new_sheet.row_dimensions[row_dim].height = sheet_to_duplicate.row_dimensions[row_dim].height


            
            
            # Assign the new sheet to `ws` so we can insert value below
            ws = new_sheet
            print(f"New sheet '{date_str}' has been created.")

        # Insert the value into B8, regardless of whether the sheet existed or was just created
        ws["B8"] = mwh_value
        wb.save("Daily production report March 2025 - updated.xlsx")
        print(f"Inserted {mwh_value} into sheet '{date_str}' cell B8.")
    else:
        print("No data found for that date.")