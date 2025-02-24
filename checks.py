from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.formatting.rule import CellIsRule

# Load the existing workbook or create a new one
workbook = Workbook()  # Replace with your filename
sheet = workbook.create_sheet("unconditional")
workbook.remove(workbook.active)  # Or specify the sheet by name

# Define the cell range where you want to apply the conditional formatting
cell_range = "A1:A23"  # Change this range as needed

# Create a light blue fill pattern
light_blue_fill = PatternFill(start_color='ADD8E6', end_color='ADD8E6', fill_type='solid')

# Add conditional formatting rule
sheet.conditional_formatting.add(
    cell_range,
    CellIsRule(operator='equal', formula=['"P"'], fill=light_blue_fill)
)

# Save the changes to the workbook
workbook.save('your_file_with_conditional_formatting.xlsx')  # Save with a new filename