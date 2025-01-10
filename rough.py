# Load the existing Excel file
file_path = 'your_excel_file.xlsx'
workbook = openpyxl.load_workbook(file_path)
# Select the active sheet (or specify a sheet by name)
sheet = workbook.active  # or workbook['Sheet1']
# Define the data you want to insert (fill one by one)

data_to_insert = ['Alice', 'Bob', 'Charlie', 'David']
# Iterate over the cells to find empty ones and fill them
row, col = 1, 1  # Start from row 1, column 1 (adjust as needed)


# Loop over the data and fill it in empty cells one by one
for data in data_to_insert:
    # Find the next empty cell
    while sheet.cell(row=row, column=col).value:  # Skip non-empty cells
        row += 1  # Move down to the next row
    # Fill the empty cell with data
    sheet.cell(row=row, column=col).value = data
    # Move to the next row for the next entry
    row += 1
# Save the modified workbook
workbook.save('modified_excel_file.xlsx')