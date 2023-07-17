import xlwings as xw

# Load File1.xlsm
file1_path = r'F:\PythonProjects\lookup-and-add\input_file\file_input1.xlsm'
file1_workbook = xw.Book(file1_path)
file1_sheet = file1_workbook.sheets['Record']

# Load File2.xlsx
file2_path = r'F:\PythonProjects\lookup-and-add\input_file\file_input2.xlsx'
file2_workbook = xw.Book(file2_path)
file2_sheet = file2_workbook.sheets['Record']

# Get the values from col1 in File2.xlsx
file2_col1_values = file2_sheet.range('A2:A' + str(file2_sheet.range('A' + str(file2_sheet.cells.last_cell.row)).end('up').row)).value
file1_col1_values = file1_sheet.range('A2:A' + str(file1_sheet.range('A' + str(file1_sheet.cells.last_cell.row)).end('up').row)).value
# Iterate over the values in col1 of File2.xlsx
for value in file2_col1_values:
    # Convert the value to string
    value = str(value)  # Convert the value to string before subscripting

    # Check if the value already exists in col1 of File1.xlsm

    if any(value == str(val[0]) for val in file1_col1_values if isinstance(val, tuple)):
        print(f"Value '{value}' already exists in File1.xlsm")
    else:
        # Find the complete row in File2.xlsx based on the value
        for row in range(2, file2_sheet.cells.last_cell.row + 1):
            if str(file2_sheet.cells(row, 1).value) == value:
                # Append the complete row to File1.xlsm
                next_row = file1_sheet.range('A' + str(file1_sheet.cells.last_cell.row)).end('up').row + 1
                file2_sheet.range(file2_sheet.cells(row, 1), file2_sheet.cells(row, 8)).api.Copy()
                file1_sheet.cells(next_row, 1).api.PasteSpecial()
                break

# Save and close File1.xlsm
file1_workbook.save()
file1_workbook.close()

# Close File2.xlsx
file2_workbook.close()
