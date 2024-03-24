Python script used for running Excel an report from .xlsx file(s)
<h1> Starting Out: </h1>
1. Download openpyxl via pip: https://pypi.org/project/openpyxl/<p></p>
2. Open Python interpretor and type <code>pip install openpyxl</code><p></p>
3. The sample code I've provided is simply what I was doing to move columns around and get alignments down based on the report(s) I needed to run. I highly recommend spending some time reading the documentation of Openpyxl to accodimate your specific needs at the following: https://openpyxl.readthedocs.io/en/stable/ <p></p>
4. The rest will be up to the user to decide, but the easiest way to run this script is by inserting it into an IDE like Visual Studio Code or PyCharm.


# Load the workbook
<code>file_path = 'C:\\Users\\USER\\...\\report_1.xlsx'
wb = openpyxl.load_workbook(file_path)</code>

# Access the first sheet in the workbook
<code>sheet = wb.active</code>

# Rename the sheet to what you want
<code>sheet.title = 'report_1'</code>

<h1> Manipulation(s): </h1>

#Cut Column Q and insert it in front of Column A:

<code>col_q = [cell.value for cell in sheet['Q']]
sheet.insert_cols(1)  # Insert a new column at position 1 (A)
for idx, value in enumerate(col_q, start=1):
sheet.cell(row=idx, column=1, value=value)
sheet.delete_cols(18)  # Delete the original column Q (now at position 18)
</code>

# Wrap text for all cells except for the first row
<code>for row in sheet.iter_rows(min_row=2):
for cell in row:
cell.alignment = Alignment(wrap_text=True)</code>

# Center align columns excluding the first row
<code>for row in sheet.iter_rows(min_row=2):
for cell in row:
if cell.column_letter in ['D', 'E', 'F', 'I', 'M', 'N', 'J']:
cell.alignment = Alignment(horizontal='center')</code>

# Add commas for columns G and H excluding the first row
<code>for row in sheet.iter_rows(min_row=2):
for cell in row:
if cell.column_letter in ['G', 'H', 'I']:
cell.number_format = '#,##0'</code>

# Format columns J, K, and L as currency excluding the first row
<code>for row in sheet.iter_rows(min_row=2):
for cell in row:
if cell.column_letter in ['K', 'L', 'M']:
cell.number_format = '"$"#,##0.00'</code>

# Save the modified workbook
<code>wb.save('C:\\Users\\USER\\...\\report_1.xlsx')</code>

</code>
