from flask import Flask, request, send_file
import openpyxl
from openpyxl.styles import Alignment, Border, Side, PatternFill
import os

app = Flask(__name__)

@app.route('/')
def index():
    return '''
        <html>
            <body>
                <form action="/upload" method="post" enctype="multipart/form-data">
                    <input type="file" name="file">
                    <input type="submit" value="Upload">
                </form>
            </body>
        </html>
    '''

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    filename = file.filename
    file.save(filename)

    # Open the Excel file and select the appropriate sheet
    wb = openpyxl.load_workbook(filename)
    sheet = wb['Sheet1']

    # Create a new sheet to hold the extracted data
    new_sheet = wb.create_sheet('Extracted Data')

    # Initialize the row and column counters for the new sheet
    new_row = 1
    new_col = 1

    # Define border style
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    # Define header fill
    header_fill = PatternFill(start_color='B8CCE4', end_color='B8CCE4', fill_type='solid')

    # Write the headers for the new sheet and apply formatting
    new_sheet.cell(row=new_row, column=new_col, value='Name').fill = header_fill
    new_sheet.cell(row=new_row, column=new_col+1, value='Violations').fill = header_fill
    new_sheet.cell(row=new_row, column=new_col+2, value='Violations Count').fill = header_fill

    # Loop through each row in the original sheet
    for row in range(2, 101):
        # Check if any of the violation columns have a value greater than 0
        has_violation = False
        violation_types = []
        for col in range(3, 11):
            cell_value = sheet.cell(row=row, column=col).value
            if cell_value is not None and cell_value > 0:
                has_violation = True
                violation_types.append(sheet.cell(row=1, column=col).value)

        # If a violation was found, write the name and violation types to the new sheet
        if has_violation:
            new_row += 1
            new_sheet.cell(row=new_row, column=new_col, value=sheet.cell(row=row, column=1).value).border = thin_border
            new_sheet.cell(row=new_row, column=new_col+1, value=', '.join(violation_types)).border = thin_border

        # Sum the violation values and write the total to the new sheet
        violation_sum = sum(sheet.cell(row=row, column=col).value for col in range(3, 11) if sheet.cell(row=row, column=col).value is not None)
        if violation_sum > 0:
            new_sheet.cell(row=new_row, column=new_col+2, value=violation_sum).border = thin_border

    # Apply formatting to all cells in new sheet
    for col in new_sheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        new_sheet.column_dimensions[column].width = adjusted_width
        for cell in col:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = thin_border

    # Save the updated Excel file
    wb.save(filename)

    return '''
        <html>
            <body>
                <a href="/download/{}">Download</a>
            </body>
        </html>
    '''.format(filename)

@app.route('/download/<filename>')
def download(filename):
    return send_file(filename)

if __name__ == '__main__':
    app.run()
