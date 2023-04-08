from flask import Flask, request, send_file
import openpyxl
from openpyxl.styles import Alignment, Border, Side, PatternFill
import io

app = Flask(__name__)

@app.route('/')
def index():
    return '''
        <form action="/upload" method="post" enctype="multipart/form-data">
            <input type="file" name="file">
            <input type="submit" value="Upload">
        </form>
        <br>
        <a href="/download">Download Processed File</a>
    '''

@app.route('/upload', methods=['POST'])
def upload():
    file = request.files['file']
    # Load the workbook from the uploaded file
    wb = openpyxl.load_workbook(file)
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
            new_sheet.cell(row=new_row, column=new_col).border = thin_border
            new_sheet.cell(row=new_row, column=new_col).alignment = Alignment(horizontal='center')
            new_sheet.cell(row=new_row, column=new_col).value = sheet.cell(row=row, column=1).value

            new_sheet.cell(row=new_row, column=new_col+1).border = thin_border
            new_sheet.cell(row=new_row, column=new_col+1).alignment = Alignment(horizontal='center')
            new_sheet.cell(row=new_row, column=new_col+1).value = ', '.join(violation_types)

            new_sheet.cell(row=new_row, column=new_col+2).border = thin_border
            new_sheet.cell(row=new_row, column=new_col+2).alignment = Alignment(horizontal='center')
            new_sheet.cell(row=new_row, column=new_col+2).value = len(violation_types)

    # Save the modified workbook to a temporary file-like object in memory
    output_file = io.BytesIO()
    wb.save(output_file)
    output_file.seek(0)

    # Return the modified workbook as a downloadable file
    return send_file(output_file,
                     attachment_filename='modified.xlsx',
                     as_attachment=True)

@app.route('/download')
def download():
    # Return the processed file as a downloadable file
    return send_file('path/to/processed/file.xlsx',
                     attachment_filename='processed.xlsx',
                     as_attachment=True)

if __name__ == '__main__':
    import os
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
