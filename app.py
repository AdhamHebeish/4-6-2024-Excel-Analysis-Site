from flask import Flask, render_template, request, jsonify
import openpyxl
import magic
import io  # Used for handling file-like objects

# Improved error handling with custom exception classes
class InvalidFileTypeError(Exception):
    pass

class FileValidationError(Exception):
    def __init__(self, errors):
        self.errors = errors

app = Flask(__name__, template_folder='templates')  # Specify template folder

@app.route('/')
def index():
    return render_template('index.html')  # Ensure correct file name

@app.route('/analyze_excel', methods=['POST'])
def analyze_excel():
    if 'file' not in request.files:
      return jsonify({'error': 'No file uploaded'}), 400
   
    file = request.files['file']

    # Check file type using filename extension (optional)
    if not file.filename.lower().endswith(('.xls', '.xlsx')):
        return jsonify({'error': 'Only Excel files (.xls or .xlsx) are allowed'}), 400

    # Convert uploaded file to a file-like object
    file_data = io.BytesIO(file.read())
    file_data.seek(0)  # Reset the pointer to the beginning

    try:
      workbook = openpyxl.load_workbook(file_data, data_only=True)
      worksheet = workbook.active
    except Exception as e:  # Catch any exception related to file opening
        return jsonify({'error': f'An error occurred while processing the file: {str(e)}'}), 400

    # Assume the header row exists on the first row (adjust as needed)
    header_row = [cell.value.strip() for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]


    # Function to validate a single row
    def validate_row(row):
        errors = []
        # Field validations based on your requirements
        if not row[header_row.index('First Name')].value.strip():
         errors.append({'field': 'First Name', 'message': 'Missing mandatory field'})
        else:
          # Check data type before stripping
         if isinstance(row[header_row.index('First Name')].value, str):
            # Only call strip() on strings
            if not row[header_row.index('First Name')].value.strip():
                errors.append({'field': 'First Name', 'message': 'Missing mandatory field'})
         else:
            # Handle case where expected value is a string but cell contains a number
            errors.append({'field': 'First Name', 'message': 'Invalid format: Must be text'})

        if not row[header_row.index('Country Code')].value:  # Check for empty value (including whitespace)
         errors.append({'field': 'Country Code', 'message': 'Missing mandatory field'})
        else:
         # You can add validation for Country Code format here (e.g., length)
         country_code = str(row[header_row.index('Country Code')].value)  # Convert to string
         if len(country_code) != 2:
          errors.append({'field': 'Country Code', 'message': 'Invalid format: Must be 2 characters'})

          mobile_number_index = header_row.index('Mobile Number')

          # Check for empty value
          # Check for empty value and handle None
          if row[mobile_number_index].value is None:
            errors.append({'field': 'Mobile Number', 'message': 'Missing mandatory field'})
          else:
            # Convert to string and perform validation (as before)
            mobile_number = str(row[mobile_number_index].value)
            if not isinstance(mobile_number, str):
              errors.append({'field': 'Mobile Number', 'message': 'Invalid format: Must be text'})

          return {'data': row, 'errors': errors}

        if not row[header_row.index('Address')].value:  # Check for empty value (including whitespace)
         errors.append({'field': 'Address', 'message': 'Missing mandatory field'})
        # No further validation needed for address (assuming any format is allowed)

        # Last Name and Notes are optional, so no validation required here

         if errors:
            raise FileValidationError(errors)

        return {
        'data': {
            'First Name': row[header_row.index('First Name')].value.strip() if isinstance(row[header_row.index('First Name')].value, str) else None,
            'Last Name': row[header_row.index('Last Name')].value.strip() if isinstance(row[header_row.index('Last Name')].value, str) else None,
            'Country Code': str(row[header_row.index('Country Code')].value) if row[header_row.index('Country Code')].value else None,
            'Mobile Number': row[header_row.index('Mobile Number')].value.strip() if isinstance(row[header_row.index('Mobile Number')].value, str) else None,
            'Address': row[header_row.index('Address')].value,
            'Notes': row[header_row.index('Notes')].value.strip() if isinstance(row[header_row.index('Notes')].value, str) else None,
            # Include other relevant fields
        },
        'valid': not errors,
        'errors': errors if errors else None
    }

    # Analyze the spreadsheet data
    try:
     data = {}  # Dictionary to store row data with index
     for row_index, row in enumerate(worksheet.iter_rows(min_row=2)):  # Skip header row
      processed_data = validate_row(row)
      data[row_index] = processed_data['data']  # Store data with index
     return jsonify(data)
    except Exception as e:
     return jsonify({'error': f'An error occurred during processing: {str(e)}'}), 400

if __name__ == '__main__':
    app.run(debug=True)

