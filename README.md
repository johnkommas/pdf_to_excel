# PDF to Excel Converter

This application converts invoice data from a PDF file to an Excel spreadsheet.

## Features

- Converts data from PDF tables to an Excel spreadsheet using pandas and tabula-py.
- Retrieves supplier's unique number from given areas in the PDF
- Adjusts column widths in Excel for better readability

## Dependencies

The application uses these libraries:

- os
- subprocess
- sys
- logging
- pandas
- tabula
- PyPDF2
- openpyxl

## Usage

1. Set the directory for the Invoice file in `FOLDER`.

2. Run `python app.py`.

Please ensure all dependencies are installed and that the Python environment is set up correctly.

## Future Scope

In its current state, this application is designed to work with a specific invoice format. In the future, it will have additional features to handle more formats, generating more comprehensive reports, and improving error handling.

## Contributing

The mission of this project is to provide a simple, efficient way to convert invoice data from PDF to Excel. Contributions towards improving this application are much appreciated.

## License

2024 Ioannis E. Kommas. All Rights Reserved.