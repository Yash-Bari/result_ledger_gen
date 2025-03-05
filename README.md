# Result Ledger Parser

A dynamic web application built with Streamlit that parses academic result ledgers from PDF files and converts them into organized Excel spreadsheets.

## Features

- PDF Result Ledger Parsing
- Dynamic Subject Detection
- Customizable Subject Selection
- Excel Export with Organized Data
- Support for Various Result Formats
- Real-time Preview
- Error Handling

## Requirements

```
streamlit
pandas
PyPDF2
xlsxwriter
```

## Installation

1. Clone this repository
2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```

2. Upload your PDF result ledger file
3. The app will automatically detect subjects from the PDF
4. Select the subjects you want to include in the Excel file
5. Click "Process File" to generate the Excel spreadsheet
6. Download the generated Excel file

## Input Format

The PDF file should contain result ledger data formatted as follows:
- Header lines with student details (SEAT NO., NAME, etc.)
- Course lines containing marks with the following structure:
  - Course code at the start
  - Course name
  - Asterisk (*) followed by marks
  - Various mark components (Insem, ESE, Total, TW, PR)

## Data Processing

The application processes:
- Student basic information (Seat No., Name)
- Subject-wise marks (Internal, External, Total)
- Laboratory/Practical marks
- MOOC course marks
- SGPA and CGPA calculations
- Total Credits earned

## Project Structure

- `app.py`: Main Streamlit application and UI logic
- `result_backend.py`: Core parsing and processing functions
- `requirements.txt`: Project dependencies

## Error Handling

The application includes robust error handling for:
- Invalid PDF formats
- Missing or malformed data
- File processing errors
- Invalid mark formats

## Contributing

Feel free to submit issues and enhancement requests.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
