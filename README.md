# Report Generator Web App

A Streamlit web application that generates multiple Word documents by merging Excel data into a Word template.

## Overview

This application provides a user-friendly web interface for generating reports from Excel data and Word templates. It's built with Streamlit, making it accessible through a web browser without requiring technical knowledge to operate.

## Features

- **Web-based Interface**: Easy-to-use Streamlit interface accessible from any browser
- **Interactive File Upload**: Drag-and-drop Excel data files and Word templates
- **Batch Processing**: Generate multiple reports at once from Excel data rows
- **Progress Tracking**: Keeps track of processed rows in the Excel file
- **Download Options**: 
  - Download individual reports
  - Download all reports as a ZIP file
  - Download updated Excel file with processing status
- **Smart Processing**: Only processes rows that haven't been processed before
- **Template System**: Uses placeholders in Word templates that get replaced with Excel data
- **Support for Complex Documents**: Handles placeholders in both paragraphs and tables

## Requirements

- Python 3.6+
- Dependencies (installed automatically):
  - streamlit>=1.0.0
  - pandas>=1.5.0
  - python-docx>=0.8.11
  - openpyxl>=3.0.10
  - docx2txt>=0.8

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/jplives4surf/report-gen-web.git
   cd report-gen-web
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Start the Streamlit app:
   ```bash
   streamlit run streamlit_app.py
   ```

2. Open your web browser and navigate to the URL shown in the terminal (typically http://localhost:8501)

3. Use the web interface to:
   - Upload your Excel data file (.xlsx)
   - Upload your Word template file (.docx)
   - Click "Generate Reports" to process the data
   - Download individual reports, a ZIP of all reports, or the updated Excel file

## Template Format

In your Word template, use double curly braces for placeholders that will be replaced with data from your Excel file:

- Example: `{{first_name}}` will be replaced with the value from the "first_name" column
- Placeholders are case-insensitive
- Placeholders work in both regular paragraphs and table cells

## How It Works

1. **Data Processing**: The app reads your Excel file and identifies rows to process (those without entries in the 'processed' column)
2. **Template Merging**: For each row, the app creates a copy of your Word template and replaces all placeholders with data from that row
3. **Output Generation**: The app generates individual Word documents for each processed row
4. **Tracking**: The app updates the 'processed' column in your Excel file with the filename of the generated report
5. **Download Options**: You can download individual reports, all reports as a ZIP, or the updated Excel file

## Notes

- If your Excel file doesn't have a 'processed' column, one will be added automatically
- Rows that have already been processed (have a value in the 'processed' column) will be skipped
- The app maintains the original Excel data structure while adding/updating the 'processed' column
