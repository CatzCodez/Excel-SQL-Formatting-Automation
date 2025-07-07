# Excel-SQL Formatting Automation
This Python script processes raw Excel or CSV files into clean, SQL-importable Excel tables. It:
- Converts column and row names to `snake_case`
- Removes extra headers, blank rows, and noise
- Converts numbers with commas into integers
- Outputs formatted `.xlsx` files using Excel table formatting
- Saves results to the `cleaned_reports/` folder with `_formatted` added to the filename

## Folder Structure
Excel_Automation/

├── raw_reports/ --> Place raw .csv or .xlsx files here

├── cleaned_reports/ --> Processed output files go here

├── table_format.py --> The script that performs formatting

## How to Use
1. Clone the repository:
2. Place your `.csv` or `.xlsx` files into the `raw_reports` folder.
3. Run the script:
4. Your formatted Excel files will be saved into the `cleaned_reports` folder with `_formatted.xlsx` added to their names.

## Requirements
- Python 3.x
- pandas
- openpyxl

Install required packages using (in terminal or command prompt):
pip install pandas openpyxl

## Example
Input file: `raw_data_1.csv`  
Output file: `cleaned_reports/raw_data_1_formatted.xlsx`

This output will include:
- Snake_case formatting for all headers and row names
- Numeric values cleaned of commas
- Excel Table format applied for SQL compatibility

## Use Cases
- Preparing public datasets for SQL import
- Cleaning messy Excel/CSV reports
- Automating spreadsheet standardization for data pipelines
