# Invoicing Application

This is a simple GUI-based invoicing application written in Python. It allows you to create invoices, calculate total income, and generate Word documents for the invoices.

## Features

- Create new invoices with customer information and item details
- Calculate the total income of the company
- Generate Word documents for the created invoices

## Installation

1. Install the required dependencies:
   - pandas: `pip install pandas`
   - tkinter: Included with Python (no separate installation required)
   - sqlite3: Included with Python (no separate installation required)
   - docxtpl: `pip install docxtpl`
   - openpyxl: `pip install openpyxl`
   -flask: 'pip install flask'
2. Download the source code or clone the repository.

3. Run the `invoicing_application.py` script.

## Usage

1. Launch the application by running the `invoicing_application.py` script.

2. Fill in the customer information, item details, and unit prices in the provided fields.

3. Click the "Add Item" button to add the item to the invoice.

4. Repeat steps 2-3 to add more items to the invoice.

5. Click the "Create Invoice" button to save the invoice data in a SQLite database and generate a Word document for the invoice.

6. Use the "New Invoice" button to start a new invoice and clear the form fields.

7. Click the "Calculate Total Income" button to calculate the total income of the company and export the income data and chart to an Excel file.

## File Structure

- `invoicing_application.py`: The main Python script that contains the code for the GUI-based application.
- `invoice_template.docx`: The Word document template for generating the invoices.
- `invoices.db`: SQLite database file for storing the invoice data.
- `README.md`: This file, providing an overview and instructions for the application.



## Acknowledgments

- invoice generator: https://github.com/codefirstio/invoice-generator-tkinter-and-doxtpl/blob/main/main.py
- [pandas](https://pandas.pydata.org/) - Python data analysis library
- [tkinter](https://docs.python.org/3/library/tkinter.html) - Python GUI toolkit
- [sqlite3](https://docs.python.org/3/library/sqlite3.html) - Python SQLite database module
- [docxtpl](https://docxtpl.readthedocs.io/) - Python library for generating Word documents from templates
- [openpyxl](https://openpyxl.readthedocs.io/) - Python library for working with Excel files
