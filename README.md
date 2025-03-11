This is a Python CLI program designed to streamline the process of encoding monthly sales reports into Excel. It provides an efficient way to record, edit, and manage sales transactions with built-in error handling and structured data organization.

Features:
Create a New Report
    Input item details (code, quantity, price) with real-time validation.
    Supports correction of incorrectly entered data.
    Saves data into an Excel file with structured formatting.

Edit Existing Reports
    Lists all available Excel sales reports for easy selection.
    Allows modification of past reports with new entries.
    Ensures data integrity by appending and updating values.

Automated Excel Handling
    Uses openpyxl to create, edit, and update Excel files.
    Categorizes sales by item type and brand for structured analysis.
    Computes total daily sales automatically.


Installation:
    Ensure Python 3.x is installed
        Download and install Python from python.org.

Install dependencies:
pip install openpyxl

Clone the repository:
git clone https://github.com/God-Fr3y/MonthlySalesReport.git
cd MonthlySalesReport

Run the script:
python main.py
