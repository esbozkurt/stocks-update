README.md

Update Stock and Cargo Transactions in Excel

Overview

This script automates the management of stock and cargo transaction data stored in an Excel file. It processes the data across multiple sheets, updates stock quantities based on transactions, and appends transaction details to a history log. The output is a new Excel file with the updated information.

Features

Reads and processes data from an Excel workbook containing Stock, Cargo Transactions, and History sheets.

Updates stock quantities in the Stock sheet based on the Cargo Transactions sheet.

Logs transaction details in the History sheet for record-keeping.

Saves all updates to a new Excel workbook.

Requirements

Python 3.7 or higher

Libraries: pandas, openpyxl, xlsxwriter