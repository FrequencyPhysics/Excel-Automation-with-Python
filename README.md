# Excel-Automation-with-Python
This project demonstrates how Python can be used to automate tasks in Excel spreadsheets.
It reads data, applies a price correction, creates a bar chart, and saves the results into a new file.

## Features
- Reads an Excel file (.xlsx) automatically.
- Prints the total number of rows in the sheet.
- Applies a 10% discount to prices in column 3.
- Writes corrected prices into column 4.
- Generates a bar chart of the corrected prices.
- Saves results into a new Excel workbook.

## Project Structure  
excel-automation/
│── data/ # sample spreadsheets (input/output)
│── main.py # main script
│── requirements.txt # dependencies (openpyxl)
│── README.md # project description

## Installation  
1. Clone the repository:  
   ```bash
   git clone https://github.com/<your-username>/excel-automation.git
   cd excel-automation
2. Install dependencies:
   ```bash
   pip install -r requirements.txt 

## Usage
  **python main.py

# Walkthrough:

**Reading Correct File and Counting Rows:
<img width="580" height="135" alt="image" src="https://github.com/user-attachments/assets/94289f02-152e-4fd3-a0ea-06f6b5c84820" />

**Input file - Reading Prices in column 3:
<img width="222" height="70" alt="image" src="https://github.com/user-attachments/assets/2fab484d-e889-4cd4-b58d-e6d693a9c95c" />

** Adding Corrected Prices & new workbook - Output file:
<img width="717" height="257" alt="image" src="https://github.com/user-attachments/assets/60b976fb-4df2-4f77-818e-0e00d9edb7dd" />

