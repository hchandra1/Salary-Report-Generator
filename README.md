# Salary-Report-Generator

This project automates the generation of visually appealing Excel reports with dynamic charts from salary data extracted from PDF files. It is designed to handle multiple tables within a single PDF, allowing for customizable column mapping to ensure flexibility and accuracy in data visualization.

## 🚀 Features
- **Dynamic PDF Table Extraction**: Supports multiple salary tables in a PDF, allowing the user to select and process specific data.
- **Customizable Column Mapping**: Provides an interactive way to map PDF table columns to desired fields like Month, Sales, Expenses, and Profit.
- **Excel Report Generation**: Produces a detailed financial report in Excel format with:
  - **Line Charts** for Sales, Expenses, and Profit trends.
  - Automatically formatted data for clear insights.
- **User-Friendly**: Minimal setup required; the script guides you through every step.

## 📂 Files
1. **`generate_salary_report.py`**  
   - The Python script to extract data from the PDF and generate an Excel report with trends.
2. **`salaries.pdf`**  
   - Example input file containing multiple salary tables.
3. **`salary_report_dashboard.xlsx`**  
   - Example output Excel file with dynamic charts.

## 🛠️ How It Works
1. The script scans the input PDF (`salaries.pdf`) for tables and extracts data using **pdfplumber**.
2. Allows the user to map relevant columns (e.g., Month, Sales, Expenses, Profit).
3. Converts the data into a structured format and creates an **Excel report** with line charts to visualize:
   - Sales trends over time.
   - Expense variations.
   - Profit analysis (if applicable).

## 📊 Output Example
The output Excel file contains:
- **Sheet 1**: Financial Data Table
- **Charts**:
  - **Sales Performance Over Time**
  - **Expenses Over Time**
  - **Profit Trends** (if mapped)

## 🔧 Requirements
- Python 3.7+
- Dependencies:
  - `pdfplumber`
  - `pandas`
  - `xlsxwriter`

## ⚙️ Setup
1. Clone this repository:
   ```bash
   git clone https://github.com/your-username/salary-report-generator.git
   cd salary-report-generator
