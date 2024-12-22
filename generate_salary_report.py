import pdfplumber
import pandas as pd
import xlsxwriter

# Function to dynamically extract tables from the PDF
def extract_tables_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        all_tables = []
        for page in pdf.pages:
            tables = page.extract_tables()
            for table in tables:
                if table:
                    # Convert table to pandas DataFrame
                    df = pd.DataFrame(table[1:], columns=table[0])  # Exclude header row from table
                    all_tables.append(df)
        return all_tables

# Function to let user map columns for sales and financial data
def map_columns(df):
    print("Table detected with the following columns:")
    print(df.columns)

    # Ask the user to map columns
    month_col = input("Enter the name of the column representing 'Month': ")
    sales_col = input("Enter the name of the column representing 'Sales': ")
    expenses_col = input("Enter the name of the column representing 'Expenses': ")
    profit_col = input("Enter the name of the column representing 'Profit' (optional): ")

    selected_columns = [month_col, sales_col, expenses_col]
    if profit_col:
        selected_columns.append(profit_col)

    # Return mapped DataFrame
    return df[selected_columns]

# Step 2: Generate Excel Report from extracted data
def generate_excel_report(financial_df, file_name):
    with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
        # Write the financial data into the first sheet
        financial_df.to_excel(writer, sheet_name='Financial Data', index=False)

        # Access the XlsxWriter workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Financial Data']

        # Create charts for Sales, Expenses, and Profit (if available)
        categories_range = f"'Financial Data'!$A$2:$A${len(financial_df)+1}"
        
        # Sales chart
        chart_sales = workbook.add_chart({'type': 'line'})
        chart_sales.add_series({
            'name': 'Sales',
            'categories': categories_range,
            'values': f"'Financial Data'!$B$2:$B${len(financial_df)+1}",
        })
        chart_sales.set_title({'name': 'Sales Performance Over Time'})
        chart_sales.set_x_axis({'name': 'Month'})
        chart_sales.set_y_axis({'name': 'Sales (in $)'})
        chart_sales.set_size({'width': 720, 'height': 480})
        worksheet.insert_chart('E2', chart_sales)

        # Expenses chart
        chart_expenses = workbook.add_chart({'type': 'line'})
        chart_expenses.add_series({
            'name': 'Expenses',
            'categories': categories_range,
            'values': f"'Financial Data'!$C$2:$C${len(financial_df)+1}",
        })
        chart_expenses.set_title({'name': 'Expenses Over Time'})
        chart_expenses.set_x_axis({'name': 'Month'})
        chart_expenses.set_y_axis({'name': 'Expenses (in $)'})
        chart_expenses.set_size({'width': 720, 'height': 480})
        worksheet.insert_chart('E20', chart_expenses)  # Adjust the location to avoid overlapping

        # Profit chart (optional)
        if 'Profit' in financial_df.columns:
            chart_profit = workbook.add_chart({'type': 'line'})
            chart_profit.add_series({
                'name': 'Profit',
                'categories': categories_range,
                'values': f"'Financial Data'!$D$2:$D${len(financial_df)+1}",
            })
            chart_profit.set_title({'name': 'Profit Over Time'})
            chart_profit.set_x_axis({'name': 'Month'})
            chart_profit.set_y_axis({'name': 'Profit (in $)'})
            chart_profit.set_size({'width': 720, 'height': 480})
            worksheet.insert_chart('E40', chart_profit)

    print(f"Financial report with sales and trends generated: {file_name}")

# Step 3: Main Function to Extract from PDF and Generate Report
def main(pdf_file):
    # Extract tables from the PDF
    extracted_tables = extract_tables_from_pdf(pdf_file)
    
    if extracted_tables:
        # Let the user choose the table they want to use (if multiple tables found)
        if len(extracted_tables) > 1:
            print(f"{len(extracted_tables)} tables detected. Please select the table you want to use:")
            for i, table in enumerate(extracted_tables):
                print(f"\nTable {i+1}:")
                print(table.head())  # Show first few rows to help user decide

            table_index = int(input("Enter the number of the table you want to use: ")) - 1
            financial_data = extracted_tables[table_index]
        else:
            financial_data = extracted_tables[0]

        # Map the columns based on user input
        mapped_financial_data = map_columns(financial_data)

        # Convert the financial columns to numeric
        for col in mapped_financial_data.columns[1:]:
            mapped_financial_data[col] = pd.to_numeric(mapped_financial_data[col], errors='coerce')

        # Generate the Excel report with charts for financial trends
        excel_file = 'financial_trends_report.xlsx'
        generate_excel_report(mapped_financial_data, excel_file)
    else:
        print("No tables found in the PDF.")

# Run the script
if __name__ == "__main__":
    pdf_file_path = 'financial_report.pdf'  # Path to your PDF file
    main(pdf_file_path)
