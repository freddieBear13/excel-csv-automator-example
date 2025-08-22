import pandas as pd
import configparser
import argparse
import os

def read_source_file(filepath, separator):
    """
        Reads source file (.csv or .xlsx) and returns DataFrame
    """
    _, file_extension = os.path.splitext(filepath)
    if file_extension == '.csv':
        return pd.read_csv(filepath, sep=separator)
    elif file_extension in ['.xlsx', 'xls']:
        return pd.read_excel(filepath)
    else:
        raise ValueError(f"Unsupported file format: {file_extension}")
            
def write_excel_report(output_path, manager_sales, top_products):
    """
        Writes final data in Excel file with chart
    """
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            manager_sales.to_excel(writer, sheet_name='Manager report')
            top_products.to_excel(writer, sheet_name='Top-5 products')

            workbook = writer.book
            worksheet_managers = writer.sheets['Manager report']
            worksheet_products = writer.sheets['Top-5 products']

            worksheet_managers.set_column('A:B', 25)

            chart = workbook.add_chart({'type': 'column'})
            num_rows = len(manager_sales)
            chart.add_series({
                'name': 'Sum of sales',
                'categories': ['Manager report', 1, 0, num_rows, 0],
                'values': ['Manager report', 1, 1, num_rows, 1],
            })

            chart.set_title({'name': 'Managers common'})
            chart.set_legend({'position': 'none'})
            worksheet_managers.insert_chart('D2', chart)

            worksheet_products.set_column('A:A', 30)
            worksheet_products.set_column('B:B', 20)

def process_data_and_create_report(config, input_file, output_file, start_date=None, end_date=None):
    """
    Main process function
    """
    try:
        separator = config['SETTINGS']['csv_separator']
        col_date = config['COLUMNS']['date']
        col_qty = config['COLUMNS']['quantity']
        col_price = config['COLUMNS']['price']
        col_manager = config['COLUMNS']['manager']
        col_product = config['COLUMNS']['product_name']
        col_sum = config['COLUMNS']['total_sum']
        
        df = read_source_file(input_file, separator)

        required_columns = [col_date, col_qty, col_price, col_manager, col_product] 
        for col in required_columns:
            if col not in df.columns:
                raise ValueError(f"Needed column is absent in the file: '{col}'")

        df[col_date] = pd.to_datetime(df[col_date]).dt.normalize()
        if start_date:
            df = df[df[col_date] >= pd.to_datetime(start_date)]
        if end_date:
            df = df[df[col_date] <= pd.to_datetime(end_date)]

        if df.empty:
            print("No data for the report")
            return 
            
        df[col_sum] = df[col_qty] * df[col_price]
        manager_sales = df.groupby(col_manager)[col_sum].sum().sort_values(ascending=False)
        top_products = df.groupby(col_product)[col_sum].sum().sort_values(ascending=False).head(5)
        
        write_excel_report(output_file, manager_sales, top_products)

        print(f"Success!")
    except (FileNotFoundError, ValueError, configparser.Error, KeyError) as e:
        print(f"ERROR: {e}")
    except Exception as e:
        print(f"Unexepted error: {e}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Sales reports automator")
    parser.add_argument("input_file", help="Path to CSV or XLSX file")
    parser.add_argument("--start", help="Start date (YYYY-MM-DD)")
    parser.add_argument("--end", help="End date (YYYY-MM-DD)")
    args = parser.parse_args()

    try:
        config = configparser.ConfigParser()
        config.read('config.ini')
        if not config.sections():
            raise configparser.Error("Can't find or read config.ini file")
        
        file_name, _ = os.path.splitext(args.input_file)
        report_file = f"{file_name}_report.xlsx"

        process_data_and_create_report(config, args.input_file, report_file, args.start, args.end)
    except configparser.Error as e:
        print(f'CONFIG ERROR: {e}')
