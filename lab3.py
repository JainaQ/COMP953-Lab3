from sys import argv, exit
import os
from datetime import date
import pandas as pd
import re

def main():
    sales_csv = get_sales_csv()
    orders_dir = create_orders_dir(sales_csv)
    process_sales_data(sales_csv, orders_dir)
    return
# Get path of sales data CSV file from the command line
def get_sales_csv():
    # Check whether command line parameter provided
    num_params = len(argv) - 1
    if num_params >= 1:
        sales_csv = argv[1]
        # Check whether provide parameter is valid path of file
        if os.path.isfile(sales_csv):
            return sales_csv
        else:
            print('Error: Invalid CSV File Path')
            exit(1) 
    else:
        print('Error: Please Define CSV File Path')
        exit(1)

    return
# Create the directory to hold the individual order Excel sheets
def create_orders_dir(sales_csv):
    # Get directory in which sales data CSV file resides
    sales_dir = os.path.dirname(os.path.abspath(sales_csv))

    # Determine the name and path of the directory to hold the order data files
    todays_date = date.today().isoformat()
    orders_dir  = os.path.join(sales_dir, f'Orders_{todays_date}')
    # Create the order directory if it does not already exist
    if not os.path.isdir(orders_dir):
        os.makedirs(orders_dir)

    return orders_dir

# Split the sales data into individual orders and save to Excel sheets
def process_sales_data(sales_csv, orders_dir):
    # Import the sales data from the CSV file into a DataFrame
    sales_df = pd.read_csv(sales_csv)

    # Insert a new "TOTAL PRICE" column into the DataFrame
    sales_df.insert(7, 'TOTAL PRICE', sales_df['ITEM QUANTITY'] * sales_df['ITEM PRICE'])

    # Remove columns from the DataFrame that are not needed
    sales_df.drop(columns=['ADDRESS', 'CITY', 'STATE', 'POSTAL CODE', 'COUNTRY'], inplace=True)

    # Group the rows in the DataFrame by order ID
    for order_id, order_df in sales_df.groupby('ORDER ID'):

        # Remove the "ORDER ID" column
        sales_df.drop(columns=['ORDER ID'], inplace=True)

        # Sort the items by item number
        order_df.sort_values(by='ITEM NUMBER', inplace=True)

        # Append a "GRAND TOTAL" row
        grand_total = order_df["TOTAL PRICE"].sum()
        grand_total_df = pd.DataFrame({'ITEM PRICE': ['GRAND TOTAL'], 'TOTAL PRICE': [grand_total]})
        order_df = pd.concat([order_df, grand_total_df])

        export_order_to_excel(order_id, order_df, orders_dir)

        break
# Determine the file name and full path of the Excel sheet
def export_order_to_excel(order_id, order_df, orders_dir):

    customer_name = order_df['CUSTOMER NAME'].values[0]
    customer_name = re.sub(r'\W', '', customer_name)
    order_file = f'Order{order_id}_{customer_name}.xlsx'
    order_path = os.path.join(orders_dir, order_file)

# Export the data to an Excel sheet
    sheet_name = f'Order #{order_id}'
    order_df.to_excel(order_path, index=False, sheet_name=sheet_name)

# TODO: Format the Excel sheet

    return

if __name__ == '__main__':
    main()