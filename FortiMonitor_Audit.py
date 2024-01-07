import pandas as pd
import requests
from pandas import json_normalize
import time

base_url = "https://api2.panopta.com/v2"
API_Key = input("Enter your Panopta API Key: ")
headers = {'Authorization': f'ApiKey {API_Key}'}
session = requests.Session()

def make_request(url, name):
    """
    Make a request to the Panopta API and fetch data.

    Args:
    - url (str): The API endpoint URL.
    - name (str): The name of the data being fetched.

    Returns:
    - data: The fetched data.
    """
    print(f"Fetching {name} data...")
    time.sleep(1)
    response = session.request("get", url=url, headers=headers)
    if response.status_code == 200:
        data = response.json().get(f"{name}_list")
        print(f"{name} data fetched successfully.")
        return data
    else:
        print(f"Failed to fetch {name} data.")
        return None

def reorder_columns(df, col_order):
    """
    Reorder DataFrame columns.

    Args:
    - df (pd.DataFrame): The DataFrame to reorder.
    - col_order (list): The desired order of columns.

    Returns:
    - pd.DataFrame: The reordered DataFrame.
    """
    return df[col_order + [col for col in df.columns if col not in col_order]]

def fetch_and_save_to_excel(data_func, sheet_name, file_writer):
    """
    Fetch data using the provided function, process and save it to an Excel file.

    Args:
    - data_func (function): The function to fetch data.
    - sheet_name (str): The name of the sheet in Excel.
    - file_writer: The ExcelWriter object for saving data.
    """
    data = data_func()
    if data:
        df = pd.json_normalize(data) if isinstance(data, list) else pd.DataFrame(data)
        if sheet_name == 'Server_data':
            df = flatten_attributes_column(df)
            # Define the desired order of columns
            col_order = ['name', 'fqdn', 'server_group', 'primary_monitoring_node']
            df = reorder_columns(df, col_order)
        df.to_excel(file_writer, sheet_name=sheet_name, index=False)
        print_separator()
        print(f"{sheet_name} data saved to Excel.")

def save_to_excel(df, sheet_name, file_writer):
    """
    Save DataFrame to an Excel sheet.

    Args:
    - df (pd.DataFrame): The DataFrame to save.
    - sheet_name (str): The name of the sheet in Excel.
    - file_writer: The ExcelWriter object for saving data.
    """
    df.to_excel(file_writer, sheet_name=sheet_name, index=False)
    print_separator()
    print(f"{sheet_name} data saved to Excel.")

def flatten_attributes_column(df):
    """
    Flatten the 'attributes' column with nested dictionaries.

    Args:
    - df (pd.DataFrame): The DataFrame to process.

    Returns:
    - pd.DataFrame: The DataFrame with flattened 'attributes'.
    """
    if 'attributes' in df.columns and isinstance(df['attributes'][0], dict):
        df_attributes = json_normalize(df['attributes'])
        df = pd.concat([df, df_attributes], axis=1)
        df = df.drop('attributes', axis=1)
    return df

def print_separator():
    """
    Print a separator line.
    """
    print("=" * 40)

def main():
    with pd.ExcelWriter('fmon_data.xlsx', engine='xlsxwriter') as writer:
        workbook = writer.book

        # Fetch and save each data to Excel
        fetch_and_save_to_excel(lambda: make_request(f'{base_url}/server', 'server'), 'Server_data', writer)
        fetch_and_save_to_excel(lambda: make_request(f'{base_url}/onsight', 'onsight'), 'Onsight_data', writer)
        fetch_and_save_to_excel(lambda: make_request(f'{base_url}/server_group', 'server_group'), 'Server_group_data', writer)
        fetch_and_save_to_excel(lambda: make_request(f'{base_url}/monitoring_node', 'monitoring_node'), 'Monitoring_node_data', writer)

        # Set Server_data sheet as the primary sheet
        if 'Server_data' in writer.sheets:
            worksheet_server_data = writer.sheets['Server_data']
            worksheet_server_data.set_first_sheet()

if __name__ == "__main__":
    main()
