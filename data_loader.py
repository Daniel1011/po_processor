import pandas as pd
from config import excel_date_to_datetime # Assuming config.py is in the same directory

def load_and_prepare_data(input_file, ocd_col_name):
    """Loads data from Excel sheets and performs initial cleaning and preparation."""
    print("Reading input data sheets...")
    xls = pd.ExcelFile(input_file)

    stock_df = xls.parse('Stock')
    po_df = xls.parse('PO')
    first_lot_df = xls.parse('1ST LOT STATUS')
    capacity_status_df = xls.parse('Capacity Status')

    # Print debug information
    print("\nDebug: First few rows of first_lot_df before cleaning:")
    print(first_lot_df.head())
    print("\nDebug: Columns in first_lot_df:")
    print(first_lot_df.columns.tolist())

    # Data Cleaning and Preparation
    # Stock Data
    stock_df['Greige ETA'] = pd.to_datetime(stock_df['Greige ETA'], errors='coerce')
    stock_df['Greige Incoming'] = pd.to_numeric(stock_df['Greige Incoming'], errors='coerce')
    stock_df.dropna(subset=['DSM Code', 'Greige ETA', 'Greige Incoming'], inplace=True)
    stock_df = stock_df[stock_df['Greige Incoming'] > 0]

    # PO Data
    if 'Mã Vải' in po_df.columns and 'DSM Code' not in po_df.columns:
            po_df.rename(columns={'Mã Vải': 'DSM Code'}, inplace=True)
    elif 'Mã Vải' in po_df.columns and 'DSM Code' in po_df.columns:
        if 'DSM Code' in po_df.columns:
                po_df['DSM Code'] = po_df['DSM Code'].fillna(po_df['Mã Vải'])

    # Convert all date columns to datetime
    date_columns = ['CHD', ocd_col_name]
    for col in date_columns:
        if col in po_df.columns:
            po_df[col] = pd.to_datetime(po_df[col], errors='coerce')
        
    po_df['Quantity request'] = pd.to_numeric(po_df['Quantity request'], errors='coerce')
    po_df['Forecasted'] = po_df['Forecasted'].astype(str).str.lower()
    
    # Ensure essential PO columns are present (ocd_col_name handled above)
    base_required_po_cols = ['PO', 'DSM Code', 'CHD', 'Quantity request', 'Forecasted',
                             'SPL', 'FG name', 'Season', 'Local/ Export',
                             'CPT Name', 'ITEM', 'COLOR']
    # Add ocd_col_name to the list for checking if it wasn't missing initially
    required_po_cols_check = base_required_po_cols + [ocd_col_name]

    for col in required_po_cols_check: # Check all, including OCD
        if col not in po_df.columns:
            if col == ocd_col_name and ocd_col_name not in xls.parse('PO').columns : # Only warn if it was truly missing from file
                 pass # Already handled and warned about OCD_COL_NAME
            else:
                print(f"Warning: Column '{col}' not found in PO sheet. Adding it as NA.")
                po_df[col] = pd.NA

    po_df.dropna(subset=['PO', 'DSM Code', 'CHD', 'Quantity request'], inplace=True)

    # 1st Lot Status Data
    first_lot_df['DUE DATE'] = pd.to_datetime(first_lot_df['DUE DATE'], errors='coerce')
    first_lot_df['DSM Code'] = pd.to_numeric(first_lot_df['DSM Code'], errors='coerce')
    first_lot_df['STATUS'] = first_lot_df['STATUS'].str.strip().str.upper()
    first_lot_df.dropna(subset=['DSM Code', 'COLOR', 'STATUS'], inplace=True)

    # Print debug information after cleaning
    print("\nDebug: First few rows of first_lot_df after cleaning:")
    print(first_lot_df.head())
    print("\nDebug: Number of rows in first_lot_df after cleaning:", len(first_lot_df))

    # Capacity Status Data
    capacity_status_df['CAPACITY DATE'] = pd.to_datetime(capacity_status_df['CAPACITY DATE'], errors='coerce')
    capacity_status_df['CAPACITY REMAIN'] = pd.to_numeric(capacity_status_df['CAPACITY REMAIN'], errors='coerce')
    capacity_status_df.dropna(subset=['CAPACITY DATE', 'CAPACITY REMAIN'], inplace=True)
    capacity_status_df = capacity_status_df.sort_values(by='CAPACITY DATE')
    
    print("Input data read and preprocessed successfully.")
    return stock_df, po_df, first_lot_df, capacity_status_df

def read_input_data(file_path):
    """Reads and preprocesses input data from Excel file."""
    print("Reading input data sheets...")
    
    # Read all sheets
    stock_df = pd.read_excel(file_path, sheet_name='Stock')
    po_df = pd.read_excel(file_path, sheet_name='PO')
    first_lot_df = pd.read_excel(file_path, sheet_name='1ST LOT STATUS')
    
    # Print debug information
    print("\nDebug: First few rows of first_lot_df before cleaning:")
    print(first_lot_df.head())
    print("\nDebug: Columns in first_lot_df:")
    print(first_lot_df.columns.tolist())
    
    # Clean and preprocess data
    stock_df = clean_stock_data(stock_df)
    po_df = clean_po_data(po_df)
    first_lot_df = clean_first_lot_data(first_lot_df)
    
    # Print debug information after cleaning
    print("\nDebug: First few rows of first_lot_df after cleaning:")
    print(first_lot_df.head())
    
    print("Input data read and preprocessed successfully.")
    return stock_df, po_df, first_lot_df

def clean_stock_data(df):
    """Cleans and preprocesses stock data."""
    # Rename columns
    df = df.rename(columns={
        'Mã Vải': 'DSM Code',
        'ETA': 'Greige ETA',
        'Available': 'Greige Incoming'
    })
    
    # Convert DSM Code to integer
    df['DSM Code'] = pd.to_numeric(df['DSM Code'], errors='coerce')
    
    # Convert dates
    df['Greige ETA'] = pd.to_datetime(df['Greige ETA'], errors='coerce')
    
    # Convert quantities to numeric
    df['Greige Incoming'] = pd.to_numeric(df['Greige Incoming'], errors='coerce')
    
    return df

def clean_po_data(df):
    """Cleans and preprocesses PO data."""
    # Convert DSM Code to integer
    df['DSM Code'] = pd.to_numeric(df['DSM Code'], errors='coerce')
    
    # Convert dates
    date_columns = ['CHD', 'OCD( Order Creation Day)']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    # Convert quantities to numeric
    df['Quantity request'] = pd.to_numeric(df['Quantity request'], errors='coerce')
    
    return df

def clean_first_lot_data(df):
    """Cleans and preprocesses 1ST LOT STATUS data."""
    # Print debug information before cleaning
    print("\nDebug: First few rows of first_lot_df before cleaning in clean_first_lot_data:")
    print(df.head())
    
    # Convert DSM Code to integer
    df['DSM Code'] = pd.to_numeric(df['DSM Code'], errors='coerce')
    
    # Convert dates
    df['DUE DATE'] = pd.to_datetime(df['DUE DATE'], errors='coerce')
    
    # Clean STATUS values
    df['STATUS'] = df['STATUS'].str.strip().str.upper()
    
    # Print debug information after cleaning
    print("\nDebug: First few rows of first_lot_df after cleaning in clean_first_lot_data:")
    print(df.head())
    
    return df 