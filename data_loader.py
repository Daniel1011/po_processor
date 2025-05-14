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
    stock_df.dropna(subset=['Greige Code', 'Greige ETA', 'Greige Incoming'], inplace=True)
    stock_df = stock_df[stock_df['Greige Incoming'] > 0]

    # PO Data
    if 'Mã Vải' in po_df.columns and 'Greige Code' not in po_df.columns:
        po_df.rename(columns={'Mã Vải': 'Greige Code'}, inplace=True)
    elif 'Mã Vải' in po_df.columns and 'Greige Code' in po_df.columns:
        if 'Greige Code' in po_df.columns:
            po_df['Greige Code'] = po_df['Greige Code'].fillna(po_df['Mã Vải'])

    # Convert all date columns to datetime
    date_columns = ['CHD', ocd_col_name]
    for col in date_columns:
        if col in po_df.columns:
            po_df[col] = pd.to_datetime(po_df[col], errors='coerce')
        
    po_df['Quantity request'] = pd.to_numeric(po_df['Quantity request'], errors='coerce')
    po_df['Forecasted'] = po_df['Forecasted'].astype(str).str.lower()
    
    # Ensure essential PO columns are present (ocd_col_name handled above)
    base_required_po_cols = ['PO', 'Greige Code', 'CHD', 'Quantity request', 'Forecasted',
                             'Greige Name', 'ITEM', 'COLOR']
    # Add ocd_col_name to the list for checking if it wasn't missing initially
    required_po_cols_check = base_required_po_cols + [ocd_col_name]

    for col in required_po_cols_check: # Check all, including OCD
        if col not in po_df.columns:
            if col == ocd_col_name and ocd_col_name not in xls.parse('PO').columns : # Only warn if it was truly missing from file
                 pass # Already handled and warned about OCD_COL_NAME
            else:
                print(f"Warning: Column '{col}' not found in PO sheet. Adding it as NA.")
                po_df[col] = pd.NA

    po_df.dropna(subset=['PO', 'Greige Code', 'CHD', 'Quantity request'], inplace=True)

    # 1st Lot Status Data
    first_lot_df['DUE DATE'] = pd.to_datetime(first_lot_df['DUE DATE'], errors='coerce')
    
    # Handle both old and new column names
    column_mapping = {
        'DSM Code': 'Greige Code',
        'CPT Name': 'Greige Name'
    }
    
    # Modified rename logic for robustness
    cols_to_rename_in_first_lot = {
        old_col: new_col 
        for old_col, new_col in column_mapping.items() 
        if old_col in first_lot_df.columns
    }
    if cols_to_rename_in_first_lot:
        first_lot_df = first_lot_df.rename(columns=cols_to_rename_in_first_lot)
    
    # Ensure required columns exist
    required_columns = ['Greige Code', 'COLOR', 'STATUS']
    for col in required_columns:
        if col not in first_lot_df.columns:
            print(f"Warning: Column '{col}' not found in 1ST LOT STATUS sheet. Adding it as NA.")
            first_lot_df[col] = pd.NA
    
    # Convert Greige Code to string type
    if 'Greige Code' in first_lot_df.columns:
        first_lot_df['Greige Code'] = first_lot_df['Greige Code'].astype(str)
    
    # Ensure STATUS is string and clean it
    if 'STATUS' in first_lot_df.columns:
        first_lot_df['STATUS'] = first_lot_df['STATUS'].astype(str).str.strip().str.upper()
    
    first_lot_df.dropna(subset=['Greige Code', 'COLOR', 'STATUS'], inplace=True)

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
    # Ensure 'Greige Code' is the target name after potential renames from 'Mã Vải' or 'DSM Code'
    rename_map_stock = {}
    if 'Mã Vải' in df.columns and 'Greige Code' not in df.columns:
        rename_map_stock['Mã Vải'] = 'Greige Code'
    elif 'DSM Code' in df.columns and 'Greige Code' not in df.columns: # If DSM Code was an intermediate name
        rename_map_stock['DSM Code'] = 'Greige Code'
    
    # Common renames
    rename_map_stock.update({
        'ETA': 'Greige ETA',
        'Available': 'Greige Incoming'
    })
    df = df.rename(columns=rename_map_stock)
    
    # Convert Greige Code to string
    if 'Greige Code' in df.columns:
        df['Greige Code'] = df['Greige Code'].astype(str)
    
    # Convert dates and numerics for other columns
    df['Greige ETA'] = pd.to_datetime(df['Greige ETA'], errors='coerce')
    df['Greige Incoming'] = pd.to_numeric(df['Greige Incoming'], errors='coerce')
    
    return df

def clean_po_data(df):
    """Cleans and preprocesses PO data."""
    # Handle renaming for Greige Code (from 'Mã Vải' or 'DSM Code')
    if 'Mã Vải' in df.columns and 'Greige Code' not in df.columns and 'DSM Code' not in df.columns :
        df.rename(columns={'Mã Vải': 'Greige Code'}, inplace=True)
    elif 'DSM Code' in df.columns and 'Greige Code' not in df.columns:
        df.rename(columns={'DSM Code': 'Greige Code'}, inplace=True)
    
    # If 'Greige Code' might have NaNs and 'Mã Vải'/'DSM Code' could fill them (after primary rename)
    # This logic might need refinement based on exact precedence if multiple source columns exist
    if 'Greige Code' in df.columns:
        if 'Mã Vải' in df.columns and df['Greige Code'].isnull().any():
             df['Greige Code'].fillna(df['Mã Vải'], inplace=True)
        # Similar fill from 'DSM Code' if it was an alternative source before renaming to Greige Code

    # Convert Greige Code to string
    if 'Greige Code' in df.columns:
        df['Greige Code'] = df['Greige Code'].astype(str)

    # Convert dates and numerics for other columns
    date_columns = ['CHD', 'OCD( Order Creation Day)']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    
    df['Quantity request'] = pd.to_numeric(df['Quantity request'], errors='coerce')
    df['Forecasted'] = df['Forecasted'].astype(str).str.lower()
    
    return df

def clean_first_lot_data(df):
    """Cleans and preprocesses 1ST LOT STATUS data."""
    # Print debug information before cleaning
    print("\nDebug: First few rows of first_lot_df before cleaning in clean_first_lot_data:")
    print(df.head())
    
    # Rename columns if they exist
    column_mapping = {
        'DSM Code': 'Greige Code',
        'CPT Name': 'Greige Name'
    }
    for old_col, new_col in column_mapping.items():
        if old_col in df.columns:
            df.rename(columns={old_col: new_col}, inplace=True)
    
    # Convert Greige Code to integer
    df['Greige Code'] = pd.to_numeric(df['Greige Code'], errors='coerce')
    
    # Convert dates
    df['DUE DATE'] = pd.to_datetime(df['DUE DATE'], errors='coerce')
    
    # Clean STATUS values
    df['STATUS'] = df['STATUS'].str.strip().str.upper()
    
    # Print debug information after cleaning
    print("\nDebug: First few rows of first_lot_df after cleaning in clean_first_lot_data:")
    print(df.head())
    
    return df 