import pandas as pd

def calculate_second_etd(draft_etd_df, first_lot_df, far_future_date):
    """Calculates the 2nd ETD based on 1st Lot Status."""
    print("Step 2: Calculating 2nd ETD with 1st Lot Status...")
    
    # Print debug information before cleaning
    print("\nDebug: First few rows of first_lot_df before cleaning:")
    print(first_lot_df[['DSM Code', 'COLOR', 'STATUS', 'DUE DATE']].head())
    
    # Clean up COLOR values in both dataframes by removing extra spaces
    draft_etd_df['COLOR'] = draft_etd_df['COLOR'].str.strip().str.replace(r'\s+', ' ')
    first_lot_df['COLOR'] = first_lot_df['COLOR'].str.strip().str.replace(r'\s+', ' ')
    
    # Convert DSM Code to integer in both dataframes
    draft_etd_df['DSM Code'] = draft_etd_df['DSM Code'].astype(int)
    first_lot_df['DSM Code'] = first_lot_df['DSM Code'].astype(int)
    
    # Convert STATUS to uppercase for consistent comparison
    first_lot_df['STATUS'] = first_lot_df['STATUS'].str.strip().str.upper()
    
    # Ensure all date columns are datetime objects
    date_columns = ['OCD( Order Creation Day)', 'CHD', 'Draft ETD', 'DUE DATE']
    for col in date_columns:
        if col in draft_etd_df.columns:
            draft_etd_df[col] = pd.to_datetime(draft_etd_df[col], errors='coerce')
        if col in first_lot_df.columns:
            first_lot_df[col] = pd.to_datetime(first_lot_df[col], errors='coerce')
    
    # Print debug information after cleaning
    print("\nDebug: First few rows of first_lot_df after cleaning:")
    print(first_lot_df[['DSM Code', 'COLOR', 'STATUS', 'DUE DATE']].head())
    
    # Merge with 1st Lot Status data
    first_lot_df_renamed = first_lot_df[['DSM Code', 'COLOR', 'STATUS', 'DUE DATE']].rename(
        columns={'STATUS': '1ST LOT STATUS'}
    )
    
    # Print debug information before merge
    print("\nDebug: First few rows of draft_etd_df before merge:")
    print(draft_etd_df[['DSM Code', 'COLOR', 'Draft ETD']].head())
    print("\nDebug: First few rows of first_lot_df_renamed before merge:")
    print(first_lot_df_renamed[['DSM Code', 'COLOR', '1ST LOT STATUS', 'DUE DATE']].head())
    
    # Use a copy to avoid SettingWithCopyWarning
    draft_etd_df_merged = pd.merge(
        draft_etd_df.copy(),
        first_lot_df_renamed,
        on=['DSM Code', 'COLOR'],
        how='left'
    )
    
    # Print debug information after merge
    print("\nDebug: First few rows of merged data:")
    print(draft_etd_df_merged[['DSM Code', 'COLOR', '1ST LOT STATUS', 'DUE DATE', 'Draft ETD']].head())
    print("\nDebug: Number of rows in merged data:", len(draft_etd_df_merged))

    # Calculate 2nd ETD
    def get_2nd_etd(row):
        draft_etd = row['Draft ETD']
        lot_status = str(row['1ST LOT STATUS']).strip().upper() if pd.notna(row['1ST LOT STATUS']) else None
        due_date = row['DUE DATE'] 

        if pd.isna(draft_etd) or draft_etd == far_future_date:
            return draft_etd 

        if lot_status is None:
            return draft_etd

        if lot_status == 'OK':
            return draft_etd
        elif lot_status == 'EXPIRED':
            if pd.isna(due_date):
                return draft_etd
            if due_date <= draft_etd:
                return draft_etd
            else: 
                return due_date
        return draft_etd

    draft_etd_df_merged['2nd ETD'] = draft_etd_df_merged.apply(get_2nd_etd, axis=1)
    
    # Ensure 2nd ETD is datetime
    draft_etd_df_merged['2nd ETD'] = pd.to_datetime(draft_etd_df_merged['2nd ETD'], errors='coerce')
    
    print("Step 2 finished.")
    return draft_etd_df_merged 