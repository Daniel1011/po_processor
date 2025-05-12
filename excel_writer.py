import pandas as pd

def write_output_to_excel(output_file, draft_etd_df, remaining_stock_df, final_etd_df, far_future_date, ocd_col_name):
    """Writes the processed DataFrames to the output Excel file."""
    print(f"Writing results to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl',
                        date_format='YYYY-MM-DD',
                        datetime_format='YYYY-MM-DD') as writer:
        # DRAFT ETD sheet
        draft_etd_output_cols = [
            'SPL', 'FG name', 'Season', 'Local/ Export', 'PO', ocd_col_name, 'CHD',
            'DSM Code', 'CPT Name', 'ITEM', 'COLOR', 'Quantity request', 'Forecasted',
            'Draft ETD', '1ST LOT STATUS', 'DUE DATE', '2nd ETD'
        ]
        # Ensure all columns exist, add if not for safety
        df_to_write_draft = draft_etd_df.copy()
        for col in draft_etd_output_cols:
            if col not in df_to_write_draft.columns:
                df_to_write_draft[col] = pd.NA
        
        df_to_write_draft = df_to_write_draft[draft_etd_output_cols]
        date_cols_draft = [ocd_col_name, 'CHD', 'Draft ETD', 'DUE DATE', '2nd ETD']
        for col in date_cols_draft:
            if col in df_to_write_draft.columns:
                # Convert to datetime and format as YYYY-MM-DD
                df_to_write_draft[col] = pd.to_datetime(df_to_write_draft[col], errors='coerce')
                df_to_write_draft[col] = df_to_write_draft[col].apply(
                    lambda x: "Insufficient Stock/Capacity" if pd.isna(x) or x == far_future_date else x.strftime('%Y-%m-%d')
                )
        df_to_write_draft.to_excel(writer, sheet_name='DRAFT ETD', index=False)

        # REMAINING STOCK sheet
        remaining_stock_df.to_excel(writer, sheet_name='REMAINING STOCK', index=False)

        # FINAL ETD sheet
        final_etd_output_cols = [
            'SPL', 'FG name', 'Season', 'Local/ Export', 'PO', ocd_col_name, 'CHD',
            'DSM Code', 'CPT Name', 'ITEM', 'COLOR', 'Quantity request', 'Forecasted',
            'Draft ETD', '1ST LOT STATUS', 'DUE DATE', '2nd ETD',
            'DEVIDED QUANTITY 1ST', 'DATE 1ST BATCH',
            'DEVIDED QUANTITY 2ND', 'DATE 2ND BATCH',
            'FINAL QUANTITY', 'FINAL ETD'
        ]
        df_to_write_final = final_etd_df.copy()
        for col in final_etd_output_cols:
            if col not in df_to_write_final.columns:
                df_to_write_final[col] = pd.NA
        
        df_to_write_final = df_to_write_final[final_etd_output_cols]
        
        # Handle date columns
        date_cols_final = [ocd_col_name, 'CHD', 'Draft ETD', 'DUE DATE', '2nd ETD',
                           'DATE 1ST BATCH', 'DATE 2ND BATCH', 'FINAL ETD']
        for col in date_cols_final:
            if col in df_to_write_final.columns:
                # Convert to datetime and format as YYYY-MM-DD
                df_to_write_final[col] = pd.to_datetime(df_to_write_final[col], errors='coerce')
                df_to_write_final[col] = df_to_write_final[col].apply(
                    lambda x: "Insufficient Stock/Capacity" if pd.isna(x) or x == far_future_date else x.strftime('%Y-%m-%d')
                )
        
        # Handle quantity columns
        quantity_cols = ['Quantity request', 'DEVIDED QUANTITY 1ST', 'DEVIDED QUANTITY 2ND', 'FINAL QUANTITY']
        for col in quantity_cols:
            if col in df_to_write_final.columns:
                # Convert to numeric and handle NaN values
                df_to_write_final[col] = pd.to_numeric(df_to_write_final[col], errors='coerce')
                df_to_write_final[col] = df_to_write_final[col].apply(
                    lambda x: 0 if pd.isna(x) else x
                )
        
        # Handle status columns
        status_cols = ['1ST LOT STATUS', 'Forecasted']
        for col in status_cols:
            if col in df_to_write_final.columns:
                # Convert to uppercase for consistency
                df_to_write_final[col] = df_to_write_final[col].apply(
                    lambda x: str(x).upper() if pd.notna(x) else x
                )
        
        df_to_write_final.to_excel(writer, sheet_name='FINAL ETD', index=False)

    print(f"Successfully wrote output to {output_file}") 