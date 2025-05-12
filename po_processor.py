import pandas as pd
import datetime
from collections import defaultdict
import uuid
# Import constants and helper from config.py
import config 
# Import functions from the new modules
from data_loader import load_and_prepare_data
from step1_draft_etd import calculate_draft_etd_and_remaining_stock
from step2_second_etd import calculate_second_etd
from step3_final_etd import schedule_production_and_final_etd
from excel_writer import write_output_to_excel

# --- Main Processing Logic --- (Orchestrator)
def process_fabric_management(input_file=config.INPUT_EXCEL_FILE, output_file=config.OUTPUT_EXCEL_FILE):
    """
    Orchestrates the fabric stock management and ETD calculation process
    by calling functions from specialized modules.
    """
    print(f"Starting fabric stock management processing for {input_file}...")
    print(f"Current date set to: {config.TODAY_DATE.strftime('%Y-%m-%d')}")
    print(f"Lead time: {config.LEAD_TIME_DAYS} days")

    try:
        # Step 0: Load and Prepare Data
        stock_df, po_df, first_lot_df, capacity_status_df = load_and_prepare_data(
            input_file,
            config.OCD_COL_NAME
        )

    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
        return
    except Exception as e:
        print(f"Error reading or validating input file: {e}")
        return

    # Step 1: Calculate Draft ETD & Prepare Remaining Stock
    draft_etd_df, remaining_stock_df = calculate_draft_etd_and_remaining_stock(
        stock_df, 
        po_df, 
        config.TODAY_DATE, 
        config.LEAD_TIME_DAYS, 
        config.FAR_FUTURE_DATE,
        config.OCD_COL_NAME
    )

    # Step 2: Calculate 2nd ETD with 1st Lot Status
    # Ensure draft_etd_df is passed to step 2
    draft_etd_df_with_2nd_etd = calculate_second_etd(
        draft_etd_df, 
        first_lot_df, 
        config.FAR_FUTURE_DATE
    )

    # Step 3: Schedule Production and Final ETD
    # Pass the output of step 2 to step 3
    final_etd_df = schedule_production_and_final_etd(
        draft_etd_df_with_2nd_etd, 
        capacity_status_df, 
        config.TODAY_DATE, 
        config.LEAD_TIME_DAYS, 
        config.FAR_FUTURE_DATE, 
        config.CAPACITY_TOLERANCE, 
        config.MIN_CAPACITY_REMAIN
    )

    # Step 4: Output Results to Excel
    write_output_to_excel(
        output_file, 
        draft_etd_df_with_2nd_etd, # This contains Draft ETD, 1st Lot, and 2nd ETD
        remaining_stock_df, 
        final_etd_df, 
        config.FAR_FUTURE_DATE, 
        config.OCD_COL_NAME
    )

# --- Entry Point --- (Remains the same)
if __name__ == "__main__":
    # Example: Create a dummy input file if it doesn't exist (for basic testing)
    # This part should be more elaborate to match the new complex structure if used for testing.
    # For now, it's better to ensure 'PO - Request.xlsx' exists with the correct sheets.
    try:
        # Check if the specific input file exists
        with open(config.INPUT_EXCEL_FILE, 'rb') as f:
            pass
        print(f"Found input file: {config.INPUT_EXCEL_FILE}. Running main process...")
        process_fabric_management()
    except FileNotFoundError:
        print(f"ERROR: Input file '{config.INPUT_EXCEL_FILE}' not found.")
        print("Please ensure the input Excel file with all required sheets (Stock, PO, 1ST LOT STATUS, Capacity Status) is in the same directory as the script.")
        print("You may need to run a script to generate a dummy input file first (e.g., create_template_file.py).") 