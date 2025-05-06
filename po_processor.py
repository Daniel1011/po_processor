import pandas as pd
import datetime
import calendar
from collections import defaultdict

# --- Configuration ---
INPUT_EXCEL_FILE = 'input.xlsx'  # Assuming the input file name
OUTPUT_EXCEL_FILE = 'output.xlsx'
TODAY_DATE_STR = '2025-05-06'
MONTHLY_CAPACITY = 10000
LEAD_TIME_DAYS = 40
FAR_FUTURE_DATE_STR = '2200-12-31'

# Convert string dates to datetime objects
TODAY_DATE = pd.to_datetime(TODAY_DATE_STR)
FAR_FUTURE_DATE = pd.to_datetime(FAR_FUTURE_DATE_STR)

# --- Helper Functions ---

def get_daily_capacity(date):
    """Calculates the daily production capacity for a given date."""
    days_in_month = calendar.monthrange(date.year, date.month)[1]
    if days_in_month == 0: # Should not happen with valid dates
        return 0
    return MONTHLY_CAPACITY / days_in_month

# --- Main Processing Logic ---

def process_pos(input_file=INPUT_EXCEL_FILE, output_file=OUTPUT_EXCEL_FILE):
    """Reads input Excel, processes POs, calculates ETDs, and writes output Excel."""
    print(f"Starting PO processing for {input_file}...")
    print(f"Current date set to: {TODAY_DATE.strftime('%Y-%m-%d')}")
    print(f"Monthly capacity: {MONTHLY_CAPACITY}")
    print(f"Lead time: {LEAD_TIME_DAYS} days")

    try:
        # 1. Read Input Data
        print("Reading Stock data...")
        stock_df = pd.read_excel(input_file, sheet_name='Stock')
        print("Reading PO data...")
        po_df = pd.read_excel(input_file, sheet_name='PO')

        # Convert date columns, coercing errors to NaT (Not a Time)
        stock_df['ETA'] = pd.to_datetime(stock_df['ETA'], errors='coerce')
        po_df['CHD'] = pd.to_datetime(po_df['CHD'], errors='coerce')

        # Validate data: Drop rows with invalid dates or missing crucial info
        stock_df.dropna(subset=['Mã Vải', 'ETA', 'Available'], inplace=True)
        po_df.dropna(subset=['PO', 'Mã Vải', 'CHD', 'Quantity request', 'Forecasted'], inplace=True)
        # Ensure quantity is numeric
        po_df['Quantity request'] = pd.to_numeric(po_df['Quantity request'], errors='coerce')
        po_df.dropna(subset=['Quantity request'], inplace=True)
        stock_df['Available'] = pd.to_numeric(stock_df['Available'], errors='coerce')
        stock_df.dropna(subset=['Available'], inplace=True)


        print("Input data read successfully.")

    except FileNotFoundError:
        print(f"Error: Input file '{input_file}' not found.")
        return
    except Exception as e:
        print(f"Error reading or validating input file: {e}")
        return

    # 2. Prepare Stock Data
    print("Preparing stock data...")
    stock_data = defaultdict(lambda: {'available': 0, 'incoming': []})
    # Convert to dictionary for faster lookup and modification
    stock_records = stock_df.to_dict('records')

    for record in stock_records:
        fabric_code = record['Mã Vải']
        eta = record['ETA']
        quantity = record['Available']

        if pd.isna(fabric_code) or pd.isna(eta) or pd.isna(quantity) or quantity <= 0:
            continue # Skip invalid stock entries

        if eta <= TODAY_DATE:
            stock_data[fabric_code]['available'] += quantity
        else:
            stock_data[fabric_code]['incoming'].append({'eta': eta, 'quantity': quantity})

    # Sort incoming batches by ETA
    for fabric_code in stock_data:
        stock_data[fabric_code]['incoming'].sort(key=lambda x: x['eta'])
    print("Stock data prepared.")


    # 3. Prioritize POs
    print("Prioritizing POs...")
    # Ensure 'Forecasted' is lowercase string for comparison
    po_df['Forecasted'] = po_df['Forecasted'].astype(str).str.lower()

    forecasted_pos = po_df[po_df['Forecasted'] == 'yes'].sort_values(by='CHD')
    non_forecasted_pos = po_df[po_df['Forecasted'] != 'yes'].sort_values(by='CHD')
    prioritized_pos = pd.concat([forecasted_pos, non_forecasted_pos], ignore_index=True)
    print(f"Total POs to process: {len(prioritized_pos)}")


    # 4. Allocate Stock to POs & Determine Desired Start Date
    print("Allocating stock and determining desired start dates...")
    po_schedule_info = [] # List to store POs with desired start date

    # Convert df to list of dicts for easier iteration and modification
    po_list = prioritized_pos.to_dict('records')

    for po in po_list:
        po_id = po['PO']
        fabric_code = po['Mã Vải']
        requested_qty = po['Quantity request']
        desired_start_date = FAR_FUTURE_DATE # Default to far future if stock insufficient
        allocated = False

        if fabric_code in stock_data:
            stock = stock_data[fabric_code]
            needed = requested_qty
            current_available = stock['available']

            # Try allocating from currently available stock
            if current_available >= needed:
                stock['available'] -= needed
                desired_start_date = TODAY_DATE
                allocated = True
            else:
                # Use all available stock and start consuming incoming batches
                needed -= current_available
                stock['available'] = 0
                last_eta_used = TODAY_DATE # Track the ETA of the last batch needed

                # Use a copy for iteration as we might modify the list
                incoming_batches_copy = list(stock['incoming'])
                stock['incoming'] = [] # Reset original list, will rebuild with remaining

                for i, batch in enumerate(incoming_batches_copy):
                    batch_eta = batch['eta']
                    batch_qty = batch['quantity']
                    last_eta_used = max(last_eta_used, batch_eta) # Update last ETA

                    if batch_qty >= needed:
                        # This batch fulfills the remaining need
                        remaining_in_batch = batch_qty - needed
                        if remaining_in_batch > 0:
                            # Add back the remainder of this batch
                            stock['incoming'].append({'eta': batch_eta, 'quantity': remaining_in_batch})
                        # Add back any subsequent batches untouched
                        stock['incoming'].extend(incoming_batches_copy[i+1:])
                        needed = 0
                        desired_start_date = last_eta_used
                        allocated = True
                        break # Allocation complete for this PO
                    else:
                        # Consume the entire batch and continue
                        needed -= batch_qty
                        # Don't add this batch back as it's fully consumed

                # Re-sort incoming batches after potential modifications
                stock['incoming'].sort(key=lambda x: x['eta'])

        if not allocated:
             print(f"Warning: Insufficient stock for PO {po_id} (Fabric: {fabric_code}, Qty: {requested_qty}). Desired start date set to {FAR_FUTURE_DATE_STR}.")

        # Record the result for scheduling
        po_schedule_info.append({
            'PO': po_id,
            'Quantity': requested_qty,
            'DesiredStartDate': desired_start_date
        })

    print("Stock allocation complete.")

    # ---> Prepare remaining stock data <--- START
    final_stock_list = []
    for fabric_code, data in stock_data.items():
        incoming_batches_str = "; ".join([
            f"ETA: {b['eta'].strftime('%Y-%m-%d')}, Qty: {b['quantity']}"
            for b in data['incoming']
        ])
        final_stock_list.append({
            'Mã Vải': fabric_code,
            'Remaining Available (as of {TODAY_DATE.strftime("%Y-%m-%d")})': data['available'],
            'Remaining Incoming Batches': incoming_batches_str if incoming_batches_str else "None"
        })
    remaining_stock_df = pd.DataFrame(final_stock_list)
    # ---> Prepare remaining stock data <--- END

    # 5. Schedule Production
    print("Scheduling production based on capacity...")
    # Sort POs by their desired start date
    po_schedule_info.sort(key=lambda x: x['DesiredStartDate'])

    daily_production_schedule = defaultdict(float) # Tracks units scheduled per day
    po_final_schedule = []

    for po_info in po_schedule_info:
        po_id = po_info['PO']
        quantity = po_info['Quantity']
        current_date = po_info['DesiredStartDate']

        if current_date == FAR_FUTURE_DATE:
            # Cannot schedule POs with insufficient stock
            actual_start_date = FAR_FUTURE_DATE
        else:
            # Find the first available date from desired start date onwards
            while True:
                daily_cap = get_daily_capacity(current_date)
                # Check if adding this PO exceeds capacity for the day
                if daily_production_schedule[current_date] + quantity <= daily_cap:
                    daily_production_schedule[current_date] += quantity
                    actual_start_date = current_date
                    break
                else:
                    # Move to the next day
                    current_date += datetime.timedelta(days=1)
                    if current_date.year > FAR_FUTURE_DATE.year + 1: # Safety break
                        print(f"Warning: Could not find capacity slot for PO {po_id} within reasonable time. Setting start date to far future.")
                        actual_start_date = FAR_FUTURE_DATE
                        break

        po_final_schedule.append({
            'PO': po_id,
            'ActualStartDate': actual_start_date
        })

    print("Production scheduling complete.")


    # 6. Calculate ETD
    print("Calculating ETDs...")
    results = []
    for schedule_item in po_final_schedule:
        po_id = schedule_item['PO']
        actual_start_date = schedule_item['ActualStartDate']

        if actual_start_date == FAR_FUTURE_DATE:
            etd = FAR_FUTURE_DATE # Or None, or specific indicator
            print(f"PO {po_id}: Cannot determine ETD due to insufficient stock or capacity issues.")
        else:
            etd = actual_start_date + datetime.timedelta(days=LEAD_TIME_DAYS)

        results.append({'PO': po_id, 'ETD': etd})

    print("ETD calculation complete.")


    # 7. Output Results
    print(f"Writing results to {output_file}...")
    etd_output_df = pd.DataFrame(results)
    # Format ETD column for better readability in Excel
    etd_output_df['ETD'] = etd_output_df['ETD'].dt.strftime('%Y-%m-%d')
    # Replace far future date string for clarity
    etd_output_df['ETD'] = etd_output_df['ETD'].replace(FAR_FUTURE_DATE.strftime('%Y-%m-%d'), 'Insufficient Stock/Capacity')


    try:
        # Use ExcelWriter to save multiple sheets
        with pd.ExcelWriter(output_file) as writer:
            etd_output_df.to_excel(writer, sheet_name='ETD Results', index=False)
            remaining_stock_df.to_excel(writer, sheet_name='Remaining Stock', index=False)
        print(f"Successfully wrote output to {output_file}")
    except Exception as e:
        print(f"Error writing output file: {e}")


# --- Entry Point ---
if __name__ == "__main__":
    # Create dummy input file for testing if it doesn't exist
    try:
        pd.read_excel(INPUT_EXCEL_FILE)
    except FileNotFoundError:
        print(f"Creating dummy input file: {INPUT_EXCEL_FILE}")
        # Create dummy Stock data
        stock_data_dummy = {
            'Mã Vải': ['A', 'A', 'B', 'B', 'C'],
            'ETA': ['2025-05-06', '2025-06-06', '2025-06-07', '2025-07-11', '2025-05-01'],
            'Available': [2000, 3000, 4000, 300, 5000]
        }
        stock_df_dummy = pd.DataFrame(stock_data_dummy)
        stock_df_dummy['ETA'] = pd.to_datetime(stock_df_dummy['ETA'])

        # Create dummy PO data
        po_data_dummy = {
            'PO': ['PO1 A', 'PO2 B', 'PO3 A', 'PO4 C', 'PO5 B'],
            'Mã Vải': ['A', 'B', 'A', 'C', 'B'],
            'SPL': ['X', 'Y', 'X', 'Z', 'Y'],
            'CHD': ['2025-08-02', '2025-09-01', '2025-07-15', '2025-06-20', '2025-10-10'],
            'Quantity request': [500, 8000, 2000, 4000, 1000],
            'Forecasted': ['yes', 'no', 'yes', 'no', 'no']
        }
        po_df_dummy = pd.DataFrame(po_data_dummy)
        po_df_dummy['CHD'] = pd.to_datetime(po_df_dummy['CHD'])

        with pd.ExcelWriter(INPUT_EXCEL_FILE) as writer:
            stock_df_dummy.to_excel(writer, sheet_name='Stock', index=False)
            po_df_dummy.to_excel(writer, sheet_name='PO', index=False)
        print(f"Dummy file {INPUT_EXCEL_FILE} created.")

    # Run the main processing function
    process_pos() 