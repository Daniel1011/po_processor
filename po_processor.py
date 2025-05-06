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
    final_stock_list = [] # This will store dictionaries, each representing a row in the output
    for fabric_code, data in stock_data.items():
        remaining_available_on_hand = data['available'] # Stock with ETA <= TODAY_DATE

        if not data['incoming']:
            # Case 1: Fabric has remaining on-hand stock, but NO future incoming batches left.
            # Create one row for this fabric.
            final_stock_list.append({
                'Mã Vải': fabric_code,
                f'Remaining Available (as of {TODAY_DATE.strftime("%Y-%m-%d")})': remaining_available_on_hand,
                'Incoming Batch ETA': pd.NA,  # Placeholder for no ETA
                'Incoming Batch Quantity': pd.NA  # Placeholder for no Quantity
            })
        else:
            # Case 2: Fabric has remaining on-hand stock AND one or more future incoming batches.
            # Create a distinct row for EACH remaining incoming batch.
            for batch in data['incoming']:
                final_stock_list.append({
                    'Mã Vải': fabric_code, # Repeated for each batch of this fabric
                    f'Remaining Available (as of {TODAY_DATE.strftime("%Y-%m-%d")})': remaining_available_on_hand, # Repeated
                    'Incoming Batch ETA': batch['eta'], # Store the datetime object
                    'Incoming Batch Quantity': batch['quantity']
                })
                
    remaining_stock_df = pd.DataFrame(final_stock_list)

    # Format the 'Incoming Batch ETA' column to 'YYYY-MM-DD' string, handling potential NaT values
    if 'Incoming Batch ETA' in remaining_stock_df.columns:
         remaining_stock_df['Incoming Batch ETA'] = pd.to_datetime(remaining_stock_df['Incoming Batch ETA'], errors='coerce').dt.strftime('%Y-%m-%d')
    
    # If you want to ensure blank cells for NA quantities as well (optional, pandas usually handles this well for Excel)
    if 'Incoming Batch Quantity' in remaining_stock_df.columns:
        remaining_stock_df['Incoming Batch Quantity'] = pd.to_numeric(remaining_stock_df['Incoming Batch Quantity'], errors='coerce')

    # ---> Prepare remaining stock data <--- END

    # 5. Schedule Production
    print("Scheduling production based on capacity...")
    # Sort POs by their desired start date
    po_schedule_info.sort(key=lambda x: x['DesiredStartDate'])

    daily_production_schedule = defaultdict(float) # Tracks units scheduled per day
    po_final_schedule = []

    for po_info in po_schedule_info:
        po_id = po_info['PO']
        quantity_to_schedule = po_info['Quantity']
        scheduling_date = po_info['DesiredStartDate'] 
        
        actual_start_date_for_etd = FAR_FUTURE_DATE # Default for this PO

        if scheduling_date == FAR_FUTURE_DATE:
            # PO has insufficient stock, cannot be scheduled
            pass # actual_start_date_for_etd remains FAR_FUTURE_DATE
        else:
            # PO has stock, try to schedule its production
            first_day_production_assigned_this_po = None # Track the first day this PO starts production

            # Loop until the entire quantity for this PO is scheduled or we hit a safety limit
            while quantity_to_schedule > 0:
                daily_cap = get_daily_capacity(scheduling_date)
                # Capacity available on this specific day, considering other POs already scheduled
                available_capacity_on_day = daily_cap - daily_production_schedule[scheduling_date]

                if available_capacity_on_day > 0:
                    if first_day_production_assigned_this_po is None:
                        # This is the first day production starts for this PO
                        first_day_production_assigned_this_po = scheduling_date
                    
                    # Determine how much of the current PO can be produced today
                    producible_today = min(quantity_to_schedule, available_capacity_on_day)
                    
                    daily_production_schedule[scheduling_date] += producible_today
                    quantity_to_schedule -= producible_today

                if quantity_to_schedule <= 0: # PO fully scheduled
                    actual_start_date_for_etd = first_day_production_assigned_this_po
                    break # Exit the while loop for this PO

                # Move to the next day to schedule remaining quantity
                scheduling_date += datetime.timedelta(days=1)

                # Safety break: If scheduling date goes into the FAR_FUTURE_DATE's year
                # and the PO is still not fully scheduled, mark it as unschedulable.
                if scheduling_date.year >= FAR_FUTURE_DATE.year:
                    print(f"Warning: PO {po_id} (remaining qty: {quantity_to_schedule}) could not be *fully* scheduled before {FAR_FUTURE_DATE_STR}. Its ETD will indicate insufficient capacity.")
                    actual_start_date_for_etd = FAR_FUTURE_DATE # Mark as unschedulable
                    break # Exit the while loop for this PO
            
            # If loop finished and quantity_to_schedule is 0, actual_start_date_for_etd is already set.
            # If loop broke due to safety, actual_start_date_for_etd is FAR_FUTURE_DATE.

        po_final_schedule.append({
            'PO': po_id,
            'ActualStartDate': actual_start_date_for_etd
        })

    print("Production scheduling complete.")


    # 6. Calculate ETD
    print("Calculating ETDs...")
    results = [] # This list will contain dictionaries like {'PO': po_id, 'ETD': etd_datetime_object}
    for schedule_item in po_final_schedule:
        po_id = schedule_item['PO']
        actual_start_date = schedule_item['ActualStartDate']

        if actual_start_date == FAR_FUTURE_DATE:
            etd = FAR_FUTURE_DATE # ETD is a datetime object (or FAR_FUTURE_DATE)
            print(f"PO {po_id}: Cannot determine ETD due to insufficient stock or capacity issues.")
        else:
            etd = actual_start_date + datetime.timedelta(days=LEAD_TIME_DAYS)

        results.append({'PO': po_id, 'ETD': etd})

    print("ETD calculation complete.")


    # 7. Output Results
    print(f"Writing results to {output_file}...")

    # Convert the ETD results list to a DataFrame
    etd_info_df = pd.DataFrame(results) # Contains 'PO' and 'ETD' (ETD is datetime)

    # The 'prioritized_pos' DataFrame (from Step 3) contains all original PO details
    # including 'Mã Vải', 'SPL', 'CHD', 'Quantity request', 'Forecasted'.
    # Merge the original PO details with the ETD information.
    # 'prioritized_pos' columns: PO, Mã Vải, SPL, CHD (datetime), Quantity request, Forecasted
    # 'etd_info_df' columns: PO, ETD (datetime)
    
    # Ensure 'CHD' in 'prioritized_pos' is in a consistent state if it hasn't been modified since read
    # (it was read as datetime and should still be)
    
    etd_output_df = pd.merge(prioritized_pos, etd_info_df, on='PO', how='left')
    
    # Format date columns for output
    # CHD is already a datetime object from the initial read and prioritization.
    etd_output_df['CHD'] = etd_output_df['CHD'].dt.strftime('%Y-%m-%d')
    
    # ETD is also a datetime object (or FAR_FUTURE_DATE which is also datetime)
    etd_output_df['ETD'] = etd_output_df['ETD'].dt.strftime('%Y-%m-%d')
    # Replace far future date string for ETD clarity
    etd_output_df['ETD'] = etd_output_df['ETD'].replace(FAR_FUTURE_DATE.strftime('%Y-%m-%d'), 'Insufficient Stock/Capacity')

    # Define the desired column order for the output sheet
    final_output_columns = [
        'PO', 'Mã Vải', 'SPL', 'CHD', 
        'Quantity request', 'Forecasted', 'ETD'
    ]
    etd_output_df = etd_output_df[final_output_columns] # Reorder/select columns

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