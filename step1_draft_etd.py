import pandas as pd
from collections import defaultdict

def calculate_draft_etd_and_remaining_stock(stock_df, po_df, today_date, lead_time_days, far_future_date, ocd_col_name):
    """Calculates Draft ETD for POs and prepares the remaining stock summary."""
    print("Step 1: Calculating Draft ETD and Preparing Remaining Stock...")

    # Prepare stock_data dictionary from stock_df
    stock_data_dict = defaultdict(lambda: {'available_on_hand': 0, 'incoming_batches': []})
    for _, row in stock_df.iterrows():
        dsm_code = row['DSM Code']
        eta = row['Greige ETA']
        quantity = row['Greige Incoming']
        if eta <= today_date:
            stock_data_dict[dsm_code]['available_on_hand'] += quantity
        else:
            stock_data_dict[dsm_code]['incoming_batches'].append({'eta': eta, 'quantity': quantity})

  
    for dsm_code in stock_data_dict:
        stock_data_dict[dsm_code]['incoming_batches'].sort(key=lambda x: x['eta'])

    # Prioritize POs
    po_df_sorted = po_df.copy()
    po_df_sorted['Forecasted_Sort'] = po_df_sorted['Forecasted'].apply(lambda x: 0 if x == 'yes' else 1)
    po_df_sorted = po_df_sorted.sort_values(by=['Forecasted_Sort', 'CHD'])

    draft_etd_results = []
    stock_for_draft_etd = defaultdict(lambda: {'available_on_hand': 0, 'incoming_batches': []})

    for code, data in stock_data_dict.items():
        stock_for_draft_etd[code]['available_on_hand'] = data['available_on_hand']
        stock_for_draft_etd[code]['incoming_batches'] = [batch.copy() for batch in data['incoming_batches']]
     

    for _, po_row in po_df_sorted.iterrows():
        dsm_code = po_row['DSM Code']
        requested_qty = po_row['Quantity request']
        draft_etd_val = far_future_date 
        current_stock_info = stock_for_draft_etd[dsm_code]
        needed = requested_qty
        material_available_date = far_future_date

        if current_stock_info['available_on_hand'] >= needed:
            current_stock_info['available_on_hand'] -= needed
            material_available_date = today_date
            needed = 0
        else:
            if current_stock_info['available_on_hand'] > 0:
                 material_available_date = today_date # Used some on-hand
            needed -= current_stock_info['available_on_hand']
            current_stock_info['available_on_hand'] = 0
            
            batches_to_keep = []
            for batch in current_stock_info['incoming_batches']:
                if needed == 0:
                    batches_to_keep.append(batch)
                    continue
                
                material_available_date = max(material_available_date if material_available_date != far_future_date else batch['eta'], batch['eta'])
                
                if batch['quantity'] >= needed:
                    batch['quantity'] -= needed
                    needed = 0
                    if batch['quantity'] > 0:
                        batches_to_keep.append(batch)
                    break 
                else:
                    needed -= batch['quantity']
            current_stock_info['incoming_batches'] = batches_to_keep

        if needed == 0:
            if material_available_date != far_future_date:
                draft_etd_val = material_available_date + pd.Timedelta(days=lead_time_days)
        
        result_row = po_row.to_dict()
        result_row['Draft ETD'] = draft_etd_val
        draft_etd_results.append(result_row)

    draft_etd_df = pd.DataFrame(draft_etd_results)

    # Prepare Remaining Stock DataFrame
    remaining_stock_list = []
    # CPT Name is sourced from po_df as per user request
    dsm_to_cpt_name_map = po_df.drop_duplicates(subset=['DSM Code']).set_index('DSM Code')['CPT Name'].to_dict()
    all_dsm_codes_for_remaining = set(stock_df['DSM Code'].unique()) | set(po_df['DSM Code'].unique())

    for dsm_code in sorted(list(all_dsm_codes_for_remaining)):
        stock_status = stock_for_draft_etd[dsm_code]
        cpt_name = dsm_to_cpt_name_map.get(dsm_code, '')
        remaining_on_hand = stock_status['available_on_hand']
        incoming_batches_str_list = []
        for batch in stock_status['incoming_batches']:
            if batch['quantity'] > 0:
                eta_str = batch['eta'].strftime('%Y-%m-%d') if pd.notna(batch['eta']) else 'N/A'
                incoming_batches_str_list.append(f"ETA: {eta_str}, Qty: {int(batch['quantity'])}")
        remaining_incoming_str = "; ".join(incoming_batches_str_list) if incoming_batches_str_list else "None"
        remaining_stock_list.append({
            'DSM Code': dsm_code,
            'CPT Name': cpt_name,
            f'Remaining Available (as of {today_date.strftime("%Y-%m-%d")})': remaining_on_hand,
            'Remaining Incoming Batches': remaining_incoming_str
        })
    remaining_stock_df = pd.DataFrame(remaining_stock_list)
    print("Step 1 finished.")
    return draft_etd_df, remaining_stock_df 