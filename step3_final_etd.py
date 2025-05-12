import pandas as pd

def schedule_production_and_final_etd(draft_etd_df_with_2nd_etd, capacity_status_df, today_date, lead_time_days, far_future_date, capacity_tolerance, min_capacity_remain):
    """Schedules production based on capacity and calculates the Final ETD."""
    print("Step 3: Scheduling Production and Final ETD...")
    
    live_capacity_df = capacity_status_df.set_index('CAPACITY DATE')[['CAPACITY REMAIN']].copy()

    schedule_pos_df = draft_etd_df_with_2nd_etd.copy()
    schedule_pos_df['Is_Schedulable_ETD'] = schedule_pos_df['2nd ETD'].apply(lambda x: x != far_future_date and pd.notna(x))
    schedule_pos_df['2nd ETD_datetime'] = pd.to_datetime(schedule_pos_df['2nd ETD'], errors='coerce')
    schedule_pos_df = schedule_pos_df.sort_values(by=['Is_Schedulable_ETD', '2nd ETD_datetime', 'CHD'], ascending=[False, True, True])

    final_etd_results = []

    for _, po_row in schedule_pos_df.iterrows():
        qty_to_schedule = po_row['Quantity request']
        second_etd_dt = po_row['2nd ETD_datetime']

        dev_qty1, date1, dev_qty2, date2 = pd.NA, pd.NaT, pd.NA, pd.NaT
        final_scheduled_qty = 0
        actual_prod_end_date = far_future_date

        if pd.isna(second_etd_dt) or second_etd_dt == far_future_date:
            pass 
        else:
            target_prod_completion_date = second_etd_dt - pd.Timedelta(days=lead_time_days)
            
            if target_prod_completion_date < today_date:
                current_day_for_scheduling = today_date
            else:
                current_day_for_scheduling = max(today_date, target_prod_completion_date - pd.Timedelta(days=30))

            scheduled_this_po = False
            search_limit_date = max(target_prod_completion_date, today_date) + pd.Timedelta(days=365)

            while current_day_for_scheduling <= search_limit_date:
                cap_day1_val = live_capacity_df.loc[live_capacity_df.index == current_day_for_scheduling, 'CAPACITY REMAIN'].iloc[0] if current_day_for_scheduling in live_capacity_df.index else 0
                schedulable_on_day1 = max(0, cap_day1_val + capacity_tolerance)
                
                if qty_to_schedule < 1000:
                    if schedulable_on_day1 >= qty_to_schedule and (cap_day1_val - qty_to_schedule >= min_capacity_remain):
                        dev_qty1 = qty_to_schedule
                        date1 = current_day_for_scheduling
                        final_scheduled_qty = qty_to_schedule
                        actual_prod_end_date = date1
                        if current_day_for_scheduling in live_capacity_df.index:
                            live_capacity_df.loc[current_day_for_scheduling, 'CAPACITY REMAIN'] -= qty_to_schedule
                        else: 
                            live_capacity_df.loc[current_day_for_scheduling] = -qty_to_schedule
                        scheduled_this_po = True
                        break
                else: # Orders >= 1,000 yards
                    if schedulable_on_day1 >= qty_to_schedule and (cap_day1_val - qty_to_schedule >= min_capacity_remain):
                        dev_qty1 = qty_to_schedule
                        date1 = current_day_for_scheduling
                        final_scheduled_qty = qty_to_schedule
                        actual_prod_end_date = date1
                        if current_day_for_scheduling in live_capacity_df.index:
                            live_capacity_df.loc[current_day_for_scheduling, 'CAPACITY REMAIN'] -= qty_to_schedule
                        else:
                            live_capacity_df.loc[current_day_for_scheduling] = -qty_to_schedule
                        scheduled_this_po = True
                        break
                    else: # Try to split
                        day2_for_scheduling = current_day_for_scheduling + pd.Timedelta(days=1)
                        if day2_for_scheduling > search_limit_date:
                            current_day_for_scheduling += pd.Timedelta(days=1)
                            continue

                        cap_day2_val = live_capacity_df.loc[live_capacity_df.index == day2_for_scheduling, 'CAPACITY REMAIN'].iloc[0] if day2_for_scheduling in live_capacity_df.index else 0
                        schedulable_on_day2 = max(0, cap_day2_val + capacity_tolerance)
                        split_qty1 = round(qty_to_schedule / 2)
                        split_qty2 = qty_to_schedule - split_qty1

                        if schedulable_on_day1 >= split_qty1 and schedulable_on_day2 >= split_qty2 and \
                           (cap_day1_val - split_qty1 >= min_capacity_remain) and (cap_day2_val - split_qty2 >= min_capacity_remain):
                            dev_qty1, date1 = split_qty1, current_day_for_scheduling
                            dev_qty2, date2 = split_qty2, day2_for_scheduling
                            final_scheduled_qty = qty_to_schedule
                            actual_prod_end_date = date2
                            if date1 in live_capacity_df.index: live_capacity_df.loc[date1, 'CAPACITY REMAIN'] -= split_qty1
                            else: live_capacity_df.loc[date1] = -split_qty1
                            if date2 in live_capacity_df.index: live_capacity_df.loc[date2, 'CAPACITY REMAIN'] -= split_qty2
                            else: live_capacity_df.loc[date2] = -split_qty2
                            scheduled_this_po = True
                            break
                
                current_day_for_scheduling += pd.Timedelta(days=1)
            
            if not scheduled_this_po:
                 actual_prod_end_date = far_future_date

        final_etd_val = far_future_date
        if actual_prod_end_date != far_future_date and pd.notna(actual_prod_end_date):
            final_etd_val = actual_prod_end_date + pd.Timedelta(days=lead_time_days)

        res = po_row.to_dict()
        res['DEVIDED QUANTITY 1ST'] = dev_qty1
        res['DATE 1ST BATCH'] = date1
        res['DEVIDED QUANTITY 2ND'] = dev_qty2
        res['DATE 2ND BATCH'] = date2
        res['FINAL QUANTITY'] = final_scheduled_qty if final_scheduled_qty > 0 else po_row['Quantity request']
        res['FINAL ETD'] = final_etd_val
        final_etd_results.append(res)

    final_etd_df = pd.DataFrame(final_etd_results)
    print("Step 3 finished.")
    return final_etd_df 