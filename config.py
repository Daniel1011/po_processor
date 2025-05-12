import pandas as pd
import datetime
from datetime import date # For date.today()

# --- Configuration ---
INPUT_EXCEL_FILE = 'PO - Request.xlsx'
OUTPUT_EXCEL_FILE = 'Output_Stock_Management.xlsx'
TODAY_DATE_STR = date.today().strftime('%Y-%m-%d')
LEAD_TIME_DAYS = 40
FAR_FUTURE_DATE_STR = '2200-12-31'
CAPACITY_TOLERANCE = 2000
MIN_CAPACITY_REMAIN = -2000
OCD_COL_NAME = "OCD( Order Creation Day)" # As per latest version of po_processor.py

# Convert string dates to datetime objects
TODAY_DATE = pd.to_datetime(TODAY_DATE_STR)
FAR_FUTURE_DATE = pd.to_datetime(FAR_FUTURE_DATE_STR)
FAR_FUTURE_DATE_DISPLAY_STR = 'Insufficient Stock/Capacity' # Still useful for potential display choices

# --- Helper Functions ---
def excel_date_to_datetime(excel_date):
    """Converts Excel serial date to datetime object, handles existing datetime objects or NaT."""
    if pd.isna(excel_date):
        return pd.NaT
    if isinstance(excel_date, (datetime.datetime, datetime.date)):
        return pd.to_datetime(excel_date)
    try:
        if isinstance(excel_date, (int, float)):
            return pd.to_datetime('1899-12-30') + pd.to_timedelta(excel_date, 'D')
    except ValueError:
        pass # Fallback to direct conversion
    return pd.to_datetime(excel_date, errors='coerce') 