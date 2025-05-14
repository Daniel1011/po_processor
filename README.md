# Fabric Stock Management and ETD Calculation

## Overview

This project processes Purchase Order (PO) data, fabric stock information, first lot status, and production capacity to calculate and schedule Estimated Time of Delivery (ETD) for fabric orders. It automates several steps, from initial data loading and cleaning to final ETD calculation and output generation in an Excel file.

A key aspect of the recent updates is the standardization of column names:

* `DSM Code` is now referred to as `Greige Code`.
* `CPT Name` is now referred to as `Greige Name`.
Furthermore, `Greige Code` is consistently treated as a **Text/String** data type throughout the process.

## Requirements

* Python 3.x
* Pandas library (`pip install pandas openpyxl`)

## Configuration (`config.py`)

The `config.py` file centralizes key parameters for the application:

* `INPUT_EXCEL_FILE`: Path to the input Excel file (default: `PO - Request.xlsx`).
* `OUTPUT_EXCEL_FILE`: Path for the output Excel file (default: `PO - Processed Output.xlsx`).
* `TODAY_DATE`: The reference date for calculations (can be set to a specific date or `datetime.date.today()`).
* `LEAD_TIME_DAYS`: Standard lead time in days.
* `FAR_FUTURE_DATE`: A placeholder date for items that cannot be scheduled.
* `OCD_COL_NAME`: Name of the Original Confirmation Date column.
* `CAPACITY_TOLERANCE`, `MIN_CAPACITY_REMAIN`: Parameters for production scheduling.

## Input Data Format

The script expects an Excel file (e.g., `PO - Request.xlsx`) with the following sheets:

1. **`Stock` Sheet:** Contains current fabric stock levels.
    * Key Columns:
        * `Greige Code` (or `Mã Vải`, `DSM Code` which will be renamed) - *Text type*
        * `Greige ETA` (or `ETA`)
        * `Greige Incoming` (or `Available`)
        * Other relevant stock details.

2. **`PO` Sheet:** Contains purchase order details.
    * Key Columns:
        * `PO`
        * `Greige Code` (or `Mã Vải`, `DSM Code`) - *Text type*
        * `Greige Name` (or `CPT Name`) - *Text type*
        * `CHD` (Confirmed Handover Date)
        * `Quantity request`
        * `Forecasted`
        * `ITEM`
        * `COLOR`
        * `OCD` (Original Confirmation Date - column name configurable in `config.py`)
        * Other PO details like `SPL`, `FG name`, `Season`, `Local/ Export`.

3. **`1ST LOT STATUS` Sheet:** Contains information about the status of the first lot of production.
    * Key Columns:
        * `Greige Code` (or `DSM Code`) - *Text type*
        * `COLOR`
        * `STATUS` (e.g., "OK", "PENDING")
        * `DUE DATE`

4. **`Capacity Status` Sheet:** Details daily or weekly production capacity.
    * Structure depends on how `step3_final_etd.py` consumes it (e.g., Date, Capacity).

**Important Note:** The script is designed to handle legacy column names like `DSM Code`, `Mã Vải`, and `CPT Name` in the input file, automatically renaming them to `Greige Code` and `Greige Name` respectively. `Greige Code` will always be processed as a text/string data type.

## Processing Logic

The main script `po_processor.py` orchestrates the following steps:

### Step 0: Data Loading and Preparation (`data_loader.py`)

* Loads data from the specified sheets in the input Excel file.
* **Column Renaming:**
  * Dynamically renames columns like `Mã Vải` or `DSM Code` to `Greige Code`.
  * Renames `CPT Name` to `Greige Name`.
* **Data Type Conversion:**
  * Ensures `Greige Code` is treated as a **string** data type across all relevant DataFrames.
  * Converts date columns to datetime objects and numeric columns to appropriate numeric types.
* Performs data cleaning (e.g., stripping whitespace, handling missing values) and validation.

### Step 1: Draft ETD Calculation & Remaining Stock (`step1_draft_etd.py`)

* Takes the cleaned stock and PO data as input.
* Calculates an initial "Draft ETD" for each PO line item based on available on-hand stock and incoming stock batches.
* Considers the `LEAD_TIME_DAYS` from `config.py`.
* Generates a `remaining_stock_df` showing stock levels after allocating to POs for the Draft ETD.

### Step 2: Second ETD Calculation (`step2_second_etd.py`)

* Takes the `draft_etd_df` (output from Step 1) and the `first_lot_df` as input.
* Merges the Draft ETD data with the 1st Lot Status information using `Greige Code` (as string) and `COLOR` as keys.
* Calculates a "2nd ETD" by potentially adjusting the Draft ETD based on the `STATUS` and `DUE DATE` from the 1st Lot information. If the 1st lot is delayed, the 2nd ETD might be pushed out.

### Step 3: Production Scheduling and Final ETD (`step3_final_etd.py`)

* Takes the `draft_etd_df_with_2nd_etd` (output from Step 2) and `capacity_status_df` as input.
* Schedules production for POs that require manufacturing (i.e., not fully covered by existing stock).
* Considers daily/weekly production capacities, `LEAD_TIME_DAYS`, `CAPACITY_TOLERANCE`, and `MIN_CAPACITY_REMAIN` from `config.py`.
* Calculates the "Final ETD" based on the production schedule. This step may involve splitting PO quantities across multiple production batches if capacity is limited.

### Step 4: Output Generation (`excel_writer.py`)

* Takes the results from the previous steps:
  * `draft_etd_df_with_2nd_etd` (containing Draft ETD, 1st Lot Status, 2nd ETD)
  * `remaining_stock_df`
  * `final_etd_df`
* Writes these DataFrames to different sheets in a new Excel file (specified by `OUTPUT_EXCEL_FILE` in `config.py`).
* Formats the output for readability.

## How to Run

1. **Ensure Prerequisites:**
    * Install Python 3.
    * Install the pandas library: `pip install pandas openpyxl`
2. **Prepare Input File:**
    * Ensure your input Excel file (e.g., `PO - Request.xlsx`) is present and structured according to the "Input Data Format" section.
    * Place it in the same directory as the scripts, or update `INPUT_EXCEL_FILE` in `config.py`.
3. **Run the Script:**
    * Open a terminal or command prompt.
    * Navigate to the directory containing the scripts.
    * Execute the main processor: `python3 po_processor.py`

## Output

The script will generate an Excel file (e.g., `PO - Processed Output.xlsx` by default) containing the following sheets:

* **`Draft_ETD_and_2nd_ETD`:** Detailed information for each PO line, including `Greige Code`, `Greige Name`, original PO details, calculated `Draft ETD`, `1ST LOT STATUS`, `DUE DATE` (from 1st Lot), and `2nd ETD`.
* **`Remaining_Stock`:** Shows the `Greige Code`, `Greige Name`, remaining on-hand stock, and remaining incoming batches after initial allocations.
* **`Final_ETD_Schedule`:** The final production schedule and ETDs for all POs, including details on split batches if applicable.
