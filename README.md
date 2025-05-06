# PO Processor Script

This Python script processes purchase orders (POs) against fabric stock levels and production capacity to determine an Estimated Time of Delivery (ETD) for each PO. It reads data from an Excel file, performs calculations based on defined logic, and outputs the results to a new Excel file.

## Features

*   Reads stock availability and PO requests from an Excel file.
*   Prioritizes POs based on 'Forecasted' status and Contractual Handover Date (CHD).
*   Allocates available stock and incoming stock batches to POs.
*   Schedules PO production based on available stock date and daily production capacity.
*   Calculates ETD based on the scheduled production start date plus a lead time.
*   Handles insufficient stock and production capacity limitations.
*   Outputs ETD results and final remaining stock levels to separate sheets in an Excel file.

## Setup

1.  **Python:** Ensure you have Python 3 installed.
2.  **Libraries:** Install the required libraries (`pandas` and `openpyxl`) using pip:
    ```bash
    python3 -m pip install pandas openpyxl
    ```

## Input File (`input.xlsx`)

The script expects an Excel file named `input.xlsx` in the same directory. This file must contain two sheets: `Stock` and `PO`.

**Sheet: `Stock`**

| Column        | Description                    | Data Type     | Example    |
| ------------- | ------------------------------ | ------------- | ---------- |
| `Mã Vải`      | Fabric Code                    | Text          | A          |
| `ETA`         | Estimated Time of Arrival      | Date/Datetime | 2025-05-06 |
| `Available`   | Quantity in the stock batch    | Number        | 2000       |

**Sheet: `PO`**

| Column             | Description                   | Data Type     | Example    |
| ------------------ | ----------------------------- | ------------- | ---------- |
| `PO`               | Purchase Order Identifier     | Text          | PO1 A      |
| `Mã Vải`           | Fabric Code                   | Text          | A          |
| `SPL`              | SPL Identifier (not used in ETD calculation but required in input) | Text          | X          |
| `CHD`              | Contractual Handover Date     | Date/Datetime | 2025-08-02 |
| `Quantity request` | Quantity required for the PO  | Number        | 500        |
| `Forecasted`       | Is the PO forecasted? ('yes'/'no', case-insensitive) | Text          | yes        |

*Note: Dates should be in a format pandas can parse (e.g., YYYY-MM-DD, MM/DD/YYYY).*

## Configuration (in `po_processor.py`)

Key parameters can be adjusted directly in the script:

*   `INPUT_EXCEL_FILE`: Name of the input file (default: `'input.xlsx'`).
*   `OUTPUT_EXCEL_FILE`: Name of the output file (default: `'output.xlsx'`).
*   `TODAY_DATE_STR`: The reference date for calculations (default: `'2025-05-06'`).
*   `MONTHLY_CAPACITY`: Assumed total production units per month (default: `10000`). Daily capacity is derived from this.
*   `LEAD_TIME_DAYS`: Days added to production start date to get ETD (default: `40`).
*   `FAR_FUTURE_DATE_STR`: Placeholder date for unschedulable POs (default: `'2200-12-31'`).

## Usage

1.  Prepare your `input.xlsx` file according to the format specified above.
2.  Place the `input.xlsx` file in the same directory as `po_processor.py`.
3.  Run the script from your terminal:
    ```bash
    python3 po_processor.py
    ```
4.  The script will print progress messages, including warnings for insufficient stock or capacity issues.
5.  An `output.xlsx` file will be generated in the same directory.

*Note: If `input.xlsx` is not found, the script will create a dummy version with sample data for testing purposes.*

## Output File (`output.xlsx`)

The output Excel file contains two sheets:

**Sheet: `ETD Results`**

| Column | Description                                                                 | Example                      |
| ------ | --------------------------------------------------------------------------- | ---------------------------- |
| `PO`   | Purchase Order Identifier                                                   | PO1 A                        |
| `ETD`  | Calculated Estimated Time of Delivery (YYYY-MM-DD) or status message        | 2025-06-15                   |
|        | If ETD cannot be calculated due to lack of stock or capacity, shows:        | `Insufficient Stock/Capacity` |

**Sheet: `Remaining Stock`**

| Column                                                    | Description                                                            | Example                                          |
| --------------------------------------------------------- | ---------------------------------------------------------------------- | ------------------------------------------------ |
| `Mã Vải`                                                  | Fabric Code                                                            | A                                                |
| `Remaining Available (as of YYYY-MM-DD)`                  | Stock quantity with ETA <= TODAY_DATE that was *not* consumed        | 1500                                             |
| `Remaining Incoming Batches`                              | List of future stock batches (ETA > TODAY_DATE) not fully consumed | `ETA: 2025-06-06, Qty: 3000; ETA: 2025-07-11, Qty: 300` |
|                                                           | If no incoming batches remain for the fabric, shows:                   | `None`                                           |

## Assumptions

*   Monthly production capacity is constant (10,000 units by default) and distributed evenly across the days of the month.
*   The 'Forecasted' status is indicated by 'yes' or 'no' in the `PO` sheet.
*   Sufficient stock exists for *most* POs, but the script handles shortages by assigning a far-future date/status message.
*   Production capacity is a single pool shared across all fabric codes.
*   The lead time from production start to delivery is fixed (40 days by default). 