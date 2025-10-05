# pivot_sheets.py - Pivot sheet creation and reordering
import pandas as pd
from xlsxwriter.utility import xl_col_to_name

def write_entrydatesuppliers_pivot(workbook, df_out, config):
    """Create Pivot EntryDate × Supplier time-series analysis sheet"""
    
    if config.get('user_feedback', True):
        print("Creating Pivot EntryDate × Supplier analysis")
    
    sheet_name = "EntrydateSuppliers Pivot"
    # Get unique suppliers from Combined Data, sorted by occurrence count (descending)
    supplier_col = 'supplier_bu'
    supplier_counts = df_out[supplier_col].value_counts()
    suppliers = list(supplier_counts.index)
    num_suppliers = len(suppliers)

    # Prepare headers: Entry Date↓, Total RIDs Reported, then supplier names
    headers = ["Entry Date↓", "Total RIDs Reported"] + suppliers

    pivot_ws = workbook.add_worksheet(sheet_name)
    if config.get('debug'):
        print(f"[DEBUG] Worksheet '{sheet_name}' created.")

    # Write headers
    for col_idx, header in enumerate(headers):
        if config.get('debug'):
            print(f"[DEBUG] Writing header '{header}' at column {col_idx}")
        header_fmt = workbook.add_format({
            'bold': True,
            'align': 'left' if col_idx == 0 else 'right',
            'border': 1,
            'text_wrap': True  # Enable word wrap for all headers
        })
        pivot_ws.write(0, col_idx, header, header_fmt)
        pivot_ws.set_column(col_idx, col_idx, 11)  # Set all columns width to 11

    # Calculate entrydate_split values in Python for pivot sheet (do not update df_out)
    entrydate_col = 'entrydate'
    entrydates_raw = df_out[entrydate_col].dropna().astype(str)
    entrydates_clean = entrydates_raw[~entrydates_raw.isin(['NaT', 'nan', '', None])]
    entrydates_dt = pd.to_datetime(entrydates_clean, errors='coerce').dropna().sort_values()
    entrydates = entrydates_dt.dt.strftime('%Y-%m-%d').unique()
    num_entrydates = len(entrydates)
    if config.get('debug'):
        print(f"[DEBUG] EntrydateSuppliers Pivot: Calculated entrydates from entrydate column: {entrydates.tolist()} (total: {num_entrydates})")

    # Write entry date names in column A (starting from row 2) with border
    for i, entrydate in enumerate(entrydates):
        safe_entrydate = entrydate if pd.notna(entrydate) else ""
        if config.get('debug'):
            print(f"[DEBUG] Writing entrydate '{safe_entrydate}' at row {i+2}")
        pivot_ws.write(i+1, 0, safe_entrydate, workbook.add_format({'bold': True, 'align': 'left', 'border': 1}))

    # Place formulas for grid (columns B onward), all with border
    for i in range(num_entrydates):
        excel_row = i + 2
        entrydate_cell = f'$A{excel_row}'
        # Column B: Total RIDs Reported
        pivot_ws.write_formula(
            excel_row-1, 1,
            f"=COUNTIF('Combined Data'!$BT:$BT,{entrydate_cell})",
            workbook.add_format({'bold': True, 'align': 'right', 'border': 1})
        )
        # Columns C onward: supplier columns
        for j, supplier in enumerate(suppliers, start=2):
            supplier_cell = f"{xl_col_to_name(j)}$1"
            # Count RIDs for entrydate and supplier
            formula = f"=COUNTIFS('Combined Data'!$BT:$BT,{entrydate_cell},'Combined Data'!$AS:$AS,{supplier_cell})"
            pivot_ws.write_formula(
                excel_row-1, j,
                formula,
                workbook.add_format({'align': 'right', 'border': 1})
            )

    # Add "Total" row after entrydates, all cells with border
    total_row = num_entrydates + 2
    if config.get('debug'):
        print(f"[DEBUG] Writing 'Total' row at {total_row}")
    pivot_ws.write(total_row-1, 0, "Total", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        col_letter = xl_col_to_name(col)
        pivot_ws.write_formula(
            total_row-1, col,
            f"=SUM({col_letter}2:{col_letter}{total_row-1})",
            workbook.add_format({'bold': True, 'border': 1})
        )

    # Add "%" row after Total row, all cells with border
    percent_row = total_row + 1
    if config.get('debug'):
        print(f"[DEBUG] Writing '%' row at {percent_row}")
    pivot_ws.write(percent_row-1, 0, "%", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        col_letter = xl_col_to_name(col)
        pivot_ws.write_formula(
            percent_row-1, col,
            f"=IF($B{total_row}>0,{col_letter}{total_row}/$B{total_row},\"\")",
            workbook.add_format({'bold': True, 'num_format': '0.00%', 'align': 'right', 'border': 1})
        )

    # Conditional formatting: Total row, columns B onward, blue data bar gradient fill (same as EntrydateFlags Pivot)
    pivot_ws.conditional_format(
        f'B{total_row}:{xl_col_to_name(len(headers)-1)}{total_row}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )

    # Conditional formatting: Column B, rows 2 to Total row (including Total), blue data bar
    pivot_ws.conditional_format(
        f'B2:B{total_row}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )
    # After writing all data rows, before totals/% rows
    # Conditional formatting: Yellow to white scale (C2:last supplier column[data_row_end])
    data_row_end = num_entrydates + 1
    max_value = len(df_out)
    last_supplier_col = len(headers) - 1
    if config.get('debug'):
        print(f"[DEBUG] EntrydateSuppliers Pivot conditional formatting max_value: {max_value}")
    # Yellow to white scale (C2:last supplier column[data_row_end])
    pivot_ws.conditional_format(
        f'C2:{xl_col_to_name(last_supplier_col)}{data_row_end}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#FFFF00',
            'max_type': 'num',
            'min_value': 0,
            'max_value': max_value
        }
    )

    # Optionally, freeze top row and first column
    pivot_ws.freeze_panes(1, 1)

    if config.get('debug'):
        print("[DEBUG] Finished EntrydateSuppliers Pivot sheet.")