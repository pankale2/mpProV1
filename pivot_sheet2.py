# pivot_sheets.py - Pivot sheet creation and reordering
import pandas as pd
from xlsxwriter.utility import xl_col_to_name

def write_prioflag_pivot(workbook, df_out, config):
    if config.get('debug'):
        print("[DEBUG] Creating PrioFlag Pivot sheet")
    
    # Add a new worksheet for the pivot table
    pivot_ws = workbook.add_worksheet('PrioFlag Pivot')

    # Hard-coded column headers
    headers = [
        "Supplier ↓",
        "Total RIDs Reported",
        "No Flags",
        "Recent User, No Enough Data",
        "New User, First survey, Bot suspect",
        "High Reversal Rate",
        "High Security Terms Rate",
        "Low Conversion Rate",
        "High LOI, Distracted",
        "Low LOI, Speeder",
        "Total Flagged RIDs"
    ]
    # Define formats
    fmt_bold = workbook.add_format({'bold': True})
    fmt_bold_right = workbook.add_format({'bold': True, 'align': 'right'})
    fmt_right = workbook.add_format({'align': 'right'})
    fmt_wrap_right = workbook.add_format({'align': 'right', 'text_wrap': True})
    fmt_bold_wrap_right = workbook.add_format({'bold': True, 'align': 'right', 'text_wrap': True})
    fmt_bold_left = workbook.add_format({'bold': True, 'align': 'left'})
    fmt_percent = workbook.add_format({'num_format': '0.00%'})
    fmt_bold_percent = workbook.add_format({'bold': True, 'num_format': '0.00%', 'align': 'right'})
    # Background colors
    bg_colors = [
        None,           # A
        '#DCE6F1',      # B
        '#D8E4BC',      # C
        '#EBF1DE',      # D
        '#F2DCDB',      # E
        '#F2DCDB',      # F
        '#F2DCDB',      # G
        '#F2DCDB',      # H
        '#F2DCDB',      # I
        '#F2DCDB',      # J
        '#E6B8B7'       # K
    ]
    # Define formats for formula cells with background color
    fmt_bold_right_bg = [
        None,  # A
        workbook.add_format({'bold': True, 'align': 'right', 'bg_color': '#DCE6F1'}),
        workbook.add_format({'align': 'right', 'bg_color': '#D8E4BC'}),
        workbook.add_format({'align': 'right', 'bg_color': '#EBF1DE'}),
        workbook.add_format({'align': 'right', 'bg_color': '#F2DCDB'}),
        workbook.add_format({'align': 'right', 'bg_color': '#F2DCDB'}),
        workbook.add_format({'align': 'right', 'bg_color': '#F2DCDB'}),
        workbook.add_format({'align': 'right', 'bg_color': '#F2DCDB'}),
        workbook.add_format({'align': 'right', 'bg_color': '#F2DCDB'}),
        workbook.add_format({'align': 'right', 'bg_color': '#F2DCDB'}),
        workbook.add_format({'bold': True, 'align': 'right', 'bg_color': '#E6B8B7'}),
    ]
    
    # Define formats for Total row with background color
    fmt_bold_bg = [
        None,  # A
        workbook.add_format({'bold': True, 'bg_color': '#DCE6F1', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#D8E4BC', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#EBF1DE', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#F2DCDB', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#F2DCDB', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#F2DCDB', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#F2DCDB', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#F2DCDB', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#F2DCDB', 'border': 1}),
        workbook.add_format({'bold': True, 'bg_color': '#E6B8B7', 'border': 1}),
    ]

    # Add border format
    fmt_border = workbook.add_format({'border': 1})

    # Write headers with formatting, background color, wordwrap, and border for columns B-K
    for col_idx, header in enumerate(headers):
        if col_idx == 0:
            header_fmt = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
            pivot_ws.write(0, col_idx, header, header_fmt)
            pivot_ws.set_column(col_idx, col_idx, 30)  # <-- Column A width set to 38
        else:
            header_fmt = workbook.add_format({
                'bold': True,
                'align': 'right',
                'text_wrap': True,
                'bg_color': bg_colors[col_idx],
                'border': 1
            })
            pivot_ws.write(0, col_idx, header, header_fmt)
            pivot_ws.set_column(col_idx, col_idx, 11, fmt_bold_right_bg[col_idx])  # <-- Columns B-K width set to 14

    # Get unique suppliers from Combined Data, sorted by occurrence count (descending)
    supplier_col = 'supplier_bu'
    supplier_counts = df_out[supplier_col].value_counts()
    suppliers = list(supplier_counts.index)
    num_suppliers = len(suppliers)

    # Write supplier names in column A (starting from row 2) with border
    for i, supplier in enumerate(suppliers):
        pivot_ws.write(i+1, 0, supplier, workbook.add_format({'bold': True, 'align': 'left', 'border': 1}))

    # Place formulas for grid (columns B to J) and column K, all with border
    for i in range(num_suppliers):
        excel_row = i + 2
        supplier_cell = f'$A{excel_row}'
        # Column B: Total RIDs Reported
        pivot_ws.write_formula(excel_row-1, 1, f"=COUNTIF('Combined Data'!$AS:$AS,{supplier_cell})",
                               workbook.add_format({'bold': True, 'align': 'right', 'bg_color': '#DCE6F1', 'border': 1}))
        # Columns C to J
        for j in range(2, 10):
            header_cell = f"{xl_col_to_name(j)}$1"
            pivot_ws.write_formula(excel_row-1, j,
                f'=COUNTIFS(\'Combined Data\'!$AS:$AS,{supplier_cell},\'Combined Data\'!$BR:$BR,{header_cell})',
                workbook.add_format({'align': 'right', 'bg_color': bg_colors[j], 'border': 1}))
        # Column K
        pivot_ws.write_formula(excel_row-1, 10, f"=SUM(E{excel_row}:J{excel_row})",
                               workbook.add_format({'bold': True, 'align': 'right', 'bg_color': '#E6B8B7', 'border': 1}))

    # Add "Total" row after suppliers, all cells with border
    total_row = num_suppliers + 2
    pivot_ws.write(total_row-1, 0, "Total", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, 11):
        col_letter = xl_col_to_name(col)
        pivot_ws.write_formula(
            total_row-1, col,
            f"=SUM({col_letter}2:{col_letter}{total_row-1})",
            workbook.add_format({'bold': True, 'bg_color': bg_colors[col], 'border': 1})
        )

    # Add "%" row after Total row, all cells with border
    percent_row = total_row + 1
    pivot_ws.write(percent_row-1, 0, "%", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, 11):
        col_letter = xl_col_to_name(col)
        pivot_ws.write_formula(percent_row-1, col,
            f"=IF($B{total_row}>0,{col_letter}{total_row}/$B{total_row},\"\")",
            workbook.add_format({'bold': True, 'num_format': '0.00%', 'align': 'right', 'bg_color': bg_colors[col], 'border': 1}))

    # After writing all data rows, before totals/% rows
    data_row_end = num_suppliers + 1
    max_value = len(df_out)
    if config.get('debug'):
        print(f"[DEBUG] PrioFlag Pivot conditional formatting max_value: {max_value}")
    # Red to white scale (E2:J[data_row_end]) - flag columns
    pivot_ws.conditional_format(
        f'E2:J{data_row_end}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#FF0000',
            'max_type': 'num',
            'min_value': 0,
            'max_value': f'={max_value}'
        }
    )
    # Green to white scale (C2:D[data_row_end])
    pivot_ws.conditional_format(
        f'C2:D{data_row_end}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#006400',
            'max_type': 'num',
            'min_value': 0,
            'max_value': f'={max_value}'
        }
    )

    # Apply borders to all data cells (headers, data, Total, % rows)
    last_row = percent_row
    for row in range(0, last_row):
        for col in range(0, 11):
            cell = pivot_ws.table[row][col]
            # If cell already has a format, combine with border
            # (xlsxwriter does not support combining formats, so we ensure all writes above include border)
            # This loop is just for completeness; actual border is set in formats above.

    # Conditional formatting: Total row, columns B-K, blue data bar gradient fill
    # Remove 'gradient' parameter (not supported by xlsxwriter)
    pivot_ws.conditional_format(
        f'B{total_row}:K{total_row}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',  # Default Excel blue
            'data_bar_2010': True,
            'bar_only': False
        }
    )

    # Conditional formatting: Column B, rows 2 to Total row (including Total), blue data bar
    pivot_ws.conditional_format(
        f'B2:B{total_row}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',  # Default Excel blue
            'data_bar_2010': True,
            'bar_only': False
        }
    )

    # Optionally, freeze top row and first column
    pivot_ws.freeze_panes(1, 1)

    # After writing first table, calculate supplier percentages
    total_rids = sum([df_out['supplier_bu'].value_counts()[s] for s in suppliers])
    supplier_percent = {s: df_out['supplier_bu'].value_counts()[s] / total_rids for s in suppliers}
    # Pass max_value to summary table via config
    config['max_value'] = max_value
    config['combined_data_count'] = len(df_out)
    write_prioflag_pivot_summary(
        workbook,
        pivot_ws,
        suppliers,
        supplier_percent,
        headers,
        first_table_start_row=0,
        first_table_total_row=total_row,
        first_table_percent_row=percent_row,
        config=config
    )
    
def write_prioflag_pivot_summary(workbook, pivot_ws, suppliers, supplier_percent, headers, first_table_start_row, first_table_total_row, first_table_percent_row, config):
    """
    Write the second summary table for PrioFlag Pivot, listing only suppliers with >5% RIDs,
    clubbing others into a single "Other" row, and referencing values from the first table.
    """
    # Leave 4 blank rows after first table's % row
    summary_start_row = first_table_percent_row + 4

    # Use workbook for formats instead of pivot_ws.book
    for col_idx, header in enumerate(headers):
        if col_idx == 0:
            header_fmt = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
            pivot_ws.write(summary_start_row, col_idx, header, header_fmt)
            pivot_ws.set_column(col_idx, col_idx, 30)
        else:
            header_fmt = workbook.add_format({
                'bold': True,
                'align': 'right',
                'text_wrap': True,
                'bg_color': [
                    None, '#DCE6F1', '#D8E4BC', '#EBF1DE', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#E6B8B7'
                ][col_idx],
                'border': 1
            })
            pivot_ws.write(summary_start_row, col_idx, header, header_fmt)
            pivot_ws.set_column(col_idx, col_idx, 11)

    # Identify suppliers with >5% RIDs and those to club
    main_suppliers = [s for s in suppliers if supplier_percent[s] > 0.05]
    other_suppliers = [s for s in suppliers if supplier_percent[s] <= 0.05]
    num_other = len(other_suppliers)

    # Map supplier to its row in the first table
    supplier_row_map = {s: i+2 for i, s in enumerate(suppliers)}

    # Write supplier rows (>5% RIDs)
    for idx, s in enumerate(main_suppliers):
        row = summary_start_row + 1 + idx
        first_row = supplier_row_map[s]
        pivot_ws.write(row, 0, s, workbook.add_format({'bold': True, 'align': 'left', 'border': 1}))
        for col in range(1, len(headers)):
            cell_ref = f"{xl_col_to_name(col)}{first_row}"
            fmt = workbook.add_format({'align': 'right', 'border': 1, 'bold': col in [1,10]})
            pivot_ws.write_formula(row, col, f"={cell_ref}", fmt)

    # Write "Other" row
    other_row = summary_start_row + 1 + len(main_suppliers)
    other_label = f"Other ({num_other}) suppliers"
    pivot_ws.write(other_row, 0, other_label, workbook.add_format({'italic': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        sum_formula = "+".join([f"{xl_col_to_name(col)}{supplier_row_map[s]}" for s in other_suppliers]) if num_other > 0 else "0"
        fmt = workbook.add_format({'align': 'right', 'border': 1, 'bold': col in [1,10]})
        pivot_ws.write_formula(other_row, col, f"={sum_formula}", fmt)

    # Write "Total" row
    total_row = other_row + 1
    pivot_ws.write(total_row, 0, "Total", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        # Sum all supplier rows in second table
        col_letter = xl_col_to_name(col)
        pivot_ws.write_formula(
            total_row, col,
            f"=SUM({col_letter}{summary_start_row+1}:{col_letter}{other_row})",
            workbook.add_format({'bold': True, 'bg_color': [
                None, '#DCE6F1', '#D8E4BC', '#EBF1DE', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#E6B8B7'
            ][col], 'border': 1})
        )

    # Write "%" row
    percent_row = total_row + 1
    pivot_ws.write(percent_row, 0, "%", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        col_letter = xl_col_to_name(col)
        pivot_ws.write_formula(
            percent_row, col,
            f"=IF($B{total_row+1}>0,{col_letter}{total_row+1}/$B{total_row+1},\"\")",
            workbook.add_format({'bold': True, 'num_format': '0.00%', 'align': 'right', 'bg_color': [
                None, '#DCE6F1', '#D8E4BC', '#EBF1DE', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#E6B8B7'
            ][col], 'border': 1})
        )

    # Apply conditional formatting (same as first table)
    data_row_end = other_row
    # Use the same max_value as the first table (len of df_out, passed via config if needed)
    max_value = config.get('max_value', 1)
    # Red to white scale (E2:J[data_row_end])
    pivot_ws.conditional_format(
        f'E{summary_start_row+1}:J{data_row_end}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#FF0000',
            'max_type': 'num',
            'min_value': 0,
            'max_value': max_value
        }
    )
    # Green to white scale (C2:D[data_row_end])
    pivot_ws.conditional_format(
        f'C{summary_start_row+1}:D{data_row_end}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#006400',
            'max_type': 'num',
            'min_value': 0,
            'max_value': max_value
        }
    )
    # Data bar for Total row
    pivot_ws.conditional_format(
        f'B{total_row+1}:K{total_row+1}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )
    # Data bar for column B
    pivot_ws.conditional_format(
        f'B{summary_start_row+2}:B{total_row+1}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )