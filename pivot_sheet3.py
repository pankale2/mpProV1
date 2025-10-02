# pivot_sheets.py - Pivot sheet creation and reordering
import pandas as pd
from xlsxwriter.utility import xl_col_to_name


def write_multiflag_pivot(workbook, df_out, config):
    if config.get('debug'):
        print("[DEBUG] Creating MultiFlag Pivot sheet")

    sheet_name = "MultiFlag Pivot"
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
    flag_map = {
        "No Flags": None,
        "Recent User, No Enough Data": "No_Enough_Data",
        "New User, First survey, Bot suspect": "New_User_Bot",
        "High Reversal Rate": "High_RR",
        "High Security Terms Rate": "High_Security",
        "Low Conversion Rate": "Poor_Conv_Rate",
        "High LOI, Distracted": "High_LOI",
        "Low LOI, Speeder": "Speeder"
    }
    # Formatting (match PrioFlag Pivot)
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
    fmt_bold_right_bg = [
        None,
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
    fmt_bold_bg = [
        None,
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

    pivot_ws = workbook.add_worksheet(sheet_name)

    # Write headers
    for col_idx, header in enumerate(headers):
        if col_idx == 0:
            header_fmt = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
            pivot_ws.write(0, col_idx, header, header_fmt)
            pivot_ws.set_column(col_idx, col_idx, 30)
        else:
            header_fmt = workbook.add_format({
                'bold': True,
                'align': 'right',
                'text_wrap': True,
                'bg_color': bg_colors[col_idx],
                'border': 1
            })
            pivot_ws.write(0, col_idx, header, header_fmt)
            pivot_ws.set_column(col_idx, col_idx, 11, fmt_bold_right_bg[col_idx])

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
        # Columns C to J (flag columns)
        for j, flag_name in enumerate(list(flag_map.keys()), start=2):
            header_cell = f"{xl_col_to_name(j)}$1"
            if flag_name == "No Flags":
                # Use same formula as PrioFlag Pivot for No Flags
                # =COUNTIFS('Combined Data'!$AS:$AS,$A2,'Combined Data'!$BR:$BR,C$1)
                formula = f"=COUNTIFS('Combined Data'!$AS:$AS,{supplier_cell},'Combined Data'!$BR:$BR,C$1)"
                pivot_ws.write_formula(excel_row-1, j, formula,
                    workbook.add_format({'align': 'right', 'bg_color': bg_colors[j], 'border': 1}))
            elif flag_name == "Recent User, No Enough Data":
                # Use same formula as PrioFlag Pivot for Recent User, No Enough Data
                # =COUNTIFS('Combined Data'!$AS:$AS,$A2,'Combined Data'!$BR:$BR,D$1)
                formula = f"=COUNTIFS('Combined Data'!$AS:$AS,{supplier_cell},'Combined Data'!$BR:$BR,D$1)"
                pivot_ws.write_formula(excel_row-1, j, formula,
                    workbook.add_format({'align': 'right', 'bg_color': bg_colors[j], 'border': 1}))
            else:
                # Multi-flag logic: count RIDs for supplier where flag is TRUE
                flag_col = flag_map[flag_name]
                flag_col_letter = {
                    "No_Enough_Data": "BP",
                    "New_User_Bot": "BN",
                    "High_RR": "BO",
                    "High_Security": "BM",
                    "Poor_Conv_Rate": "BL",
                    "High_LOI": "BK",
                    "Speeder": "BJ"
                }[flag_col]
                formula = f"=COUNTIFS('Combined Data'!$AS:$AS,{supplier_cell},'Combined Data'!${flag_col_letter}:${flag_col_letter},TRUE)"
                pivot_ws.write_formula(excel_row-1, j, formula,
                    workbook.add_format({'align': 'right', 'bg_color': bg_colors[j], 'border': 1}))
        # Column K: Total Flagged RIDs - lookup from PrioFlag Pivot sheet using XLOOKUP
        pivot_ws.write_formula(
            excel_row-1, 10,
            f'=XLOOKUP({supplier_cell},\'PrioFlag Pivot\'!$A:$A,\'PrioFlag Pivot\'!$K:$K,"")',
            workbook.add_format({'bold': True, 'align': 'right', 'bg_color': '#E6B8B7', 'border': 1})
        )

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
    # After writing all data rows, before totals/% rows
    data_row_end = num_suppliers + 1
    max_value = len(df_out)
    if config.get('debug'):
        print(f"[DEBUG] MultiFlag Pivot conditional formatting max_value: {max_value}")
    # Red to white scale (E2:J[data_row_end]) - flag columns
    pivot_ws.conditional_format(
        f'E2:J{data_row_end}',
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
        f'C2:D{data_row_end}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#006400',
            'max_type': 'num',
            'min_value': 0,
            'max_value': max_value
        }
    )
    # Optionally, freeze top row and first column
    pivot_ws.freeze_panes(1, 1)
    pivot_ws.freeze_panes(1, 1)

    # Calculate total RIDs and supplier percentages for summary table
    total_rids = sum([supplier_counts[s] for s in suppliers])
    supplier_percent = {s: supplier_counts[s] / total_rids for s in suppliers}

    # After writing main table, call summary table function
    write_multiflag_pivot_summary(
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

def write_multiflag_pivot_summary(workbook, pivot_ws, suppliers, supplier_percent, headers, first_table_start_row, first_table_total_row, first_table_percent_row, config):
    # Leave 4 blank rows after first table's % row
    summary_start_row = first_table_percent_row + 4

    # Write headers (same as main table)
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

    # Write "Other" row using SUM formula over adjacent block
    other_row = summary_start_row + 1 + len(main_suppliers)
    other_label = f"Other ({num_other}) suppliers"
    pivot_ws.write(other_row, 0, other_label, workbook.add_format({'italic': True, 'align': 'right', 'border': 1}))
    if num_other > 0:
        # Find first and last row for "Other" suppliers in original table
        other_indices = [supplier_row_map[s] for s in other_suppliers]
        min_row = min(other_indices)
        max_row = max(other_indices)
        for col in range(1, len(headers)):
            col_letter = xl_col_to_name(col)
            sum_formula = f"=SUM({col_letter}{min_row}:{col_letter}{max_row})"
            fmt = workbook.add_format({'align': 'right', 'border': 1, 'bold': col in [1,10]})
            pivot_ws.write_formula(other_row, col, sum_formula, fmt)
    else:
        for col in range(1, len(headers)):
            fmt = workbook.add_format({'align': 'right', 'border': 1, 'bold': col in [1,10]})
            pivot_ws.write(other_row, col, 0, fmt)

    # Write "Total" row - reference original table's Total row
    total_row = other_row + 1
    orig_total_row = first_table_total_row
    pivot_ws.write(total_row, 0, "Total", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        cell_ref = f"{xl_col_to_name(col)}{orig_total_row}"
        pivot_ws.write_formula(
            total_row, col,
            f"={cell_ref}",
            workbook.add_format({'bold': True, 'bg_color': [
                None, '#DCE6F1', '#D8E4BC', '#EBF1DE', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#E6B8B7'
            ][col], 'border': 1})
        )

    # Write "%" row - reference original table's % row
    percent_row = total_row + 1
    orig_percent_row = first_table_percent_row
    pivot_ws.write(percent_row, 0, "%", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        cell_ref = f"{xl_col_to_name(col)}{orig_percent_row}"
        pivot_ws.write_formula(
            percent_row, col,
            f"={cell_ref}",
            workbook.add_format({'bold': True, 'num_format': '0.00%', 'align': 'right', 'bg_color': [
                None, '#DCE6F1', '#D8E4BC', '#EBF1DE', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#F2DCDB', '#E6B8B7'
            ][col], 'border': 1})
        )

    # Apply conditional formatting (same as main table)
    data_row_end = other_row
    max_value = config.get('max_value', 1)
    if config.get('debug'):
        print(f"[DEBUG] MultiFlag Pivot summary table conditional formatting max_value: {max_value}")
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
    pivot_ws.conditional_format(
        f'B{total_row+1}:K{total_row+1}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )
    pivot_ws.conditional_format(
        f'B{summary_start_row+2}:B{total_row+1}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )
