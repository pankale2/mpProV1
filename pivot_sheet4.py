# pivot_sheets.py - Pivot sheet creation and reordering
import pandas as pd
from xlsxwriter.utility import xl_col_to_name

def write_entrydateflags_pivot(workbook, df_out, config):
    if config.get('debug'):
        print("[DEBUG] Creating EntrydateFlags Pivot sheet")

    sheet_name = "EntrydateFlags Pivot"
    headers = [
        "Entry Date↓",
        "Total RIDs Reported",
        "No Flags",
        "Recent User, No Enough Data",
        "New User, First survey, Bot suspect",
        "High Reversal Rate",
        "High Security Terms Rate",
        "Low Conversion Rate",
        "High LOI, Distracted",
        "Low LOI, Speeder"
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
    if config.get('debug'):
        print(f"[DEBUG] Worksheet '{sheet_name}' created.")

    # Write headers
    for col_idx, header in enumerate(headers):
        if config.get('debug'):
            print(f"[DEBUG] Writing header '{header}' at column {col_idx}")
        if col_idx == 0:
            header_fmt = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
            pivot_ws.write(0, col_idx, header, header_fmt)
            pivot_ws.set_column(col_idx, col_idx, 12)
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

    entrydate_col = 'entrydate'
    # Calculate entrydate_split values in Python for pivot sheet (do not update df_out)
    entrydates_raw = df_out[entrydate_col].dropna().astype(str)
    entrydates_clean = entrydates_raw[~entrydates_raw.isin(['NaT', 'nan', '', None])]
    # Extract date part (YYYY-MM-DD) from ISO format
    entrydates_dt = pd.to_datetime(entrydates_clean, errors='coerce').dropna().sort_values()
    entrydates = entrydates_dt.dt.strftime('%Y-%m-%d').unique()
    num_entrydates = len(entrydates)
    if config.get('debug'):
        print(f"[DEBUG] EntrydateFlags Pivot: Calculated entrydates from entrydate column: {entrydates.tolist()} (total: {num_entrydates})")

    # Write entry date names in column A (starting from row 2) with border
    for i, entrydate in enumerate(entrydates):
        safe_entrydate = entrydate if pd.notna(entrydate) else ""
        if config.get('debug'):
            print(f"[DEBUG] Writing entrydate '{safe_entrydate}' at row {i+2}")
        pivot_ws.write(i+1, 0, safe_entrydate, workbook.add_format({'bold': True, 'align': 'left', 'border': 1}))

    # Place formulas for grid (columns B to J), all with border
    for i in range(num_entrydates):
        excel_row = i + 2
        entrydate_cell = f'$A{excel_row}'
        if config.get('debug'):
            print(f"[DEBUG] Writing formulas for entrydate row {excel_row} ({entrydates[i]})")
        # Column B: Total RIDs Reported
        pivot_ws.write_formula(excel_row-1, 1, f"=COUNTIF('Combined Data'!$BT:$BT,{entrydate_cell})",
                               workbook.add_format({'bold': True, 'align': 'right', 'bg_color': '#DCE6F1', 'border': 1}))
        # Columns C to J (flag columns)
        for j, flag_name in enumerate(list(flag_map.keys()), start=2):
            header_cell = f"{xl_col_to_name(j)}$1"
            if flag_name == "No Flags":
                formula = f"=COUNTIFS('Combined Data'!$BT:$BT,{entrydate_cell},'Combined Data'!$BR:$BR,C$1)"
                pivot_ws.write_formula(excel_row-1, j, formula,
                    workbook.add_format({'align': 'right', 'bg_color': bg_colors[j], 'border': 1}))
            elif flag_name == "Recent User, No Enough Data":
                formula = f"=COUNTIFS('Combined Data'!$BT:$BT,{entrydate_cell},'Combined Data'!$BR:$BR,D$1)"
                pivot_ws.write_formula(excel_row-1, j, formula,
                    workbook.add_format({'align': 'right', 'bg_color': bg_colors[j], 'border': 1}))
            else:
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
                formula = f"=COUNTIFS('Combined Data'!$BT:$BT,{entrydate_cell},'Combined Data'!${flag_col_letter}:${flag_col_letter},TRUE)"
                pivot_ws.write_formula(excel_row-1, j, formula,
                    workbook.add_format({'align': 'right', 'bg_color': bg_colors[j], 'border': 1}))

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
            workbook.add_format({'bold': True, 'bg_color': bg_colors[col], 'border': 1})
        )

    # Add "%" row after Total row, all cells with border
    percent_row = total_row + 1
    if config.get('debug'):
        print(f"[DEBUG] Writing '%' row at {percent_row}")
    pivot_ws.write(percent_row-1, 0, "%", workbook.add_format({'bold': True, 'align': 'right', 'border': 1}))
    for col in range(1, len(headers)):
        col_letter = xl_col_to_name(col)
        pivot_ws.write_formula(percent_row-1, col,
            f"=IF($B{total_row}>0,{col_letter}{total_row}/$B{total_row},\"\")",
            workbook.add_format({'bold': True, 'num_format': '0.00%', 'align': 'right', 'bg_color': bg_colors[col], 'border': 1}))

    # Conditional formatting: Total row, columns B-J, blue data bar gradient fill
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
    # Conditional formatting: Red to white scale for E2:J[data_row_end]
    data_row_end = num_entrydates + 1
    max_value = len(df_out)
    if config.get('debug'):
        print(f"[DEBUG] EntrydateFlags Pivot conditional formatting max_value: {max_value}")
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

    if config.get('debug'):
        print("[DEBUG] Finished EntrydateFlags Pivot sheet.")
