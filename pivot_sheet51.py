import pandas as pd
from xlsxwriter.utility import xl_col_to_name

def write_tenure_supplier_pivot(workbook, df_out, config):
    """
    Create TenureSupplier Pivot sheet: suppliers × tenure groups analysis.
    """
    ws = workbook.add_worksheet('TenureSupplier Pivot')

    # Define headers and column widths
    headers = [
        "Supplier ↓",
        "Less than a week",
        "Less than a month",
        "Less than 3 months",
        "Less than 6 months",
        "Less than a year",
        "More than a year",
        "Total"
    ]
    col_widths = [30] + [10] * (len(headers) - 1)

    # Formats
    fmt_bold_left = workbook.add_format({'bold': True, 'align': 'left', 'border': 1, 'text_wrap': True})
    fmt_bold_right = workbook.add_format({'bold': True, 'align': 'right', 'border': 1, 'text_wrap': True})
    fmt_right = workbook.add_format({'align': 'right', 'border': 1})
    fmt_bold_total = workbook.add_format({'bold': True, 'align': 'right', 'border': 1})
    fmt_percent = workbook.add_format({'num_format': '0.00%', 'align': 'right', 'border': 1})
    fmt_italic = workbook.add_format({'italic': True, 'align': 'right', 'border': 1})

    # Write headers and set column widths (word wrap enabled)
    for col_idx, (header, width) in enumerate(zip(headers, col_widths)):
        fmt = fmt_bold_left if col_idx == 0 else fmt_bold_right
        ws.write(0, col_idx, header, fmt)
        ws.set_column(col_idx, col_idx, width)

    # Get unique suppliers sorted by count
    supplier_col = 'supplier_bu'
    tenure_col = 'Tenure_Group'
    supplier_counts = df_out[supplier_col].value_counts()
    suppliers = list(supplier_counts.index)
    num_suppliers = len(suppliers)
    total_rids = supplier_counts.sum()

    # Tenure group values (must match headers B-G)
    tenure_groups = [
        "Less than a week",
        "Less than a month",
        "Less than 3 months",
        "Less than 6 months",
        "Less than a year",
        "More than a year"
    ]

    # Write supplier names in column A
    for i, supplier in enumerate(suppliers):
        ws.write(i + 1, 0, supplier, fmt_bold_left)

    # Write formulas for counts per supplier × tenure group
    for i in range(num_suppliers):
        excel_row = i + 2
        supplier_cell = f'$A{excel_row}'
        # Columns B-G: tenure group counts
        for j, tenure in enumerate(tenure_groups, start=1):
            header_cell = f"{xl_col_to_name(j)}$1"
            ws.write_formula(
                excel_row - 1, j,
                f'=COUNTIFS(\'Combined Data\'!$AS:$AS,{supplier_cell},\'Combined Data\'!$BS:$BS,{header_cell})',
                fmt_right
            )
        # Column H: Total (sum of B-G)
        ws.write_formula(
            excel_row - 1, 7,
            f'=SUM(B{excel_row}:G{excel_row})',
            fmt_bold_total
        )

    # Add "Total" row
    total_row = num_suppliers + 2
    ws.write(total_row - 1, 0, "Total", fmt_bold_total)
    for col in range(1, 8):
        col_letter = xl_col_to_name(col)
        ws.write_formula(
            total_row - 1, col,
            f'=SUM({col_letter}2:{col_letter}{total_row-1})',
            fmt_bold_total
        )

    # Add "%" row
    percent_row = total_row + 1
    ws.write(percent_row - 1, 0, "%", fmt_bold_total)
    for col in range(1, 8):
        col_letter = xl_col_to_name(col)
        ws.write_formula(
            percent_row - 1, col,
            f'=IF($H{total_row}>0,{col_letter}{total_row}/$H{total_row},"")',
            fmt_percent
        )

    # Freeze top row and first column
    ws.freeze_panes(1, 1)

    # Prepare summary table (main suppliers >5% RIDs, others clubbed)
    supplier_percent = {s: supplier_counts[s] / total_rids for s in suppliers}
    main_suppliers = [s for s in suppliers if supplier_percent[s] > 0.05]
    other_suppliers = [s for s in suppliers if supplier_percent[s] <= 0.05]
    num_other = len(other_suppliers)
    supplier_row_map = {s: i + 2 for i, s in enumerate(suppliers)}

    # Leave 4 blank rows after % row
    summary_start_row = percent_row + 4

    # Write summary headers (word wrap enabled)
    for col_idx, (header, width) in enumerate(zip(headers, col_widths)):
        fmt = fmt_bold_left if col_idx == 0 else fmt_bold_right
        ws.write(summary_start_row, col_idx, header, fmt)
        ws.set_column(col_idx, col_idx, width)

    # Write main suppliers
    for idx, s in enumerate(main_suppliers):
        row = summary_start_row + 1 + idx
        first_row = supplier_row_map[s]
        ws.write(row, 0, s, fmt_bold_left)
        for col in range(1, 8):
            cell_ref = f"{xl_col_to_name(col)}{first_row}"
            ws.write_formula(row, col, f"={cell_ref}", fmt_right)

    # Write "Other" row
    other_row = summary_start_row + 1 + len(main_suppliers)
    other_label = f"Other ({num_other}) suppliers"
    ws.write(other_row, 0, other_label, fmt_italic)
    for col in range(1, 8):
        sum_formula = "+".join([f"{xl_col_to_name(col)}{supplier_row_map[s]}" for s in other_suppliers]) if num_other > 0 else "0"
        ws.write_formula(other_row, col, f"={sum_formula}", fmt_right)

    # Write "Total" row
    total_row2 = other_row + 1
    ws.write(total_row2, 0, "Total", fmt_bold_total)
    for col in range(1, 8):
        main_table_total_cell = f"{xl_col_to_name(col)}{total_row}"
        ws.write_formula(total_row2, col, f"={main_table_total_cell}", fmt_bold_total)

    # Write "%" row
    percent_row2 = total_row2 + 1
    ws.write(percent_row2, 0, "%", fmt_bold_total)
    for col in range(1, 8):
        col_letter = xl_col_to_name(col)
        ws.write_formula(
            percent_row2, col,
            f'=IF($H{total_row2+1}>0,{col_letter}{total_row2+1}/$H{total_row2+1},"")',
            fmt_percent
        )

    # OPTIMIZED CONDITIONAL FORMATTING
    # Main table - 2-color scale for tenure columns (B-G), supplier rows only, excluding Total and % rows
    data_row_end = total_row - 1  # Exclude Total row from 2-color scale
    ws.conditional_format(
        f'B2:G{data_row_end}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#FFFF00',
            'max_type': 'num',
            'min_value': 0,
            'max_value': total_rids
        }
    )

    # Main table - Data bar for Total column (H), including Total row, excluding % row
    ws.conditional_format(
        f'H2:H{total_row}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )

    # Main table - Data bar for Total row across tenure columns (B-G)
    ws.conditional_format(
        f'B{total_row}:G{total_row}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )

    # Summary table - 2-color scale for tenure columns (B-G), supplier rows only, excluding Total and % rows
    summary_data_row_start = summary_start_row + 1
    summary_data_row_end = total_row2 - 1  # Exclude Total row from 2-color scale
    ws.conditional_format(
        f'B{summary_data_row_start+1}:G{summary_data_row_end+1}',
        {
            'type': '2_color_scale',
            'min_color': '#FFFFFF',
            'max_color': '#FFFF00',
            'max_type': 'num',
            'min_value': 0,
            'max_value': total_rids
        }
    )

    # Summary table - Data bar for Total column (H), including Total row, excluding % row
    ws.conditional_format(
        f'H{summary_data_row_start+1}:H{total_row2+1}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )

    # Summary table - Data bar for Total row across tenure columns (B-G)
    ws.conditional_format(
        f'B{total_row2+1}:G{total_row2+1}',
        {
            'type': 'data_bar',
            'bar_color': '#5B9BD5',
            'data_bar_2010': True,
            'bar_only': False
        }
    )