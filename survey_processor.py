# survey_processor.py
import pandas as pd
from datetime import datetime
import os
from openpyxl.styles import PatternFill, Font
from openpyxl.formatting.rule import ColorScaleRule

# openpyxl is used by pandas for writing .xlsx files, ensure it's in requirements.txt

def read_rid_file_from_stream(file_stream):
    """Reads the RID lookup CSV file from a file stream."""
    df = pd.read_csv(file_stream)
    df.columns = df.columns.str.strip().str.lower()
    return df

def read_metrics_file_from_stream(file_stream):
    """Reads the Marketplace Metrics Excel file from a file stream."""
    df = pd.read_excel(
        file_stream,
        sheet_name='Marketplace Metrics by PID',
        skiprows=5
    )
    df.columns = df.columns.str.strip().str.lower()
    if 'pid' in df.columns:
        df['pid'] = df['pid'].astype(str).str.strip()
    return df

def get_formatting_ranges(header, data_rows, total_column=None, has_col_total_row=False):
    """
    Helper function to calculate formatting exclusions and last data row for conditional formatting.
    header: list of column names (Excel order, 1-based for openpyxl)
    data_rows: total number of rows written (including totals row)
    total_column: name of total column (e.g. 'Total_Flagged', 'Row Total')
    has_col_total_row: True if the last row is a column total row
    Returns: (exclude_cols, last_data_row)
    """
    exclude_cols = [1]  # Always exclude first column (index 1)
    if '-n/a-' in header:
        exclude_cols.append(header.index('-n/a-') + 1)
    if total_column and total_column in header:
        exclude_cols.append(header.index(total_column) + 1)
    # Only format up to last data row (exclude column total row if present)
    last_data_row = data_rows - 1 if has_col_total_row else data_rows
    return exclude_cols, last_data_row

def apply_conditional_formatting(worksheet, start_col, end_col, data_rows, exclude_cols=None, total_rows=None):
    """Apply conditional formatting to specified range."""
    if exclude_cols is None:
        exclude_cols = []
    color_scale_rule = ColorScaleRule(
        start_type='num',
        start_value=0,
        start_color='FFFFFF',
        end_type='num',
        end_value=total_rows,
        end_color='f82b1b'
    )
    for col in range(start_col, end_col + 1):
        if col not in exclude_cols:
            col_letter = worksheet.cell(row=1, column=col).column_letter
            cell_range = f"{col_letter}2:{col_letter}{data_rows}"
            worksheet.conditional_formatting.add(cell_range, color_scale_rule)

def apply_na_column_formatting(worksheet, header):
    """Apply dark green formatting to -n/a- column if present."""
    if '-n/a-' in header:
        na_col_idx = header.index('-n/a-') + 1  # openpyxl is 1-based
        dark_green_font = Font(color="006400")
        for row in worksheet.iter_rows(min_row=2, min_col=na_col_idx, max_col=na_col_idx, max_row=worksheet.max_row):
            for cell in row:
                cell.font = dark_green_font
        worksheet.cell(row=1, column=na_col_idx).font = dark_green_font

def format_pivot_sheet(worksheet, header, data_rows, total_column, total_rows, has_col_total_row=True):
    """Standardized pivot formatting: conditional formatting + -n/a- styling."""
    exclude_cols, last_data_row = get_formatting_ranges(
        header, data_rows, total_column=total_column, has_col_total_row=has_col_total_row
    )
    apply_conditional_formatting(
        worksheet,
        start_col=2,
        end_col=len(header),
        data_rows=last_data_row,
        exclude_cols=exclude_cols,
        total_rows=total_rows
    )
    apply_na_column_formatting(worksheet, header)

def add_check_results_pivot(writer, df_merged):
    """Add pivot table showing all check flags by supplier."""
    workbook = writer.book
    check_columns = [
        "Poor_Conv_Rate", "New_User_Bot", "High_Security",
        "Speeder", "High_LOI", "High_RR"
    ]
    
    # Create pivot for check results
    ws_pivot = workbook.create_sheet('Flags Pivot (Multi)')
    
    # Calculate counts by supplier
    supplier_stats = []
    for supplier in df_merged['supplier_bu'].unique():
        supplier_df = df_merged[df_merged['supplier_bu'] == supplier]
        
        # Count rows with any True flag
        has_any_flag = supplier_df[check_columns].any(axis=1)
        total_flagged = has_any_flag.sum()
        
        # Count rows with no flags (true -n/a- count)
        no_flags = ~has_any_flag
        na_count = no_flags.sum()
        
        # Get individual flag counts
        flag_counts = supplier_df[check_columns].sum()
        
        stats = {
            'supplier_bu': supplier,
            '-n/a-': na_count,
            **flag_counts.to_dict(),
            'Total_Flagged': total_flagged
        }
        supplier_stats.append(stats)
    
    pivot_df = pd.DataFrame(supplier_stats)
    pivot_df = pivot_df.sort_values('Total_Flagged', ascending=False)
    
    # Add column totals with proper supplier_bu value
    totals = pd.DataFrame([{
        'supplier_bu': 'Column Total',
        '-n/a-': pivot_df['-n/a-'].sum(),
        **{col: pivot_df[col].sum() for col in check_columns},
        'Total_Flagged': pivot_df['Total_Flagged'].sum()
    }])
    pivot_df = pd.concat([pivot_df, totals], ignore_index=True)
    
    # Arrange columns in desired order:
    # 1. supplier_bu first
    # 2. -n/a- second
    # 3. check columns sorted by total count
    # 4. Total_Flagged last
    check_totals = pivot_df[check_columns].sum()
    sorted_check_cols = sorted(check_columns, key=lambda x: check_totals[x], reverse=True)
    
    header = ['supplier_bu', '-n/a-'] + sorted_check_cols + ['Total_Flagged']
    
    # Write to Excel with formatting
    ws_pivot.append(header)
    
    # Write data rows
    for _, row in pivot_df.iterrows():
        ws_pivot.append([row['supplier_bu']] + [row[col] for col in header[1:]])
    
    # Set supplier_bu column width to 200 pixels
    # Convert pixels to Excel column width units (approximately)
    excel_width = 200 / 7  # Excel width units are roughly 7 pixels
    ws_pivot.column_dimensions['A'].width = excel_width

    # Use helper for exclusions and data range
    exclude_cols, last_data_row = get_formatting_ranges(
        header, len(pivot_df), total_column='Total_Flagged', has_col_total_row=True
    )
    total_rows = len(df_merged)
    apply_conditional_formatting(
        ws_pivot,
        start_col=2,
        end_col=len(header),
        data_rows=last_data_row,
        exclude_cols=exclude_cols,
        total_rows=total_rows
    )
    
    # Apply existing -n/a- column formatting
    dark_green_font = Font(color="006400")
    for row in ws_pivot.iter_rows(min_row=2, min_col=exclude_cols[1], max_col=exclude_cols[1]):
        for cell in row:
            cell.font = dark_green_font
    ws_pivot.cell(row=1, column=exclude_cols[1]).font = dark_green_font

def add_pivot_and_format(writer, df_merged):
    """
    Adds pivot tables to the Excel workbook and applies basic formatting.
    """
    workbook = writer.book
    
    # --- Combined Data Sheet Formatting (from original) ---
    ws_combined = writer.sheets.get('Combined Data') # Get the sheet by name
    if ws_combined:
        ws_combined.auto_filter.ref = ws_combined.dimensions
        ws_combined.freeze_panes = ws_combined['A2']

    # --- Create Pivot Table ---
    if 'supplier_bu' not in df_merged.columns or 'Observation' not in df_merged.columns:
        print("Warning: 'supplier_bu' or 'Observation' column not found. Skipping pivot table.")
        return

    try:
        pivot = (
            df_merged.groupby(['supplier_bu', 'Observation'])
            .size()
            .reset_index(name='Count')
            .pivot(index='supplier_bu', columns='Observation', values='Count')
            .fillna(0)
            .astype(int)
        )

        if pivot.empty:
            print("Warning: Pivot table is empty. Skipping advanced formatting and sheet creation.")
            return

        # Add row totals
        pivot['Row Total'] = pivot.sum(axis=1)

        # Add column totals
        col_totals = pivot.sum(axis=0)
        col_totals.name = 'Column Total'
        pivot = pd.concat([pivot, pd.DataFrame([col_totals], index=['Column Total'])])


        # Sort rows by row total (descending), keep total row at the end
        if 'Row Total' in pivot.columns and 'Column Total' in pivot.index:
            # Separate the 'Column Total' row
            total_row_df = pivot.loc[['Column Total']]
            pivot_data_rows = pivot.drop('Column Total')
            
            # Sort the data rows
            pivot_data_rows = pivot_data_rows.sort_values(by='Row Total', ascending=False)
            
            # Concatenate sorted data rows with the total row
            pivot = pd.concat([pivot_data_rows, total_row_df])


        # Sort columns: "-n/a-" first, then by column total (descending), then Row Total last
        cols = list(pivot.columns)
        # Use .get(c, 0) for robustness if a column name is unexpectedly missing from col_totals
        col_total_values = pivot.loc['Column Total'] if 'Column Total' in pivot.index else pivot.sum(axis=0)

        obs_cols = [c for c in cols if c not in ['Row Total']]
        sorted_obs_cols = []

        if '-n/a-' in obs_cols:
            sorted_obs_cols.append('-n/a-')
            obs_cols_no_na = [c for c in obs_cols if c != '-n/a-']
        else:
            obs_cols_no_na = list(obs_cols)
            
        # Sort remaining observation columns by their total count
        obs_cols_sorted = sorted(obs_cols_no_na, key=lambda c: col_total_values.get(c, 0), reverse=True)
        sorted_obs_cols.extend(obs_cols_sorted)

        if 'Row Total' in cols:
            sorted_obs_cols.append('Row Total')
        
        pivot = pivot[sorted_obs_cols]

        # Write pivot table to new sheet
        ws_pivot = workbook.create_sheet('Flags Pivot (Priority)')
        
        # Write header (supplier_bu and then pivot columns)
        header = ['supplier_bu'] + list(pivot.columns)
        ws_pivot.append(header)
        
        # Write data rows
        for supplier_bu_index, row_data in pivot.iterrows():
            ws_pivot.append([supplier_bu_index] + list(row_data.values))

        # Set supplier_bu column width to 200 pixels
        excel_width = 200 / 7  # Excel width units are roughly 7 pixels
        ws_pivot.column_dimensions['A'].width = excel_width

        exclude_cols, last_data_row = get_formatting_ranges(
            header, len(pivot), total_column='Row Total', has_col_total_row=True
        )
        total_rows = len(df_merged)
        apply_conditional_formatting(
            ws_pivot,
            start_col=2,
            end_col=len(header),
            data_rows=last_data_row,
            exclude_cols=exclude_cols,
            total_rows=total_rows
        )

        # --- Style "-n/a-" column in dark green ---
        dark_green_font = Font(color="006400")  # Hex for dark green

        # Find the column index for "-n/a-"
        try:
            na_col_idx = header.index('-n/a-') + 1  # openpyxl is 1-based
            for row in ws_pivot.iter_rows(min_row=2, min_col=na_col_idx, max_col=na_col_idx, max_row=ws_pivot.max_row):
                for cell in row:
                    cell.font = dark_green_font
            # Also style the header cell
            ws_pivot.cell(row=1, column=na_col_idx).font = dark_green_font
        except ValueError:
            pass  # "-n/a-" column not present

        # --- Additional Pivots ---

        # Helper: get just the date part from entrydate
        df_merged['entrydate_only'] = df_merged['entrydate'].astype(str).str.split('T').str[0]

        # 1. Pivot: entrydate_only vs supplier_bu
        if 'entrydate_only' in df_merged.columns and 'supplier_bu' in df_merged.columns:
            pivot_entrydate_supplier = (
                df_merged.groupby(['entrydate_only', 'supplier_bu'])
                .size()
                .reset_index(name='Count')
                .pivot(index='entrydate_only', columns='supplier_bu', values='Count')
                .fillna(0)
                .astype(int)
            )
            
            # Sort index by date (ascending)
            pivot_entrydate_supplier = pivot_entrydate_supplier.sort_index()
            
            # Add row totals
            pivot_entrydate_supplier['Row Total'] = pivot_entrydate_supplier.sum(axis=1)
            
            # Add column totals
            col_totals = pivot_entrydate_supplier.sum()
            col_totals.name = 'Column Total'
            pivot_entrydate_supplier = pd.concat([pivot_entrydate_supplier, pd.DataFrame([col_totals], index=['Column Total'])])
            
            # Sort columns by total count (highest first) but keep Row Total last
            cols = list(pivot_entrydate_supplier.columns)
            if 'Row Total' in cols:
                cols.remove('Row Total')
            col_totals_sorted = pivot_entrydate_supplier.loc[pivot_entrydate_supplier.index != 'Column Total', cols].sum()
            sorted_cols = col_totals_sorted.sort_values(ascending=False).index
            sorted_cols = list(sorted_cols) + ['Row Total']
            pivot_entrydate_supplier = pivot_entrydate_supplier[sorted_cols]

            ws_pivot1 = workbook.create_sheet('Pivot EntryDate x Supplier')
            header = ['entrydate'] + list(pivot_entrydate_supplier.columns)
            ws_pivot1.append(header)
            for idx, row_data in pivot_entrydate_supplier.iterrows():
                ws_pivot1.append([idx] + list(row_data.values))
            # Set all columns to 100 pixels width
            excel_width = 100 / 7  # Excel width units are roughly 7 pixels
            for col_idx in range(1, len(header) + 1):
                col_letter = ws_pivot1.cell(row=1, column=col_idx).column_letter
                ws_pivot1.column_dimensions[col_letter].width = excel_width

            format_pivot_sheet(ws_pivot1, header, len(pivot_entrydate_supplier), 'Row Total', len(df_merged))

        # 2. Pivot: entrydate_only vs Observation
        if 'entrydate_only' in df_merged.columns and 'Observation' in df_merged.columns:
            pivot_entrydate_obs = (
                df_merged.groupby(['entrydate_only', 'Observation'])
                .size()
                .reset_index(name='Count')
                .pivot(index='entrydate_only', columns='Observation', values='Count')
                .fillna(0)
                .astype(int)
            )
            
            # Sort index by date (ascending)
            pivot_entrydate_obs = pivot_entrydate_obs.sort_index()
            
            # Add row totals
            pivot_entrydate_obs['Row Total'] = pivot_entrydate_obs.sum(axis=1)
            
            # Add column totals
            col_totals = pivot_entrydate_obs.sum()
            col_totals.name = 'Column Total'
            pivot_entrydate_obs = pd.concat([pivot_entrydate_obs, pd.DataFrame([col_totals], index=['Column Total'])])
            
            # Sort columns: "-n/a-" first, then others by total (highest first), then Row Total last
            cols = list(pivot_entrydate_obs.columns)
            sorted_cols = []
            if '-n/a-' in cols:
                sorted_cols.append('-n/a-')
                cols.remove('-n/a-')
            if 'Row Total' in cols:
                cols.remove('Row Total')
            col_totals_sorted = pivot_entrydate_obs.loc[pivot_entrydate_obs.index != 'Column Total', cols].sum()
            remaining_cols = col_totals_sorted.sort_values(ascending=False).index
            sorted_cols.extend(remaining_cols)
            sorted_cols.append('Row Total')
            pivot_entrydate_obs = pivot_entrydate_obs[sorted_cols]

            ws_pivot2 = workbook.create_sheet('Pivot EntryDate x Flags')
            header = ['entrydate'] + list(pivot_entrydate_obs.columns)
            ws_pivot2.append(header)
            for idx, row in pivot_entrydate_obs.iterrows():
                ws_pivot2.append([idx] + list(row.values))
            ws_pivot2.column_dimensions['A'].width = 100 / 7

            format_pivot_sheet(ws_pivot2, header, len(pivot_entrydate_obs), 'Row Total', len(df_merged))

            # Format -n/a- column in dark green (like other pivots)
            try:
                na_col_idx = header.index('-n/a-') + 1  # openpyxl is 1-based
                dark_green_font = Font(color="006400")
                for row in ws_pivot2.iter_rows(min_row=2, min_col=na_col_idx, max_col=na_col_idx, max_row=ws_pivot2.max_row):
                    for cell in row:
                        cell.font = dark_green_font
                ws_pivot2.cell(row=1, column=na_col_idx).font = dark_green_font
            except ValueError:
                pass  # "-n/a-" column not present

    except Exception as e:
        print(f"Error creating pivot table: {e}")
        # Optionally, write a small error message to the sheet if it exists
        try:
            ws_error_pivot = workbook.create_sheet('Pivot_Error')
            ws_error_pivot.append([f"Could not generate pivot table: {e}"])
        except: # pragma: no cover
            pass # If sheet creation also fails

def generate_survey_report(
    rid_file_stream, metrics_file_stream, actual_loi, output_dir,
    conversion_rate_threshold=10,
    security_terms_threshold=30,
    speeder_multiplier=3,
    high_loi_multiplier=3,
    negative_recs_rate_threshold=15,
    process_status_26_only=True,
    use_datetime_for_newuser=True
):
    """
    Processes the survey files and generates an Excel report with observations and a pivot table.
    """
    if not (3 <= actual_loi <= 100):
        raise ValueError("Survey Actual LOI must be between 3 and 100.")

    rid_df = read_rid_file_from_stream(rid_file_stream)
    metrics_df = read_metrics_file_from_stream(metrics_file_stream)

    if 'pid' not in rid_df.columns or 'pid' not in metrics_df.columns:
        raise ValueError("Critical Error: 'pid' column not found in one or both input files. Please check column headers (e.g., 'PID', 'Pid').")
    
    # Ensure 'pid' columns are strings for reliable merging
    rid_df['pid'] = rid_df['pid'].astype(str).str.strip()
    metrics_df['pid'] = metrics_df['pid'].astype(str).str.strip()

    merged_df = pd.merge(
        rid_df,
        metrics_df,
        on='pid',
        how='left',
        indicator=False
    )

    # Filter for status=26 if requested
    if process_status_26_only and 'client_responsestatusid' in merged_df.columns:
        merged_df = merged_df[merged_df['client_responsestatusid'].astype(str) == '26'].copy()

    # Apply observation logic and add check columns
    merged_df = apply_pid_observation_logic(
        merged_df,
        actual_loi=actual_loi,
        conversion_rate_threshold=conversion_rate_threshold,
        security_terms_threshold=security_terms_threshold,
        speeder_multiplier=speeder_multiplier,
        high_loi_multiplier=high_loi_multiplier,
        negative_recs_rate_threshold=negative_recs_rate_threshold,
        session_loi_checks=True,
        use_datetime_for_newuser=use_datetime_for_newuser
    )

    # Remove blank columns (all values are NaN or empty) before writing to Excel
    merged_df = merged_df.dropna(axis=1, how='all')

    # --- Generate Excel File ---
    os.makedirs(output_dir, exist_ok=True) # Ensure output directory exists
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"survey_report_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='Combined Data', index=False)
        add_check_results_pivot(writer, merged_df)  # First create multi-check pivot
        add_pivot_and_format(writer, merged_df)     # Then create other pivots
    
    return output_path

def generate_pid_only_report(
    metrics_file_stream,
    output_dir,
    conversion_rate_threshold=10,
    security_terms_threshold=30,
    negative_recs_rate_threshold=15,
    use_datetime_for_newuser=True  # <-- add this
):
    """
    Processes only the PID Metrics file and generates an Excel report with observations (PID-only mode).
    Skips session_loi-based checks and ignores speeder/high_loi multipliers.
    Adds a simple Observation pivot sheet with formatting.
    Also deletes all blank columns from the output.
    In the Observation Pivot sheet, do not add a row total column since there is only one column with numbers.
    """
    import pandas as pd
    from datetime import datetime
    import os
    from openpyxl.styles import Font
    metrics_df = pd.read_excel(metrics_file_stream, sheet_name='Marketplace Metrics by PID', skiprows=5)
    metrics_df.columns = metrics_df.columns.str.strip().str.lower()
    if 'pid' not in metrics_df.columns:
        raise ValueError("Critical Error: 'pid' column not found in PID Metrics file.")
    # Remove blank columns (all values are NaN or empty)
    metrics_df = metrics_df.dropna(axis=1, how='all')
    # Pass None for speeder_multiplier and high_loi_multiplier, and session_loi_checks=False
    apply_pid_observation_logic(
        metrics_df,
        actual_loi=None,
        conversion_rate_threshold=conversion_rate_threshold,
        security_terms_threshold=security_terms_threshold,
        speeder_multiplier=None,
        high_loi_multiplier=None,
        negative_recs_rate_threshold=negative_recs_rate_threshold,
        session_loi_checks=False,
        use_datetime_for_newuser=use_datetime_for_newuser  # <-- pass this
    )
    os.makedirs(output_dir, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"pid_metrics_report_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, output_filename)
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        metrics_df.to_excel(writer, sheet_name='PID Metrics Data', index=False)
        workbook = writer.book
        ws_data = writer.sheets['PID Metrics Data']
        ws_data.freeze_panes = ws_data['B2']
        ws_data.auto_filter.ref = ws_data.dimensions
        # Add Observation Pivot with formatting and conditional totals
        if 'Observation' in metrics_df.columns:
            obs_counts = metrics_df['Observation'].value_counts().reset_index()
            obs_counts.columns = ['Observation', 'Count']
            na_row = obs_counts[obs_counts['Observation'] == '-n/a-']
            other_rows = obs_counts[obs_counts['Observation'] != '-n/a-'].sort_values('Count', ascending=False)
            obs_counts_sorted = pd.concat([na_row, other_rows], ignore_index=True)
            # Only add totals if there are multiple unique observations
            add_totals = len(obs_counts_sorted) > 1
            # Do NOT add row total column in PID-only mode
            obs_counts_sorted.to_excel(writer, sheet_name='Observation Pivot', index=False)
            ws_pivot = writer.sheets['Observation Pivot']
            ws_pivot.auto_filter.ref = ws_pivot.dimensions
            ws_pivot.freeze_panes = ws_pivot['B2']
            dark_green_font = Font(color="006400")
            for row in ws_pivot.iter_rows(min_row=2, max_row=2, min_col=1, max_col=2):
                for cell in row:
                    if cell.value == '-n/a-' or (cell.row == 2 and ws_pivot['A2'].value == '-n/a-'):
                        cell.font = dark_green_font
            # Style Row Total row bold and add column total if needed
            if add_totals:
                total_row_idx = ws_pivot.max_row
                # Only add bold if there is a Row Total row (but not a row total column)
                if ws_pivot.cell(row=total_row_idx, column=1).value == 'Row Total':
                    for cell in ws_pivot[total_row_idx]:
                        cell.font = Font(bold=True)
                # Add column total (sum of Count column) in the cell below last row
                ws_pivot.cell(row=ws_pivot.max_row+1, column=1, value='Column Total')
                ws_pivot.cell(row=ws_pivot.max_row, column=2, value=obs_counts_sorted['Count'].sum())
                ws_pivot.cell(row=ws_pivot.max_row, column=2).font = Font(bold=True)
    return output_path

def apply_pid_observation_logic(
    df,
    actual_loi=None,
    conversion_rate_threshold=10,
    security_terms_threshold=30,
    speeder_multiplier=3,
    high_loi_multiplier=3,
    negative_recs_rate_threshold=15,
    session_loi_checks=True,
    use_datetime_for_newuser=True  # Add this parameter
):
    """
    Applies all non-RID-based observation logic to the DataFrame in-place.
    If session_loi_checks is False, skips Speeder and High LOI checks.
    """
    import pandas as pd
    # Dates: Convert to datetime objects if they are not already.
    try:
        df["first_entry_date_time"] = pd.to_datetime(df.get("first_entry_date_time"), errors='coerce')
        df["last_entry_date_time"] = pd.to_datetime(df.get("last_entry_date_time"), errors='coerce')
    except Exception:
        pass
    # Set default value
    df["Observation"] = "-n/a-"
    
    # Initialize check columns as False (not blank)
    check_columns = [
        "Poor_Conv_Rate", "New_User_Bot", "High_Security",
        "Speeder", "High_LOI", "High_RR"
    ]
    
    for col in check_columns:
        df[col] = False
    
    # 1. Poor Conversion Rate
    mask_poor_conversion = df["system_conversion_rate"] < (conversion_rate_threshold / 100.0)
    df.loc[mask_poor_conversion.fillna(False), "Poor_Conv_Rate"] = True
    df.loc[mask_poor_conversion.fillna(False), "Observation"] = "Poor Conversion Rate"

    # 2. New User (bot?)
    if use_datetime_for_newuser:
        mask_new_user = (df["first_entry_date_time"] == df["last_entry_date_time"]) & \
                       (df["first_entry_date_time"].notna()) & \
                       (df["last_entry_date_time"].notna())
    else:
        mask_new_user = (df["first_entry_date"] == df["last_entry_date"]) & \
                       (df["first_entry_date"].notna()) & \
                       (df["last_entry_date"].notna())
    df.loc[mask_new_user.fillna(False), "New_User_Bot"] = True
    df.loc[mask_new_user.fillna(False), "Observation"] = "New User (bot?)"

    # 3. High Security Terms
    mask_high_security = (df["total_system_entrants"] > 0) & \
                       ((df["sum_f_and_g_column"] / df["total_system_entrants"]) > (security_terms_threshold / 100.0))
    df.loc[mask_high_security.fillna(False), "High_Security"] = True
    df.loc[mask_high_security.fillna(False), "Observation"] = "High Security Terms"

    # 4. Speeder & 5. High LOI
    if session_loi_checks and actual_loi is not None and speeder_multiplier and high_loi_multiplier:
        mask_speeder = df["session_loi"] < (actual_loi / speeder_multiplier)
        df.loc[mask_speeder.fillna(False), "Speeder"] = True
        df.loc[mask_speeder.fillna(False), "Observation"] = "Speeder"

        mask_high_loi = df["session_loi"] > (actual_loi * high_loi_multiplier)
        df.loc[mask_high_loi.fillna(False), "High_LOI"] = True
        df.loc[mask_high_loi.fillna(False), "Observation"] = "High LOI, Distracted"

    # 6. High RR%
    mask_high_rr = df["negative_recs_rate"] > (negative_recs_rate_threshold / 100.0)
    df.loc[mask_high_rr.fillna(False), "High_RR"] = True
    df.loc[mask_high_rr.fillna(False), "Observation"] = "High RR%"
    
    return df