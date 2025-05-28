# survey_processor.py
import pandas as pd
from datetime import datetime
import os
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

def add_pivot_and_format(writer, df_merged):
    """
    Adds a pivot table sheet to the Excel workbook and applies basic formatting.
    This function is adapted from the original MPpro.py.
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
        ws_pivot = workbook.create_sheet('Observation Pivot')
        
        # Write header (supplier_bu and then pivot columns)
        header = ['supplier_bu'] + list(pivot.columns)
        ws_pivot.append(header)
        
        # Write data rows
        for supplier_bu_index, row_data in pivot.iterrows():
            ws_pivot.append([supplier_bu_index] + list(row_data.values))

        # --- Style "-n/a-" column in dark green ---
        from openpyxl.styles import Font
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
            ws_pivot1 = workbook.create_sheet('Pivot EntryDate x Supplier')
            # Write header
            ws_pivot1.append(['entrydate'] + list(pivot_entrydate_supplier.columns))
            # Write data rows
            for idx, row in pivot_entrydate_supplier.iterrows():
                ws_pivot1.append([idx] + list(row.values))

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
            ws_pivot2 = workbook.create_sheet('Pivot EntryDate x Observation')
            # Write header
            ws_pivot2.append(['entrydate'] + list(pivot_entrydate_obs.columns))
            # Write data rows
            for idx, row in pivot_entrydate_obs.iterrows():
                ws_pivot2.append([idx] + list(row.values))

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
    process_status_26_only=True
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

    # --- Observation Logic (from original MPpro.py) ---
    # Dates: Convert to datetime objects if they are not already. Assuming standard date formats.
    # Pandas might auto-convert if excel cells are date formatted. If not, explicit conversion is needed.
    # For robustness, let's attempt conversion and handle potential errors.
    try:
        merged_df["first_entry_date_time"] = pd.to_datetime(merged_df.get("first_entry_date_time"), errors='coerce')
        merged_df["last_entry_date_time"] = pd.to_datetime(merged_df.get("last_entry_date_time"), errors='coerce')
    except Exception: # pragma: no cover
        # If date columns are not critical for an observation or already handled by .get(), this can be pass
        # Otherwise, raise an error or log a warning
        print("Warning: Could not parse date columns 'first_entry_date' or 'last_entry_date'.")

    # Set default value
    merged_df["Observation"] = "-n/a-"

    # 1. Poor Conversion Rate (<conversion_rate_threshold%)
    mask_poor_conversion = merged_df["system_conversion_rate"] < (conversion_rate_threshold / 100.0)
    merged_df.loc[mask_poor_conversion.fillna(False), "Observation"] = "Poor Conversion Rate"

    # 2. New User (bot?) (first_entry_date_time == last_entry_date_time and not NaT)
    mask_new_user = (merged_df["first_entry_date_time"] == merged_df["last_entry_date_time"]) & \
                    (merged_df["first_entry_date_time"].notna()) & \
                    (merged_df["last_entry_date_time"].notna())
    merged_df.loc[mask_new_user.fillna(False), "Observation"] = "New User (bot?)"

    # 3. High Security Terms (sum_f_and_g_column / total_surveys_entered > security_terms_threshold%)
    mask_high_security = (merged_df["total_surveys_entered"] > 0) & \
                         ((merged_df["sum_f_and_g_column"] / merged_df["total_surveys_entered"]) > (security_terms_threshold / 100.0))
    merged_df.loc[mask_high_security.fillna(False), "Observation"] = "High Security Terms"

    # 4. Speeder (session_loi < actual_loi / speeder_multiplier)
    mask_speeder = merged_df["session_loi"] < (actual_loi / speeder_multiplier)
    merged_df.loc[mask_speeder.fillna(False), "Observation"] = "Speeder"

    # 5. High LOI, Distracted (session_loi > actual_loi * high_loi_multiplier)
    mask_high_loi = merged_df["session_loi"] > (actual_loi * high_loi_multiplier)
    merged_df.loc[mask_high_loi.fillna(False), "Observation"] = "High LOI, Distracted"

    # 6. High RR% (negative_recs_rate > negative_recs_rate_threshold%)
    mask_high_rr = merged_df["negative_recs_rate"] > (negative_recs_rate_threshold / 100.0)
    merged_df.loc[mask_high_rr.fillna(False), "Observation"] = "High RR%"

    # --- Generate Excel File ---
    os.makedirs(output_dir, exist_ok=True) # Ensure output directory exists
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_filename = f"survey_report_{timestamp}.xlsx"
    output_path = os.path.join(output_dir, output_filename)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        merged_df.to_excel(writer, sheet_name='Combined Data', index=False)
        add_pivot_and_format(writer, merged_df) # Call the pivot table function

    return output_path

def generate_pid_only_report(
    metrics_file_stream,
    output_dir,
    conversion_rate_threshold=10,
    security_terms_threshold=30,
    negative_recs_rate_threshold=15
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
        session_loi_checks=False
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
    session_loi_checks=True
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
    # 1. Poor Conversion Rate (<conversion_rate_threshold%)
    mask_poor_conversion = df["system_conversion_rate"] < (conversion_rate_threshold / 100.0)
    df.loc[mask_poor_conversion.fillna(False), "Observation"] = "Poor Conversion Rate"
    # 2. New User (bot?) (first_entry_date_time == last_entry_date_time and not NaT)
    mask_new_user = (df["first_entry_date_time"] == df["last_entry_date_time"]) & \
                    (df["first_entry_date_time"].notna()) & \
                    (df["last_entry_date_time"].notna())
    df.loc[mask_new_user.fillna(False), "Observation"] = "New User (bot?)"
    # 3. High Security Terms (sum_f_and_g_column / total_surveys_entered > security_terms_threshold%)
    mask_high_security = (df["total_surveys_entered"] > 0) & \
                         ((df["sum_f_and_g_column"] / df["total_surveys_entered"]) > (security_terms_threshold / 100.0))
    df.loc[mask_high_security.fillna(False), "Observation"] = "High Security Terms"
    if session_loi_checks and actual_loi is not None and speeder_multiplier and high_loi_multiplier:
        # 4. Speeder (session_loi < actual_loi / speeder_multiplier)
        mask_speeder = df["session_loi"] < (actual_loi / speeder_multiplier)
        df.loc[mask_speeder.fillna(False), "Observation"] = "Speeder"
        # 5. High LOI, Distracted (session_loi > actual_loi * high_loi_multiplier)
        mask_high_loi = df["session_loi"] > (actual_loi * high_loi_multiplier)
        df.loc[mask_high_loi.fillna(False), "Observation"] = "High LOI, Distracted"
    # 6. High RR% (negative_recs_rate > negative_recs_rate_threshold%)
    mask_high_rr = df["negative_recs_rate"] > (negative_recs_rate_threshold / 100.0)
    df.loc[mask_high_rr.fillna(False), "Observation"] = "High RR%"
    return df