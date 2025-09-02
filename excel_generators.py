# excel_generators.py - Excel file creation, formatting, and pivot generation
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import ColorScaleRule, Rule
from openpyxl.styles.differential import DifferentialStyle
import numpy as np

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
    # Guard: skip if no data rows
    if data_rows < 2:
        return
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
            # Guard: skip if cell_range is not valid
            if int(data_rows) < 2:
                continue
            try:
                worksheet.conditional_formatting.add(cell_range, color_scale_rule)
            except Exception as e:
                # Log and skip this column if openpyxl fails
                print(f"Conditional formatting skipped for {cell_range}: {e}")
                continue

def apply_na_column_formatting(worksheet, header):
    """Apply dark green formatting to -n/a- column if present."""
    print("DEBUG: Starting apply_na_column_formatting...")
    print("DEBUG: Header type and content:", type(header), header)
    
    if '-n/a-' in header:
        print("DEBUG: Found -n/a- in header")
        na_col_idx = header.index('-n/a-') + 1  # openpyxl is 1-based
        dark_green_font = Font(color="006400")
        print(f"DEBUG: Applying formatting to column {na_col_idx}")
        
        try:
            for row in worksheet.iter_rows(min_row=2, min_col=na_col_idx, max_col=na_col_idx, max_row=worksheet.max_row):
                for cell in row:
                    cell.font = dark_green_font
            worksheet.cell(row=1, column=na_col_idx).font = dark_green_font
            print("DEBUG: -n/a- column formatting completed successfully")
        except Exception as e:
            print(f"DEBUG: Error in apply_na_column_formatting: {e}")
            raise
    else:
        print("DEBUG: -n/a- not found in header")

def format_pivot_sheet(worksheet, header, data_rows, total_column, total_rows, has_col_total_row=True):
    """Standardized pivot formatting: conditional formatting + -n/a- styling."""
    print("DEBUG: Starting format_pivot_sheet...")
    print("DEBUG: Header:", header)
    print("DEBUG: Data rows:", data_rows)
    print("DEBUG: Total column:", total_column)
    
    try:
        exclude_cols, last_data_row = get_formatting_ranges(
            header, data_rows, total_column=total_column, has_col_total_row=has_col_total_row
        )
        print("DEBUG: Got formatting ranges successfully")
        
        apply_conditional_formatting(
            worksheet,
            start_col=2,
            end_col=len(header),
            data_rows=last_data_row,
            exclude_cols=exclude_cols,
            total_rows=total_rows
        )
        print("DEBUG: Applied conditional formatting successfully")
        
        apply_na_column_formatting(worksheet, header)
        print("DEBUG: format_pivot_sheet completed successfully")
        
    except Exception as e:
        print(f"DEBUG: Error in format_pivot_sheet: {e}")
        import traceback
        traceback.print_exc()
        raise

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
            'supplier_bu': str(supplier) if supplier is not None else '',  # Ensure string
            '-n/a-': int(na_count),
            **{col: int(flag_counts[col]) for col in check_columns},  # Ensure integers
            'Total_Flagged': int(total_flagged)
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
    
    # Write to Excel with formatting - ensure all values are properly typed
    header_strings = [str(h) for h in header]  # Ensure header items are strings
    ws_pivot.append(header_strings)
    
    # Write data rows with explicit type conversion
    for _, row in pivot_df.iterrows():
        row_values = []
        for col in header:
            value = row[col]
            if col == 'supplier_bu':
                # Ensure supplier_bu is string
                row_values.append(str(value) if value is not None else '')
            else:
                # Ensure numeric values are proper integers
                try:
                    row_values.append(int(value) if not pd.isna(value) else 0)
                except (ValueError, TypeError):
                    row_values.append(0)
        ws_pivot.append(row_values)
    
    # Set supplier_bu column width to 200 pixels
    try:
        excel_width = float(200.0 / 7.0)
        ws_pivot.column_dimensions['A'].width = excel_width
    except Exception as e:
        print(f"DEBUG: Error setting supplier_bu column width: {e}")
        pass

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
    apply_na_column_formatting(ws_pivot, header)

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

    # Debug: Print available columns before checking for required columns
    print("DEBUG: Columns in merged DataFrame before pivot:", df_merged.columns.tolist())

    # --- Create Pivot Table ---
    if 'supplier_bu' not in df_merged.columns or 'Observation' not in df_merged.columns:
        available_cols = df_merged.columns.tolist()
        raise ValueError(
            f"Required columns 'supplier_bu' or 'Observation' not found in processed data. "
            f"Available columns: {available_cols}. This may indicate a problem with the input files or merge logic."
        )

    try:
        print("DEBUG: Starting pivot creation...")
        pivot = (
            df_merged.groupby(['supplier_bu', 'Observation'])
            .size()
            .reset_index(name='Count')
            .pivot(index='supplier_bu', columns='Observation', values='Count')
            .fillna(0)
            .astype(int)
        )
        print("DEBUG: Main pivot created successfully")

        if pivot.empty:
            raise ValueError("No data available for pivot table generation. Please check that your input files contain valid data.")

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
        
        # Write header (supplier_bu and then pivot columns) - ensure strings
        header = ['supplier_bu'] + [str(col) for col in pivot.columns]
        ws_pivot.append(header)
        
        # Write data rows with explicit type conversion
        for supplier_bu_index, row_data in pivot.iterrows():
            row_values = [str(supplier_bu_index) if supplier_bu_index is not None else '']  # Ensure supplier name is string
            for value in row_data.values:
                try:
                    row_values.append(int(value) if not pd.isna(value) else 0)  # Ensure integers
                except (ValueError, TypeError):
                    row_values.append(0)
            ws_pivot.append(row_values)

        # Set supplier_bu column width to 200 pixels
        try:
            excel_width = float(200.0 / 7.0)
            ws_pivot.column_dimensions['A'].width = excel_width
        except Exception as e:
            print(f"DEBUG: Error setting supplier_bu column width: {e}")
            pass

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
        print("DEBUG: Starting entrydate processing...")

        # Helper: get just the date part from entrydate - robust conversion for mixed data types
        def safe_convert_entrydate(x):
            if pd.isna(x) or x is None:
                return ''
            # Convert to string and handle any data type
            str_val = str(x)
            # Split by 'T' and take first part (date portion)
            return str_val.split('T')[0] if 'T' in str_val else str_val
        
        # Ensure entrydate column exists and convert it safely
        if 'entrydate' in df_merged.columns:
            print("DEBUG: Converting entrydate column...")
            print("DEBUG: entrydate sample values before conversion:", df_merged['entrydate'].head().tolist())
            print("DEBUG: entrydate data types:", df_merged['entrydate'].dtype)
            df_merged['entrydate_only'] = df_merged['entrydate'].apply(safe_convert_entrydate)
            print("DEBUG: entrydate_only sample values after conversion:", df_merged['entrydate_only'].head().tolist())
        else:
            print("DEBUG: entrydate column not found, creating empty column")
            # If entrydate doesn't exist, create empty column
            df_merged['entrydate_only'] = ''

        # 1. Pivot: entrydate_only vs supplier_bu
        if 'entrydate_only' in df_merged.columns and 'supplier_bu' in df_merged.columns:
            print("DEBUG: Starting Pivot EntryDate x Supplier...")
            print("DEBUG: supplier_bu sample values before conversion:", df_merged['supplier_bu'].head().tolist())
            print("DEBUG: supplier_bu data types:", df_merged['supplier_bu'].dtype)
            
            # Ensure supplier_bu is also string type to avoid groupby issues
            df_merged['supplier_bu'] = df_merged['supplier_bu'].astype(str)
            print("DEBUG: supplier_bu converted to string")
            
            try:
                print("DEBUG: Creating entrydate vs supplier pivot...")
                pivot_entrydate_supplier = (
                    df_merged.groupby(['entrydate_only', 'supplier_bu'])
                    .size()
                    .reset_index(name='Count')
                    .pivot(index='entrydate_only', columns='supplier_bu', values='Count')
                    .fillna(0)
                    .astype(int)
                )
                print("DEBUG: Pivot EntryDate x Supplier created successfully")
                
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
                header = ['entrydate'] + [str(col) for col in pivot_entrydate_supplier.columns]  # Ensure strings
                ws_pivot1.append(header)
                
                # Write data with explicit type conversion
                for idx, row_data in pivot_entrydate_supplier.iterrows():
                    row_values = [str(idx) if idx is not None else '']  # Ensure date string
                    for value in row_data.values:
                        try:
                            row_values.append(int(value) if not pd.isna(value) else 0)
                        except (ValueError, TypeError):
                            row_values.append(0)
                    ws_pivot1.append(row_values)

                # Set all columns to 100 pixels width - fix potential float/int issue
                excel_width = 100.0 / 7.0  # Ensure float division
                for col_idx in range(1, len(header) + 1):
                    col_letter = ws_pivot1.cell(row=1, column=col_idx).column_letter
                    try:
                        ws_pivot1.column_dimensions[col_letter].width = float(excel_width)
                    except Exception as e:
                        print(f"DEBUG: Error setting column width for {col_letter}: {e}")
                        # Skip this column width setting if it fails
                        pass

                format_pivot_sheet(ws_pivot1, header, len(pivot_entrydate_supplier), 'Row Total', len(df_merged))

            except Exception as e:
                print(f"DEBUG: Error in Pivot EntryDate x Supplier: {e}")
                raise

        # 2. Pivot: entrydate_only vs Observation
        if 'entrydate_only' in df_merged.columns and 'Observation' in df_merged.columns:
            print("DEBUG: Starting Pivot EntryDate x Flags...")
            print("DEBUG: Observation sample values before conversion:", df_merged['Observation'].head().tolist())
            print("DEBUG: Observation data types:", df_merged['Observation'].dtype)
            
            # Ensure Observation is also string type
            df_merged['Observation'] = df_merged['Observation'].astype(str)
            print("DEBUG: Observation converted to string")
            
            try:
                print("DEBUG: Creating entrydate vs observation pivot...")
                pivot_entrydate_obs = (
                    df_merged.groupby(['entrydate_only', 'Observation'])
                    .size()
                    .reset_index(name='Count')
                    .pivot(index='entrydate_only', columns='Observation', values='Count')
                    .fillna(0)
                    .astype(int)
                )
                print("DEBUG: Pivot EntryDate x Flags created successfully")
                
                print("DEBUG: Starting pivot operations...")
                # Sort index by date (ascending)
                pivot_entrydate_obs = pivot_entrydate_obs.sort_index()
                print("DEBUG: Index sorted")
                
                # Add row totals
                pivot_entrydate_obs['Row Total'] = pivot_entrydate_obs.sum(axis=1)
                print("DEBUG: Row totals added")
                
                # Add column totals
                col_totals = pivot_entrydate_obs.sum()
                col_totals.name = 'Column Total'
                pivot_entrydate_obs = pd.concat([pivot_entrydate_obs, pd.DataFrame([col_totals], index=['Column Total'])])
                print("DEBUG: Column totals added")
                
                # Sort columns: "-n/a-" first, then others by total (highest first), then Row Total last
                cols = list(pivot_entrydate_obs.columns)
                print("DEBUG: Columns before sorting:", cols)
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
                print("DEBUG: Columns sorted:", sorted_cols)

                print("DEBUG: Creating worksheet...")
                ws_pivot2 = workbook.create_sheet('Pivot EntryDate x Flags')
                header = ['entrydate'] + [str(h) for h in pivot_entrydate_obs.columns]  # Ensure all strings
                print("DEBUG: Final header:", header)
                
                ws_pivot2.append(header)
                print("DEBUG: Header written to worksheet")
                
                # Write data with explicit type conversion
                for idx, row in pivot_entrydate_obs.iterrows():
                    row_values = [str(idx) if idx is not None else '']  # Ensure date string
                    for value in row.values:
                        try:
                            row_values.append(int(value) if not pd.isna(value) else 0)
                        except (ValueError, TypeError):
                            row_values.append(0)
                    ws_pivot2.append(row_values)
                print("DEBUG: Data written to worksheet")
                
                # Fix column width setting
                try:
                    ws_pivot2.column_dimensions['A'].width = float(100.0 / 7.0)
                except Exception as e:
                    print(f"DEBUG: Error setting column width for column A: {e}")
                    # Skip if column width setting fails
                    pass
                print("DEBUG: Column width set")

                print("DEBUG: About to call format_pivot_sheet...")
                format_pivot_sheet(ws_pivot2, header, len(pivot_entrydate_obs), 'Row Total', len(df_merged))
                print("DEBUG: format_pivot_sheet completed")

            except Exception as e:
                print(f"DEBUG: Error in Pivot EntryDate x Flags: {e}")
                import traceback
                traceback.print_exc()
                raise

    except Exception as e:
        print(f"DEBUG: Error in add_pivot_and_format at line: {e}")
        import traceback
        print("DEBUG: Full traceback:")
        traceback.print_exc()
        print(f"Error creating pivot table: {e}")
        # Return early to prevent further errors
        return

def create_denylist_draft_sheet(workbook, merged_df):
    """Create the DenyList_Draft sheet with proper formatting."""
    print("DEBUG: Creating DenyList_Draft sheet...")
    
    # Define columns for DenyList_Draft
    deny_cols = [
        "pid", "supplierid", "name", "Observation",
        "Poor_Conv_Rate", "New_User_Bot", "High_Security",
        "Speeder", "High_LOI", "High_RR", "Flag_Count",
        "Diff Days"
    ]
    
    # Only keep columns that exist in merged_df
    deny_cols_present = [c for c in deny_cols if c in merged_df.columns]
    deny_df = merged_df[deny_cols_present].copy()

    # Ensure proper data types for DenyList_Draft
    if 'pid' in deny_df.columns:
        deny_df['pid'] = deny_df['pid'].astype(str).replace('nan', '')
    if 'supplierid' in deny_df.columns:
        deny_df['supplierid'] = deny_df['supplierid'].astype(str).replace('nan', '')
    if 'name' in deny_df.columns:
        deny_df['name'] = deny_df['name'].astype(str).replace('nan', '')
    if 'Flag_Count' in deny_df.columns:
        deny_df['Flag_Count'] = pd.to_numeric(deny_df['Flag_Count'], errors='coerce')
    if 'Observation' in deny_df.columns:
        deny_df['Observation'] = deny_df['Observation'].astype(str).replace('nan', '')

    # Rename columns first, then insert Deny Criteria
    col_map = {
        "pid": "PID",
        "supplierid": "Supplier ID", 
        "name": "Supplier Name"
    }
    deny_df = deny_df.rename(columns=col_map)
    
    # Insert "Deny Criteria" column after "Supplier Name"
    cols_list = list(deny_df.columns)
    if "Supplier Name" in cols_list:
        name_idx = cols_list.index("Supplier Name")
        deny_df.insert(name_idx + 1, "Deny Criteria", 10)
    else:
        # Fallback: insert after second column
        deny_df.insert(2, "Deny Criteria", 10)

    # Sort by Flag_Count in descending order (highest values at the top)
    if 'Flag_Count' in deny_df.columns:
        deny_df = deny_df.sort_values('Flag_Count', ascending=False)
        print("DEBUG: DenyList_Draft sorted by Flag_Count (descending)")
    
    # Create the sheet and write data properly with explicit data types
    deny_sheet = workbook.create_sheet("DenyList_Draft")
    
    # Write headers - ensure all headers are strings
    for col_idx, col_name in enumerate(deny_df.columns, 1):
        header_value = str(col_name) if col_name is not None else ''
        cell = deny_sheet.cell(row=1, column=col_idx, value=header_value)
        cell.data_type = 's'
    
    # Write data rows with proper data types and string conversion
    for row_idx, (_, row_data) in enumerate(deny_df.iterrows(), 2):
        for col_idx, (col_name, value) in enumerate(row_data.items(), 1):
            cell = deny_sheet.cell(row=row_idx, column=col_idx)
            
            # Ensure value is properly typed and converted
            if col_name in ['Flag_Count', 'Deny Criteria', 'Diff Days']:
                if pd.isna(value):
                    cell.value = 0
                else:
                    try:
                        cell.value = float(value) if value != '' else 0
                    except (ValueError, TypeError):
                        cell.value = 0
                cell.data_type = 'n'  # Numeric
            elif col_name in ['Poor_Conv_Rate', 'New_User_Bot', 'High_Security', 'Speeder', 'High_LOI', 'High_RR']:
                # Ensure boolean values are properly handled
                if pd.isna(value):
                    cell.value = False
                else:
                    cell.value = bool(value)
                cell.data_type = 'b'  # Boolean
            else:
                # Convert to string and handle None/NaN
                if pd.isna(value) or value is None:
                    cell.value = ''
                else:
                    cell.value = str(value)
                cell.data_type = 's'  # String

    # Enable auto-filter and freeze first row
    deny_sheet.auto_filter.ref = deny_sheet.dimensions
    deny_sheet.freeze_panes = deny_sheet['A2']
    
    return deny_sheet, deny_df

def apply_denylist_conditional_formatting(deny_sheet, deny_df, merged_df):
    """Apply conditional formatting to DenyList_Draft sheet."""
    print("DEBUG: Starting conditional formatting for DenyList_Draft sheet...")
    deny_header = [cell.value for cell in deny_sheet[1]]
    deny_n_rows = deny_sheet.max_row
    print("DEBUG: DenyList_Draft header:", deny_header)
    print("DEBUG: DenyList_Draft rows:", deny_n_rows)

    # Helper to get min/max for DenyList_Draft data
    def get_denylist_col_min_max(col_name):
        if col_name in deny_df.columns:
            col_data = pd.to_numeric(deny_df[col_name], errors='coerce')
            col_min = np.nanmin(col_data)
            col_max = np.nanmax(col_data)
            return col_min, col_max
        return None, None

    # Apply conditional formatting for Diff Days and Flag_Count using same specs as Combined Data
    deny_format_specs = {
        "Flag_Count": {
            "min_color": "FFFFFF", "max_color": "f82b1b", "min": 0, "max": 5, "reverse": False
        },
        "Diff Days": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        }
    }

    for col_name, spec in deny_format_specs.items():
        if col_name in deny_header:
            try:
                col_idx = deny_header.index(col_name) + 1
                col_letter = deny_sheet.cell(row=1, column=col_idx).column_letter
                cell_range = f"{col_letter}2:{col_letter}{deny_n_rows}"

                if "min" in spec and "max" in spec:
                    col_min, col_max = spec["min"], spec["max"]
                else:
                    col_min, col_max = get_denylist_col_min_max(col_name)
                    if col_min is None or col_max is None or col_min == col_max:
                        print(f"DEBUG: Skipping {col_name} - no valid data range")
                        continue
                
                color_rule = ColorScaleRule(
                    start_type='num', start_value=col_min, start_color=spec["min_color"],
                    end_type='num', end_value=col_max, end_color=spec["max_color"]
                )
                deny_sheet.conditional_formatting.add(cell_range, color_rule)
                print(f"DEBUG: Applied conditional formatting to {col_name} in DenyList_Draft (range {col_min}-{col_max}).")
            except Exception as e:
                print(f"DEBUG: Could not apply conditional formatting to {col_name} in DenyList_Draft: {e}")
                import traceback
                traceback.print_exc()

    # Apply red font formatting for TRUE values in boolean flag columns
    print("DEBUG: Applying red font formatting for TRUE flag columns in DenyList_Draft...")
    red_font = Font(color="FF0000")
    red_dxf = DifferentialStyle(font=red_font)
    
    deny_flag_cols_to_format = [
        "Poor_Conv_Rate", "New_User_Bot", "High_Security",
        "Speeder", "High_LOI", "High_RR"
    ]

    for col_name in deny_flag_cols_to_format:
        if col_name in deny_header:
            try:
                col_idx = deny_header.index(col_name) + 1
                col_letter = deny_sheet.cell(row=1, column=col_idx).column_letter
                cell_range = f"{col_letter}2:{col_letter}{deny_n_rows}"
                
                # Rule to apply red font if cell value is TRUE
                rule = Rule(type="expression", dxf=red_dxf)
                rule.formula = [f'{col_letter}2=TRUE']
                
                deny_sheet.conditional_formatting.add(cell_range, rule)
                print(f"DEBUG: Applied red font formatting to {col_name} for TRUE values in DenyList_Draft.")
            except Exception as e:
                print(f"DEBUG: Could not apply red font formatting for {col_name} in DenyList_Draft: {e}")

    # Apply dark green formatting for "-n/a-" values in Observation column
    print("DEBUG: Styling Observation column in DenyList_Draft...")
    if "Observation" in deny_header:
        obs_col_idx = deny_header.index("Observation") + 1
        dark_green_font = Font(color="006400")
        for row in deny_sheet.iter_rows(min_row=2, min_col=obs_col_idx, max_col=obs_col_idx, max_row=deny_sheet.max_row):
            for cell in row:
                if str(cell.value) == "-n/a-":
                    cell.font = dark_green_font
        print("DEBUG: Observation column styling completed in DenyList_Draft")

    print("DEBUG: DenyList_Draft conditional formatting completed")

def apply_combined_data_formatting(combined_sheet, merged_df):
    """Apply comprehensive conditional formatting to Combined Data sheet."""
    print("DEBUG: Starting conditional formatting for Combined Data...")
    
    header = [cell.value for cell in combined_sheet[1]]
    n_rows = combined_sheet.max_row
    print("DEBUG: Header for conditional formatting:", header)
    print("DEBUG: Number of rows for formatting:", n_rows)

    # Helper to get min/max and ensure numeric
    def get_col_min_max(col_name):
        if col_name in merged_df.columns:
            col_data = pd.to_numeric(merged_df[col_name], errors='coerce')
            col_min = np.nanmin(col_data)
            col_max = np.nanmax(col_data)
            return col_min, col_max
        return None, None

    # Conditional formatting rules using exact column names from specification
    format_specs = {
        # system_conversion_rate: high is good, red for low (bad), white for high (good), scale 0-100
        "system_conversion_rate": {
            "min_color": "FFFFFF", "max_color": "f82b1b", "min": 0, "max": 100, "reverse": True
        },
        # Security_Terms_Rate: high is bad, red for high (bad), white for low (good), scale 0-100
        "Security_Terms_Rate": {
            "min_color": "FFFFFF", "max_color": "f82b1b", "min": 0, "max": 100, "reverse": False
        },
        # net recs rate: high is bad, red for high (bad), white for low (good), scale 0-100
        "net recs rate": {
            "min_color": "FFFFFF", "max_color": "f82b1b", "min": 0, "max": 100, "reverse": False
        },
        # negative_recs_rate: high is bad, red for high (bad), white for low (good), scale 0-100
        "negative_recs_rate": {
            "min_color": "FFFFFF", "max_color": "f82b1b", "min": 0, "max": 100, "reverse": False
        },
        # Flag_Count: fixed scale 0-5, white for low, red for high
        "Flag_Count": {
            "min_color": "FFFFFF", "max_color": "f82b1b", "min": 0, "max": 5, "reverse": False
        },
        # client_responsestatusid: orange to pickle green
        "client_responsestatusid": {
            "min_color": "FFA500", "max_color": "4f9e4f", "reverse": False
        },
        # session_loi: 3-color scale yellow-white-yellow
        "session_loi": {
            "min_color": "FFFF00", "mid_color": "FFFFFF", "max_color": "FFFF00", "reverse": False, "three_color": True
        },
        # supplier_bu_id: sky blue to gray
        "supplier_bu_id": {
            "min_color": "87CEEB", "max_color": "808080", "reverse": False
        },
        # survey_ccpi: white to yellow
        "survey_ccpi": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        # survey_qcpi: white to yellow
        "survey_qcpi": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        # Diff Days: white to yellow
        "Diff Days": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        # Updated PID sheet columns using exact names: white (0) to yellow (high values)
        "total_system_entrants": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        "total_surveys_entered": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        "total_completes": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        "total_negative_recs": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        "total_security_terms_on_marketplace_side": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        "total_security_terms_on_client_side": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        },
        "total security terms": {
            "min_color": "FFFFFF", "max_color": "FFFF00", "reverse": False
        }
    }

    print("DEBUG: Applying conditional formatting rules...")
    for col_name, spec in format_specs.items():
        if col_name in header:
            print(f"DEBUG: Applying formatting to column: {col_name}")
            try:
                col_idx = header.index(col_name) + 1
                col_letter = combined_sheet.cell(row=1, column=col_idx).column_letter
                cell_range = f"{col_letter}2:{col_letter}{n_rows}"
                
                # Get min/max
                if "min" in spec and "max" in spec:
                    col_min, col_max = spec["min"], spec["max"]
                else:
                    col_min, col_max = get_col_min_max(col_name)
                    if col_min is None or col_max is None or col_min == col_max:
                        continue
                        
                # 3-color scale for session_loi
                if spec.get("three_color"):
                    col_median = np.nanmedian(pd.to_numeric(merged_df[col_name], errors='coerce'))
                    color_rule = ColorScaleRule(
                        start_type='num', start_value=col_min, start_color=spec["min_color"],
                        mid_type='num', mid_value=col_median, mid_color=spec["mid_color"],
                        end_type='num', end_value=col_max, end_color=spec["max_color"]
                    )
                else:
                    if spec.get("reverse"):
                        color_rule = ColorScaleRule(
                            start_type='num',
                            start_value=col_min,
                            start_color=spec["max_color"],
                            end_type='num',
                            end_value=col_max,
                            end_color=spec["min_color"]
                        )
                    else:
                        color_rule = ColorScaleRule(
                            start_type='num',
                            start_value=col_min,
                            start_color=spec["min_color"],
                            end_type='num',
                            end_value=col_max,
                            end_color=spec["max_color"]
                        )
                
                combined_sheet.conditional_formatting.add(cell_range, color_rule)
                print(f"DEBUG: Successfully applied formatting to {col_name}")
            except (TypeError, AttributeError) as e:
                # Graceful degradation for openpyxl compatibility issues
                print(f"DEBUG: Warning: Could not apply conditional formatting for {col_name}: {e}")
                continue
            except Exception as e:
                print(f"DEBUG: Error applying formatting to {col_name}: {e}")
                import traceback
                traceback.print_exc()
                raise

    print("DEBUG: Styling -n/a- column in Combined Data...")
    # Style -n/a- column in dark green if present
    if "-n/a-" in header:
        na_col_idx = header.index("-n/a-") + 1
        dark_green_font = Font(color="006400")
        for row in combined_sheet.iter_rows(min_row=2, min_col=na_col_idx, max_col=na_col_idx, max_row=combined_sheet.max_row):
            for cell in row:
                cell.font = dark_green_font
        combined_sheet.cell(row=1, column=na_col_idx).font = dark_green_font
        print("DEBUG: -n/a- column styling completed")

    print("DEBUG: Styling Observation column...")
    # Style Observation column "-n/a-" values in dark green
    if "Observation" in header:
        obs_col_idx = header.index("Observation") + 1
        dark_green_font = Font(color="006400")
        for row in combined_sheet.iter_rows(min_row=2, min_col=obs_col_idx, max_col=obs_col_idx, max_row=combined_sheet.max_row):
            for cell in row:
                if str(cell.value) == "-n/a-":
                    cell.font = dark_green_font
                    
    print("DEBUG: Applying conditional formatting for TRUE flag columns...")
    red_font = Font(color="FF0000")
    dxf = DifferentialStyle(font=red_font)
    
    flag_cols_to_format = [
        "Poor_Conv_Rate", "New_User_Bot", "High_Security",
        "Speeder", "High_LOI", "High_RR"
    ]

    for col_name in flag_cols_to_format:
        if col_name in header:
            try:
                col_idx = header.index(col_name) + 1
                col_letter = combined_sheet.cell(row=1, column=col_idx).column_letter
                cell_range = f"{col_letter}2:{col_letter}{n_rows}"
                
                # Rule to apply red font if cell value is TRUE
                rule = Rule(type="expression", dxf=dxf)
                # The formula applies to the top-left cell of the range.
                rule.formula = [f'{col_letter}2=TRUE']
                
                combined_sheet.conditional_formatting.add(cell_range, rule)
                print(f"DEBUG: Applied red font formatting to {col_name} for TRUE values.")
            except Exception as e:
                print(f"DEBUG: Could not apply red font formatting for {col_name}: {e}")
                rule.formula = [f'{col_letter}2=TRUE']
                
                combined_sheet.conditional_formatting.add(cell_range, rule)
                print(f"DEBUG: Applied red font formatting to {col_name} for TRUE values.")
            except Exception as e:
                print(f"DEBUG: Could not apply red font formatting for {col_name}: {e}")
