# pivot_sheets.py - Pivot sheet creation and reordering
import pandas as pd
from xlsxwriter.utility import xl_col_to_name

def write_flag_pivot(workbook, df_out, config):
    if config.get('user_feedback', True):
        print("Creating Counts sheet with summary tables")
    
    # Add a new worksheet for the counts
    counts_ws = workbook.add_worksheet('Counts')
    
    # Start position for first table
    current_row = 0
    
    # Write PrioFlag Count table
    prioflag_total_row = write_prioflag_count_table(counts_ws, df_out, current_row, workbook, config)
    current_row = prioflag_total_row + 3  # Add 2 blank rows between tables
    
    # Write MultiFlag Count table - pass prioflag_total_row for reference
    current_row = write_multiflag_count_table(counts_ws, df_out, current_row, workbook, config, prioflag_total_row)
    current_row += 3  # Add 2 blank rows between tables
    
    # Write additional count tables in the specified order
    tables_info = [
        ('client_responsestatus', 'Client Response Status Counts'),
        ('Tenure_Group', 'Tenure Group Counts'),
        ('surveyid', 'Survey ID Counts'),
        ('survey_country', 'Survey Country Counts'),
        ('fulcrum_responsestatus', 'Fulcrum Response Status Counts'),
        ('project_manager', 'Project Manager Counts'),
        ('buyer_bu', 'Buyer BU Counts'),
        ('link_type', 'Link Type Counts')
    ]
    
    for column_name, table_title in tables_info:
        current_row = write_general_count_table(counts_ws, df_out, column_name, table_title, current_row, workbook, config)
        current_row += 3  # Add 2 blank rows between tables
    
    # Set column widths
    counts_ws.set_column(0, 0, 35)  # Column A - labels
    counts_ws.set_column(1, 1, 8)   # Column B - counts (changed from 15 to 8)
    counts_ws.set_column(2, 2, 15)  # Column C - percentages
    
    # Freeze top row
    counts_ws.freeze_panes(1, 0)

def write_prioflag_count_table(counts_ws, df_out, start_row, workbook, config):
    """Write PrioFlag Count table using TRANSPOSE formulas from PrioFlag Pivot sheet"""
    
    if config.get('debug', False):
        print(f"DEBUG: write_prioflag_count_table called with start_row: {start_row}")
    
    # Background colors matching PrioFlag Pivot columns B through K
    bg_colors = [
        '#DCE6F1',  # B
        '#D8E4BC',  # C  
        '#EBF1DE',  # D
        '#F2DCDB',  # E
        '#F2DCDB',  # F
        '#F2DCDB',  # G
        '#F2DCDB',  # H
        '#F2DCDB',  # I
        '#F2DCDB',  # J
        '#E6B8B7'   # K
    ]
    
    # Define formats with background colors
    fmt_header = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
    fmt_count_header = workbook.add_format({'align': 'right', 'border': 1, 'bold': True})
    fmt_percent_header = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%'})
    
    current_row = start_row
    
    # Write table headers
    if config.get('debug', False):
        print(f"DEBUG: Writing headers at row {current_row}")
    counts_ws.write(current_row, 0, "PrioFlag Counts:", fmt_header)
    counts_ws.write(current_row, 1, "Count", fmt_count_header)
    counts_ws.write(current_row, 2, "%", fmt_percent_header)
    current_row += 1
    
    data_start_row = current_row
    
    if config.get('debug', False):
        print(f"DEBUG: data_start_row: {data_start_row}")
        print(f"DEBUG: About to write single TRANSPOSE formulas in first row")

    # Write single TRANSPOSE formulas in first data row (let Excel handle spill)
    # Column A: Labels
    counts_ws.write_formula(data_start_row, 0, "=TRANSPOSE(XLOOKUP(\"Supplier ↓\",'PrioFlag Pivot'!A:A,'PrioFlag Pivot'!B:K))",
                           workbook.add_format({'align': 'left', 'border': 1}))
    
    # Column B: Counts
    counts_ws.write_formula(data_start_row, 1, "=TRANSPOSE(XLOOKUP(\"Total\",'PrioFlag Pivot'!A:A,'PrioFlag Pivot'!B:K))",
                           workbook.add_format({'align': 'right', 'border': 1, 'bold': True}))
    
    # Column C: Percentages
    counts_ws.write_formula(data_start_row, 2, "=TRANSPOSE(XLOOKUP(\"%\",'PrioFlag Pivot'!A:A,'PrioFlag Pivot'!B:K))",
                           workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%'}))
    
    if config.get('debug', False):
        print(f"DEBUG: All TRANSPOSE formulas written in row {data_start_row}")

    # Apply background colors to individual cells for all three columns (keep current method)
    for i in range(10):  # 10 flag types (B through K columns)
        row_num = data_start_row + i
        if i > 0: # skip addional formatting for first row
            if config.get('debug', False) and i < 3:  # Only print first 3 for brevity
                print(f"DEBUG: Applying background colors to row {row_num} (i={i}) with color {bg_colors[i]}")
        
            # Apply background color formats for each column
            fmt_label = workbook.add_format({'align': 'left', 'border': 1, 'bg_color': bg_colors[i]})
            fmt_count = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'bg_color': bg_colors[i]})
            fmt_percent = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%', 'bg_color': bg_colors[i]})
            
            # Apply formatting to individual cells
            counts_ws.set_row(row_num - 1, None, None)  # Clear any existing row format
            counts_ws.write(row_num, 0, None, fmt_label)    # Apply format to label cell
            counts_ws.write(row_num, 1, None, fmt_count)    # Apply format to count cell
            counts_ws.write(row_num, 2, None, fmt_percent)  # Apply format to percent cell 
        
    # Add blue data bars to percentage column
    percent_data_range = f"C{data_start_row + 1}:C{data_start_row + 10}"
    counts_ws.conditional_format(percent_data_range, {
        'type': 'data_bar',
        'bar_color': '#5B9BD5',
        'data_bar_2010': True,
        'bar_only': False
    })
    
    if config.get('debug', False):
        print(f"DEBUG: Returning final row: {data_start_row + 10}")
    
    return data_start_row + 10  # Return row after last data row (no Total row)

def write_multiflag_count_table(counts_ws, df_out, start_row, workbook, config, prioflag_total_row):
    """Write MultiFlag Count table using TRANSPOSE formulas from MultiFlag Pivot sheet"""
    
    if config.get('debug', False):
        print(f"DEBUG: write_multiflag_count_table called with start_row: {start_row}")
    
    # Background colors matching MultiFlag Pivot columns B through K (same as PrioFlag)
    bg_colors = [
        '#DCE6F1',  # B
        '#D8E4BC',  # C  
        '#EBF1DE',  # D
        '#F2DCDB',  # E
        '#F2DCDB',  # F
        '#F2DCDB',  # G
        '#F2DCDB',  # H
        '#F2DCDB',  # I
        '#F2DCDB',  # J
        '#E6B8B7'   # K
    ]
    
    # Define formats with background colors
    fmt_header = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
    fmt_count_header = workbook.add_format({'align': 'right', 'border': 1, 'bold': True})
    fmt_percent_header = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%'})
    
    current_row = start_row
    
    # Write table headers
    if config.get('debug', False):
        print(f"DEBUG: Writing MultiFlag headers at row {current_row}")
    counts_ws.write(current_row, 0, "MultiFlag Counts:", fmt_header)
    counts_ws.write(current_row, 1, "Count", fmt_count_header)
    counts_ws.write(current_row, 2, "%", fmt_percent_header)
    current_row += 1
    
    data_start_row = current_row
    
    if config.get('debug', False):
        print(f"DEBUG: MultiFlag data_start_row: {data_start_row}")
        print(f"DEBUG: About to write TRANSPOSE formulas referencing MultiFlag Pivot sheet")

    # Write single TRANSPOSE formulas in first data row (let Excel handle spill)
    # Column A: Labels from MultiFlag Pivot
    counts_ws.write_formula(data_start_row, 0, "=TRANSPOSE(XLOOKUP(\"Supplier ↓\",'MultiFlag Pivot'!A:A,'MultiFlag Pivot'!B:K))",
                           workbook.add_format({'align': 'left', 'border': 1}))
    
    # Column B: Counts from MultiFlag Pivot
    counts_ws.write_formula(data_start_row, 1, "=TRANSPOSE(XLOOKUP(\"Total\",'MultiFlag Pivot'!A:A,'MultiFlag Pivot'!B:K))",
                           workbook.add_format({'align': 'right', 'border': 1, 'bold': True}))
    
    # Column C: Percentages from MultiFlag Pivot
    counts_ws.write_formula(data_start_row, 2, "=TRANSPOSE(XLOOKUP(\"%\",'MultiFlag Pivot'!A:A,'MultiFlag Pivot'!B:K))",
                           workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%'}))
    
    if config.get('debug', False):
        print(f"DEBUG: All MultiFlag TRANSPOSE formulas written in row {data_start_row}")

    # Apply background colors to individual cells for all three columns (same method as PrioFlag)
    for i in range(10):  # 10 flag types (B through K columns)
        row_num = data_start_row + i
        if i > 0: # skip additional formatting for first row
            if config.get('debug', False) and i < 3:  # Only print first 3 for brevity
                print(f"DEBUG: Applying MultiFlag background colors to row {row_num} (i={i}) with color {bg_colors[i]}")
        
            # Apply background color formats for each column
            fmt_label = workbook.add_format({'align': 'left', 'border': 1, 'bg_color': bg_colors[i]})
            fmt_count = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'bg_color': bg_colors[i]})
            fmt_percent = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%', 'bg_color': bg_colors[i]})
            
            # Apply formatting to individual cells
            counts_ws.set_row(row_num - 1, None, None)  # Clear any existing row format
            counts_ws.write(row_num, 0, None, fmt_label)    # Apply format to label cell
            counts_ws.write(row_num, 1, None, fmt_count)    # Apply format to count cell
            counts_ws.write(row_num, 2, None, fmt_percent)  # Apply format to percent cell
        
    # Add blue data bars to percentage column
    percent_data_range = f"C{data_start_row + 1}:C{data_start_row + 10}"
    counts_ws.conditional_format(percent_data_range, {
        'type': 'data_bar',
        'bar_color': '#5B9BD5',
        'data_bar_2010': True,
        'bar_only': False
    })
    
    if config.get('debug', False):
        print(f"DEBUG: Returning MultiFlag final row: {data_start_row + 10}")
    
    return data_start_row + 10  # Return row after last data row (no Total row)

def write_general_count_table(counts_ws, df_out, column_name, table_title, start_row, workbook, config):
    """Write general count table for specified column"""
    
    if column_name not in df_out.columns:
        if config.get('debug', False):
            print(f"DEBUG: Column '{column_name}' not found in Combined Data, skipping table")
        return start_row
    
    # Special handling for Tenure_Group
    if column_name == 'Tenure_Group':
        return write_tenure_group_count_table(counts_ws, table_title, start_row, workbook, config)
    
    # Regular handling for other columns
    # Get unique values and their counts, sorted by count descending
    value_counts = df_out[column_name].fillna('(blank)').value_counts()
    
    # Updated column letter mapping for Combined Data
    col_letters = {
        'rid': 'A', 'buyer_account_id': 'B', 'buyer_account': 'C', 'buyer_bu': 'D',
        'buyer_bu_id': 'E', 'survey_client': 'F', 'client_responsestatusid': 'G',
        'client_responsestatus': 'H', 'link_type_id': 'I', 'external_survey_name': 'J',
        'fulcrum_responsestatusid': 'K', 'fulcrum_responsestatus': 'L',
        'internal_survey_name': 'M', 'marketplace_projectid': 'N', 'marketplace_project': 'O',
        'mid': 'P', 'parentsid': 'Q', 'pid': 'R', 'respondentsid': 'S',
        'entrydate': 'T', 'lastdate': 'U', 'id': 'V', 'name': 'W',
        'supplier_bu_id': 'X', 'link_type': 'Y', 'supplierid': 'Z',
        'survey_country': 'AA', 'survey_country_langauge': 'AB', 'survey_ccpi': 'AC',
        'survey_HASH_status': 'AD', 'survey_SCCB_status': 'AE', 'survey_https_status': 'AF',
        'project_manager': 'AG', 'pm_email': 'AH', 'survey_qcpi': 'AI',
        'total_system_entrants': 'AJ', 'total_completes': 'AK', 'total_negative_recs': 'AL',
        'total_security_terms_on_marketplace_side': 'AM', 'total_security_terms_on_client_side': 'AN',
        'total_security_terms': 'AO', 'first_entry_time': 'AP', 'last_exit_time': 'AQ',
        'net_recs_rate': 'AR', 'supplier_bu': 'AS', 'first_entry_date': 'AT',
        'last_entry_date': 'AU', 'Tenure': 'AV', 'total_surveys_entered': 'AW',
        'system_conversion_rate': 'AX', 'security_terms_rate': 'AY', 'negative_recs_rate': 'AZ',
        'surveyid': 'BA', 'CompLOI': 'BB', 'session_loi': 'BC', 'speeder_multiplier': 'BD',
        'high_loi_multiplier': 'BE', 'surveys_entered_threshold': 'BF', 'conversion_rate_threshold': 'BG',
        'security_terms_threshold': 'BH', 'negative_recs_rate_threshold': 'BI', 'Speeder': 'BJ',
        'High_LOI': 'BK', 'Poor_Conv_Rate': 'BL', 'High_Security': 'BM', 'New_User_Bot': 'BN',
        'High_RR': 'BO', 'No_Enough_Data': 'BP', 'Flag_Count': 'BQ', 'PrioFlag': 'BR',
        'Tenure_Group': 'BS'
    }
    
    col_letter = col_letters.get(column_name, 'A')
    
    # Create count data using formulas - UPDATED to exclude blank RID rows
    count_data = {}
    total_records = len(df_out)
    
    for value in value_counts.index:
        if value == '(blank)':
            # Count blank/empty values in target column, but only where RID is not blank
            count_data[value] = f'=COUNTIFS(\'Combined Data\'!{col_letter}:{col_letter},"",\'Combined Data\'!A:A,"<>")+COUNTIFS(\'Combined Data\'!{col_letter}:{col_letter}," ",\'Combined Data\'!A:A,"<>")'
        else:
            # Escape quotes in the value for formula and add RID non-blank condition
            escaped_value = str(value).replace('"', '""')
            count_data[value] = f'=COUNTIFS(\'Combined Data\'!{col_letter}:{col_letter},"{escaped_value}",\'Combined Data\'!A:A,"<>")'
    
    # Write the table
    return write_count_table(counts_ws, count_data, table_title, start_row, workbook, total_records)

def write_tenure_group_count_table(counts_ws, table_title, start_row, workbook, config):
    """Special handler for Tenure_Group count table with hardcoded categories"""
    
    if config.get('user_feedback', True):
        print(f"Writing {table_title} table")
    
    if config.get('debug', False):
        print(f"DEBUG: Writing Tenure_Group table at row {start_row}")
    
    # Hardcoded tenure categories in exact order
    tenure_categories = [
        "Less than a week",
        "Less than a month", 
        "Less than 3 months",
        "Less than 6 months",
        "Less than a year",
        "More than a year"
    ]
    
    # Define formats
    fmt_header = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
    fmt_data = workbook.add_format({'align': 'left', 'border': 1})
    fmt_count = workbook.add_format({'align': 'right', 'border': 1, 'bold': True})
    fmt_percent = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%'})
    fmt_total = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
    
    current_row = start_row
    
    # Write table headers
    counts_ws.write(current_row, 0, table_title, fmt_header)
    counts_ws.write(current_row, 1, "Count", fmt_count)
    counts_ws.write(current_row, 2, "%", fmt_percent)
    current_row += 1
    
    data_start_row = current_row
    
    # Write data rows with dynamic COUNTIFS formulas - UPDATED to exclude blank RID rows
    for category in tenure_categories:
        counts_ws.write(current_row, 0, category, fmt_data)
        
        # Column B: COUNTIFS formula referencing current row in column A and excluding blank RIDs
        count_formula = f"=COUNTIFS('Combined Data'!BS:BS,A{current_row + 1},'Combined Data'!A:A,\"<>\")"
        counts_ws.write_formula(current_row, 1, count_formula, fmt_count)
        
        # Column C: Percentage formula
        count_cell = f"B{current_row + 1}"
        total_cell = f"B{data_start_row + len(tenure_categories) + 1}"  # Reference to Total row
        percent_formula = f"=IF({total_cell}>0,{count_cell}/{total_cell},0)"
        counts_ws.write_formula(current_row, 2, percent_formula, fmt_percent)
        
        current_row += 1
    
    # Write Total row - UPDATED to count only rows with non-blank RIDs and exclude header
    total_row = current_row
    counts_ws.write(total_row, 0, "Total", fmt_total)
    
    # Count total non-blank RID rows instead of summing individual categories, subtract 1 for header
    counts_ws.write_formula(total_row, 1, "=COUNTIF('Combined Data'!A:A,\"<>\") - 1", fmt_count)
    
    # Total percentage is always 100%
    counts_ws.write_formula(total_row, 2, "=1", fmt_percent)
    
    # Add blue data bars to percentage column (including Total row)
    percent_range = f"C{data_start_row + 1}:C{total_row + 1}"
    counts_ws.conditional_format(percent_range, {
        'type': 'data_bar',
        'bar_color': '#5B9BD5',
        'data_bar_2010': True,
        'bar_only': False
    })
    
    if config.get('debug', False):
        print(f"DEBUG: Tenure_Group table completed, returning row {total_row + 1}")
    
    return total_row + 1

def write_count_table(counts_ws, count_data, table_title, start_row, workbook, total_records, prioflag_total_row=None):
    """Helper function to write count table with consistent formatting"""
    
    # Define formats
    fmt_header = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
    fmt_data = workbook.add_format({'align': 'left', 'border': 1})
    fmt_count = workbook.add_format({'align': 'right', 'border': 1, 'bold': True})
    fmt_percent = workbook.add_format({'align': 'right', 'border': 1, 'bold': True, 'num_format': '0.00%'})
    fmt_total = workbook.add_format({'bold': True, 'align': 'left', 'border': 1})
    
    current_row = start_row
    
    # Write table headers
    counts_ws.write(current_row, 0, table_title, fmt_header)
    counts_ws.write(current_row, 1, "Count", fmt_count)
    counts_ws.write(current_row, 2, "%", fmt_percent)
    current_row += 1
    
    data_start_row = current_row
    
    # Write data rows
    for label, count_formula in count_data.items():
        counts_ws.write(current_row, 0, label, fmt_data)
        counts_ws.write_formula(current_row, 1, count_formula, fmt_count)
        
        # Calculate percentage formula
        count_cell = f"B{current_row + 1}"
        total_cell = f"B{data_start_row + len(count_data) + 1}"  # Reference to Total row
        percent_formula = f"=IF({total_cell}>0,{count_cell}/{total_cell},0)"
        counts_ws.write_formula(current_row, 2, percent_formula, fmt_percent)
        
        current_row += 1
    
    # Write Total row - UPDATED to count only non-blank RID rows and exclude header
    total_row = current_row
    counts_ws.write(total_row, 0, "Total", fmt_total)
    
    if prioflag_total_row is not None:
        # Special case for MultiFlag table - reference PrioFlag Total
        counts_ws.write_formula(total_row, 1, f"=B{prioflag_total_row}", fmt_count)
    else:
        # UPDATED: Count total non-blank RID rows instead of summing categories, subtract 1 for header
        counts_ws.write_formula(total_row, 1, "=COUNTIF('Combined Data'!A:A,\"<>\") - 1", fmt_count)
    
    # Use formula instead of hardcoded value for percentage
    counts_ws.write_formula(total_row, 2, "=1", fmt_percent)
    
    # Add blue data bars to percentage column (including Total row)
    percent_range = f"C{data_start_row + 1}:C{total_row + 1}"
    counts_ws.conditional_format(percent_range, {
        'type': 'data_bar',
        'bar_color': '#5B9BD5',
        'data_bar_2010': True,
        'bar_only': False
    })
    
    return total_row + 1