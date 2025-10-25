import pandas as pd
from xlsxwriter.utility import xl_col_to_name

def write_denylist_draft(workbook, df_out, config):
    """Create DenyList_Draft sheet with filtered flagged records"""
    
    if config.get('user_feedback', True):
        print("Creating DenyList Draft sheet")
    
    if config.get('debug'):
        print("[DEBUG] Creating DenyList_Draft sheet")
        print("[DEBUG] DenyList_Draft: df_out shape:", df_out.shape)
        print("[DEBUG] DenyList_Draft: df_out columns:", df_out.columns.tolist())

    # Define columns for DenyList_Draft - RID added at the end
    denylist_columns = [
        "PID", "Supplier ID", "Supplier Name", "Deny Criteria",
        "Speeder", "High_LOI", "Poor_Conv_Rate", "High_Security", "New_User_Bot", "High_RR", "No_Enough_Data",
        "Flag_Count", "PrioFlag", "Tenure_Group", "entrydate_split", "RID"
    ]
    # Mapping from DenyList_Draft columns to Combined Data columns
    combined_map = {
        "Supplier ID": "supplierid",
        "Supplier Name": "name",
        "Speeder": "Speeder",
        "High_LOI": "High_LOI",
        "Poor_Conv_Rate": "Poor_Conv_Rate",
        "High_Security": "High_Security",
        "New_User_Bot": "New_User_Bot",
        "High_RR": "High_RR",
        "No_Enough_Data": "No_Enough_Data",
        "Flag_Count": "Flag_Count",
        "PrioFlag": "PrioFlag",
        "Tenure_Group": "Tenure_Group",
        "entrydate_split": "entrydate_split"
        # RID not in map - uses special TEXTJOIN formula
    }

    # Get all unique PIDs (no sorting)
    pid_col = "pid"
    if pid_col not in df_out.columns:
        if config.get('debug'):
            print("[DEBUG] PID column missing in Combined Data")
        return

    unique_pids = pd.Series(df_out[pid_col].unique()).dropna().tolist()
    if config.get('debug'):
        print(f"[DEBUG] DenyList_Draft: unique PIDs count: {len(unique_pids)}")

    deny_ws = workbook.add_worksheet("DenyList_Draft")

    # Header formatting
    fmt_header = workbook.add_format({'bold': True})
    fmt_header_bg = workbook.add_format({'bold': True, 'bg_color': '#E4DFEC'})
    # Data column background color
    fmt_data_bg = workbook.add_format({'bg_color': '#E4DFEC'})
    # No bg for first 4 columns
    fmt_no_bg = workbook.add_format({})

    # Write headers
    for col_idx, col_name in enumerate(denylist_columns):
        if col_idx < 4:
            deny_ws.write(0, col_idx, col_name, fmt_header)
        else:
            deny_ws.write(0, col_idx, col_name, fmt_header_bg)

    # Write data rows with XLOOKUP formulas
    for row_idx, pid in enumerate(unique_pids, start=1):
        if config.get('debug') and row_idx <= 5:
            print(f"[DEBUG] DenyList_Draft: Writing PID {pid} at row {row_idx+1}")
        deny_ws.write(row_idx, 0, pid, fmt_no_bg)  # PID column

        # Supplier ID
        deny_ws.write_formula(row_idx, 1,
            f'=XLOOKUP(A{row_idx+1},\'Combined Data\'!R:R,\'Combined Data\'!Z:Z,"")',
            fmt_no_bg)
        # Supplier Name
        deny_ws.write_formula(row_idx, 2,
            f'=XLOOKUP(A{row_idx+1},\'Combined Data\'!R:R,\'Combined Data\'!W:W,"")',
            fmt_no_bg)
        # Deny Criteria
        deny_ws.write(row_idx, 3, 10, fmt_no_bg)

        # Data columns (Speeder, High_LOI, etc.)
        for col_idx, col_name in enumerate(denylist_columns[4:-1], start=4):  # Exclude RID (last column)
            combined_col = combined_map[col_name]
            # Find column letter in Combined Data
            if combined_col in df_out.columns:
                combined_col_idx = df_out.columns.get_loc(combined_col)
                combined_col_letter = xl_col_to_name(combined_col_idx)
                deny_ws.write_formula(row_idx, col_idx,
                    f'=XLOOKUP(A{row_idx+1},\'Combined Data\'!R:R,\'Combined Data\'!{combined_col_letter}:{combined_col_letter},"")',
                    fmt_data_bg)
            else:
                deny_ws.write(row_idx, col_idx, "", fmt_data_bg)

        # RID column (last column) - XLOOKUP formula (returns first matching RID)
        rid_col_idx = len(denylist_columns) - 1
        deny_ws.write_formula(row_idx, rid_col_idx,
            f'=XLOOKUP(A{row_idx+1},\'Combined Data\'!R:R,\'Combined Data\'!A:A,"")',
            fmt_data_bg)

    # Autofilter
    deny_ws.autofilter(0, 0, len(unique_pids), len(denylist_columns)-1)
    # Freeze first row and first 4 columns
    deny_ws.freeze_panes(1, 4)

    # Conditional formatting for flag columns: red font if TRUE
    flag_cols = ["Speeder", "High_LOI", "Poor_Conv_Rate", "High_Security", "New_User_Bot", "High_RR"]
    for col_name in flag_cols:
        col_idx = denylist_columns.index(col_name)
        deny_ws.conditional_format(1, col_idx, len(unique_pids), col_idx, {
            'type': 'formula',
            'criteria': f'=AND(${xl_col_to_name(col_idx)}2=TRUE)',
            'format': workbook.add_format({'font_color': 'red'})
        })
   
   # Conditional formatting for flag column "No_Enough_Data": #666666 font if TRUE
    flag_cols = ["No_Enough_Data"]
    for col_name in flag_cols:
        col_idx = denylist_columns.index(col_name)
        deny_ws.conditional_format(1, col_idx, len(unique_pids), col_idx, {
            'type': 'formula',
            'criteria': f'=AND(${xl_col_to_name(col_idx)}2=TRUE)',
            'format': workbook.add_format({'font_color': '#666666'})
        })
        
    # Conditional formatting for PrioFlag column
    prioflag_idx = denylist_columns.index("PrioFlag")
    deny_ws.conditional_format(1, prioflag_idx, len(unique_pids), prioflag_idx, {
        'type': 'formula',
        'criteria': f'=${xl_col_to_name(prioflag_idx)}2="Recent User, No Enough Data"',
        'format': workbook.add_format({'font_color': "#666666"})
    })
    deny_ws.conditional_format(1, prioflag_idx, len(unique_pids), prioflag_idx, {
        'type': 'formula',
        'criteria': f'=${xl_col_to_name(prioflag_idx)}2="No Flags"',
        'format': workbook.add_format({'font_color': "#026102"})
    })

    # Conditional formatting (same rules as Combined Data)
    # Flag_Count red to white color scale, max value 7
    flag_count_idx = denylist_columns.index("Flag_Count")
    deny_ws.conditional_format(1, flag_count_idx, len(unique_pids), flag_count_idx, {
        'type': '2_color_scale',
        'max_color': "#FF0000",
        'min_color': "#E5E0EC",
        'max_type': 'num',
        'min_value': 0,
        'max_value': 7
    })
        
    # Set column widths for readability
    deny_ws.set_column(0, 0, 18)  # PID
    deny_ws.set_column(1, 1, 11.5)  # Supplier ID
    deny_ws.set_column(2, 2, 22)    # Supplier Name
    deny_ws.set_column(3, 3, 7.5)   # Deny Criteria
    # Data columns widths
    for idx, col_name in enumerate(denylist_columns[4:], start=4):
        if col_name in ["Speeder", "High_LOI", "Poor_Conv_Rate", "High_Security", "New_User_Bot", "High_RR", "No_Enough_Data", "Flag_Count"]:
            deny_ws.set_column(idx, idx, 9, fmt_data_bg)
        elif col_name == "PrioFlag":
            deny_ws.set_column(idx, idx, 32, fmt_data_bg)
        elif col_name == "Tenure_Group":
            deny_ws.set_column(idx, idx, 20, fmt_data_bg)
        elif col_name == "RID":
            deny_ws.set_column(idx, idx, 37.71, fmt_data_bg)  # Match Combined Data RID column width
        else:
            deny_ws.set_column(idx, idx, 16, fmt_data_bg)

    if config.get('debug'):
        print("[DEBUG] DenyList_Draft sheet created with", len(unique_pids), "rows.")

