# excel_generators.py - Excel file creation, formatting, and pivot generation
import pandas as pd
import re
from xlsxwriter.utility import xl_col_to_name

def get_combined_data_columns():
    """Return the centralized column order for Combined Data sheet."""
    return [
        'rid', 'buyer_account_id', 'buyer_account', 'buyer_bu', 'buyer_bu_id', 'survey_client',
        'client_responsestatusid', 'client_responsestatus', 'link_type_id', 'external_survey_name',
        'fulcrum_responsestatusid', 'fulcrum_responsestatus', 'internal_survey_name', 'marketplace_projectid',
        'marketplace_project', 'mid', 'parentsid', 'pid', 'respondentsid', 'entrydate', 'lastdate', 'id', 'name',
        'supplier_bu_id', 'link_type', 'supplierid', 'survey_country', 'survey_country_langauge', 'survey_ccpi',
        'survey_HASH_status', 'survey_SCCB_status', 'survey_https_status', 'project_manager', 'pm_email',
        'survey_qcpi', 'total_system_entrants', 'total_completes', 'total_negative_recs',
        'total_security_terms_on_marketplace_side', 'total_security_terms_on_client_side', 'total_security_terms',
        'first_entry_time', 'last_exit_time', 'net_recs_rate', 'supplier_bu', 'first_entry_date', 'last_entry_date',
        'Tenure', 'total_surveys_entered', 'system_conversion_rate', 'security_terms_rate', 'negative_recs_rate',
        'surveyid', 'CompLOI', 'session_loi', 'speeder_multiplier', 'high_loi_multiplier', 'surveys_entered_threshold',
        'conversion_rate_threshold', 'security_terms_threshold', 'negative_recs_rate_threshold',
        'Speeder', 'High_LOI', 'Poor_Conv_Rate', 'High_Security', 'New_User_Bot', 'High_RR', 'No_Enough_Data',
        'Flag_Count', 'PrioFlag', 'Tenure_Group', 'entrydate_split'  # Added 'entrydate_split'
    ]

def _normalize_name(s):
    """Normalize column name for matching: lowercase + remove non-alphanumerics."""
    if s is None:
        return ''
    return re.sub(r'[^0-9a-z]', '', str(s).lower())

def reorder_and_fill_combined_data(
    df,
    surveys_entered_threshold=None,
    conversion_rate_threshold=None,
    security_terms_threshold=None,
    negative_recs_rate_threshold=None
):
    """Reorder columns and fill missing columns with blanks for Combined Data sheet."""
    # Normalize column names: replace spaces with underscores and lowercase all names
    df.columns = df.columns.str.strip().str.replace(' ', '_').str.lower()

    # Convert first_entry_date and last_entry_date to date format (YYYY-MM-DD)
    date_columns = ['first_entry_date', 'last_entry_date']
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce').dt.strftime('%Y-%m-%d')

    columns = get_combined_data_columns()
    input_cols = list(df.columns)
    norm_map = { _normalize_name(c): c for c in input_cols }
    data = {}
    n_rows = len(df)
    for col in columns:
        target_norm = _normalize_name(col)
        if target_norm in norm_map:
            src_col = norm_map[target_norm]
            # For threshold columns, fill blank/empty/NaN values with provided value
            if col in [
                'surveys_entered_threshold',
                'conversion_rate_threshold',
                'security_terms_threshold',
                'negative_recs_rate_threshold'
            ]:
                col_values = df[src_col].values.tolist()
                fill_val = None
                if col == 'surveys_entered_threshold':
                    fill_val = surveys_entered_threshold
                elif col == 'conversion_rate_threshold':
                    fill_val = conversion_rate_threshold
                elif col == 'security_terms_threshold':
                    fill_val = security_terms_threshold
                elif col == 'negative_recs_rate_threshold':
                    fill_val = negative_recs_rate_threshold
                # Fill blank/empty/NaN values with threshold
                data[col] = [
                    fill_val if (v is None or str(v).strip() == '' or (isinstance(v, float) and pd.isna(v))) else v
                    for v in col_values
                ] if fill_val is not None else col_values
            else:
                data[col] = df[src_col].values.tolist()
        else:
            # For threshold columns, fill with provided value if available
            if col == 'surveys_entered_threshold' and surveys_entered_threshold is not None:
                data[col] = [surveys_entered_threshold] * n_rows
            elif col == 'conversion_rate_threshold' and conversion_rate_threshold is not None:
                data[col] = [conversion_rate_threshold] * n_rows
            elif col == 'security_terms_threshold' and security_terms_threshold is not None:
                data[col] = [security_terms_threshold] * n_rows
            elif col == 'negative_recs_rate_threshold' and negative_recs_rate_threshold is not None:
                data[col] = [negative_recs_rate_threshold] * n_rows
            else:
                data[col] = [''] * n_rows

    out_df = pd.DataFrame(data)

    # Coerce numeric columns (attempt) and boolean flags for better Excel formatting
    numeric_targets = [
        'total_system_entrants', 'total_completes', 'total_negative_recs',
        'total_security_terms_on_marketplace_side', 'total_security_terms_on_client_side',
        'total security terms', 'net_recs_rate',  # Removed 'first_entry_time', 'last_exit_time'
        'total_surveys_entered', 'system_conversion_rate', 'security_terms_rate', 'negative_recs_rate',
        'speeder_multiplier', 'high_loi_multiplier', 'conversion_rate_threshold',
        'security_terms_threshold', 'negative_recs_rate_threshold', 'Flag_Count', 'Tenure'  # Changed 'Diff Days' to 'Tenure'
    ]
    for nc in numeric_targets:
        if nc in out_df.columns:
            out_df[nc] = pd.to_numeric(out_df[nc], errors='coerce')

    bool_targets = ['Speeder', 'High_LOI', 'Poor_Conv_Rate', 'High_Security', 'New_User_Bot', 'High_RR', 'No_Enough_Data']
    for bc in bool_targets:
        if bc in out_df.columns:
            # convert truthy strings and numbers to boolean
            out_df[bc] = out_df[bc].map(lambda v: bool(v) if pd.notna(v) and str(v).strip() != '' else False)

    return out_df

def write_combined_data_xlsx(
    output_path,
    df_merged,
    surveys_entered_threshold=None,
    conversion_rate_threshold=None,
    security_terms_threshold=None,
    negative_recs_rate_threshold=None,
    is_pid_only_mode=False,
    config=None
):
    """Write Combined Data sheet to Excel file with formulas for dynamic calculations."""
    df_out = reorder_and_fill_combined_data(
        df_merged,
        surveys_entered_threshold=surveys_entered_threshold,
        conversion_rate_threshold=conversion_rate_threshold,
        security_terms_threshold=security_terms_threshold,
        negative_recs_rate_threshold=negative_recs_rate_threshold
    )
    print("DEBUG: df_out security_terms_rate sample:", df_out['security_terms_rate'].head() if 'security_terms_rate' in df_out.columns else "not found")
    
    # Define formula columns that should not be written as data
    formula_columns = ['Speeder', 'High_LOI', 'Poor_Conv_Rate', 'High_Security', 'New_User_Bot', 'High_RR', 'No_Enough_Data', 'Flag_Count', 'Tenure', 'PrioFlag', 'Tenure_Group', 'entrydate_split']  # Added 'entrydate_split'
    
    # Create a copy without formula columns for initial write
    df_for_excel = df_out.copy()
    for col in formula_columns:
        if col in df_for_excel.columns:
            df_for_excel[col] = ''  # Clear the column data
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df_for_excel.to_excel(writer, sheet_name='Combined Data', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Combined Data']

        # Add hyperlinks for surveyid column
        if 'surveyid' in df_out.columns:
            surveyid_col_idx = df_out.columns.get_loc('surveyid')
            fmt_hyper_num = workbook.add_format({'align': 'right', 'font_color': 'blue', 'underline': 1, 'num_format': '0'})
            for row_idx in range(len(df_out)):
                surveyid_value = df_out.at[row_idx, 'surveyid']
                if pd.notna(surveyid_value) and str(surveyid_value).strip():
                    url = f"https://www.samplicio.us/fulcrum/next/surveys/{surveyid_value}/reports"
                    worksheet.write_url(row_idx + 1, surveyid_col_idx, url, string=str(surveyid_value), cell_format=fmt_hyper_num)

        # Define formats
        fmt_num = workbook.add_format({'num_format': '0.00'})
        fmt_int = workbook.add_format({'num_format': '0'})
        fmt_bool = workbook.add_format({'align': 'center'})
        fmt_header = workbook.add_format({'bold': True, 'align': 'left'})  # Left align headers
        fmt_percent = workbook.add_format({'num_format': '0.00%'})
        fmt_red = workbook.add_format({'font_color': 'red'})
        fmt_text = workbook.add_format({'num_format': '@'})  # Text format
        fmt_date = workbook.add_format({'num_format': 'yyyy-mm-dd'})  # Date format

        # Apply header format
        for col_idx, col_name in enumerate(df_out.columns):
            worksheet.write(0, col_idx, col_name, fmt_header)

        # Define comprehensive column width mapping
        column_widths = {
            'rid': 37.71,
            'buyer_account_id': 6.43,
            'buyer_account': 17.29,
            'buyer_bu': 17.29,
            'buyer_bu_id': 6,
            'survey_client': 17.29,
            'client_responsestatusid': 5.29,
            'client_responsestatus': 15,
            'link_type_id': 4.71,
            'external_survey_name': 15,
            'fulcrum_responsestatusid': 5.43,
            'fulcrum_responsestatus': 24.29,
            'internal_survey_name': 21.86,
            'marketplace_projectid': 10.86,
            'marketplace_project': 25.14,
            'mid': 22.71,
            'parentsid': 17.71,
            'pid': 35.57,
            'respondentsid': 12,
            'entrydate': 12.57,
            'lastdate': 11,
            'id': 6.71,
            'name': 30,
            'supplier_bu_id': 8.57,
            'link_type': 20.57,
            'supplierid': 10.43,
            'survey_country': 11.29,
            'survey_country_langauge': 15,
            'survey_ccpi': 5,
            'survey_HASH_status': 5,
            'survey_SCCB_status': 5,
            'survey_https_status': 5,
            'project_manager': 15,
            'pm_email': 15,
            'survey_qcpi': 9.71,
            'total_system_entrants': 15,
            'total_completes': 15,
            'total_negative_recs': 15,
            'total_security_terms_on_marketplace_side': 15,
            'total_security_terms_on_client_side': 15,
            'total_security_terms': 15,
            'first_entry_time': 15,
            'last_exit_time': 15,
            'net_recs_rate': 9.86,
            'supplier_bu': 30,
            'first_entry_date': 11.71,
            'last_entry_date': 11.71,
            'Tenure': 10.71,
            'total_surveys_entered': 10.71,
            'system_conversion_rate': 10.71,
            'security_terms_rate': 10.71,
            'negative_recs_rate': 10.71,
            'surveyid': 10.14,
            'CompLOI': 10.57,
            'session_loi': 4.86,
            'speeder_multiplier': 4.86,
            'high_loi_multiplier': 4.86,
            'surveys_entered_threshold': 4.86,
            'conversion_rate_threshold': 4.86,
            'security_terms_threshold': 4.86,
            'negative_recs_rate_threshold': 4.86,
            'Speeder': 9.86,
            'High_LOI': 9.86,
            'Poor_Conv_Rate': 9.86,
            'High_Security': 9.86,
            'New_User_Bot': 9.86,
            'High_RR': 9.86,
            'No_Enough_Data': 9.86,
            'Flag_Count': 12.43,
            'PrioFlag': 31.86,
            'Tenure_Group': 17.29,
            'entrydate_split': 10.29
        }

        # Set column formats
        for col_idx, col_name in enumerate(df_out.columns):
            lc = _normalize_name(col_name)
            width = column_widths.get(col_name, 15)  # Use mapping or default to 15
            
            if lc in [ _normalize_name(n) for n in ['system_conversion_rate','security_terms_rate','negative_recs_rate','netrecsrate'] ]:
                worksheet.set_column(col_idx, col_idx, width, fmt_percent)
            elif lc in [ _normalize_name(n) for n in ['flag_count','total_system_entrants','total_completes','total_surveys_entered','tenure'] ]:
                worksheet.set_column(col_idx, col_idx, width, fmt_int)
            elif col_name in ['first_entry_time', 'last_exit_time']:
                worksheet.set_column(col_idx, col_idx, width, fmt_text)  # Ensure text format
            elif col_name in ['first_entry_date', 'last_entry_date']:
                worksheet.set_column(col_idx, col_idx, width, fmt_date)  # Ensure date format
            # Apply background color #D8D8D8 for specified columns (removed session_loi)
            elif col_name in [
                'speeder_multiplier',
                'high_loi_multiplier',
                'surveys_entered_threshold',
                'conversion_rate_threshold',
                'security_terms_threshold',
                'negative_recs_rate_threshold'
            ]:
                worksheet.set_column(col_idx, col_idx, width, workbook.add_format({'bg_color': '#D8D8D8'}))
            # Apply background color #E5E0EC for flag columns
            elif col_name in [
                'Speeder',
                'High_LOI',
                'Poor_Conv_Rate',
                'High_Security',
                'New_User_Bot',
                'High_RR'
            ]:
                worksheet.set_column(col_idx, col_idx, width, workbook.add_format({'bg_color': '#E5E0EC', 'align': 'center'}))
                col_letter = xl_col_to_name(col_idx)
                worksheet.conditional_format(f'{col_letter}2:{col_letter}{len(df_out)+1}', {
                    'type': 'formula',
                    'criteria': f'=${col_letter}2=TRUE',
                    'format': fmt_red
                })
            # Apply background color #E5E0EC for specified columns
            elif col_name in [
                'No_Enough_Data',
                'Flag_Count',
                'PrioFlag',
                'Tenure_Group',
                'entrydate_split'
            ]:
                worksheet.set_column(col_idx, col_idx, width, workbook.add_format({'bg_color': '#E5E0EC'}))
            else:
                worksheet.set_column(col_idx, col_idx, width)

        # Freeze header row
        worksheet.freeze_panes(1, 0)
        worksheet.autofilter(0, 0, len(df_out), len(df_out.columns)-1)

        # Column letters mapping (fully updated from user)
        col_letters = {
            'rid': 'A',
            'buyer_account_id': 'B',
            'buyer_account': 'C',
            'buyer_bu': 'D',
            'buyer_bu_id': 'E',
            'survey_client': 'F',
            'client_responsestatusid': 'G',
            'client_responsestatus': 'H',
            'link_type_id': 'I',
            'external_survey_name': 'J',
            'fulcrum_responsestatusid': 'K',
            'fulcrum_responsestatus': 'L',
            'internal_survey_name': 'M',
            'marketplace_projectid': 'N',
            'marketplace_project': 'O',
            'mid': 'P',
            'parentsid': 'Q',
            'pid': 'R',
            'respondentsid': 'S',
            'entrydate': 'T',
            'lastdate': 'U',
            'id': 'V',
            'name': 'W',
            'supplier_bu_id': 'X',
            'link_type': 'Y',
            'supplierid': 'Z',
            'survey_country': 'AA',
            'survey_country_langauge': 'AB',
            'survey_ccpi': 'AC',
            'survey_HASH_status': 'AD',
            'survey_SCCB_status': 'AE',
            'survey_https_status': 'AF',
            'project_manager': 'AG',
            'pm_email': 'AH',
            'survey_qcpi': 'AI',
            'total_system_entrants': 'AJ',
            'total_completes': 'AK',
            'total_negative_recs': 'AL',
            'total_security_terms_on_marketplace_side': 'AM',
            'total_security_terms_on_client_side': 'AN',
            'total_security_terms': 'AO',
            'first_entry_time': 'AP',
            'last_exit_time': 'AQ',
            'net_recs_rate': 'AR',
            'supplier_bu': 'AS',
            'first_entry_date': 'AT',
            'last_entry_date': 'AU',
            'Tenure': 'AV',
            'total_surveys_entered': 'AW',
            'system_conversion_rate': 'AX',
            'security_terms_rate': 'AY',
            'negative_recs_rate': 'AZ',
            'surveyid': 'BA',
            'CompLOI': 'BB',
            'session_loi': 'BC',
            'speeder_multiplier': 'BD',
            'high_loi_multiplier': 'BE',
            'surveys_entered_threshold': 'BF',
            'conversion_rate_threshold': 'BG',
            'security_terms_threshold': 'BH',
            'negative_recs_rate_threshold': 'BI',
            'Speeder': 'BJ',
            'High_LOI': 'BK',
            'Poor_Conv_Rate': 'BL',
            'High_Security': 'BM',
            'New_User_Bot': 'BN',
            'High_RR': 'BO',
            'No_Enough_Data': 'BP',
            'Flag_Count': 'BQ',
            'PrioFlag': 'BR',
            'Tenure_Group': 'BS',
            'entrydate_split': 'BT'
        }

        row_start = 2  # Define the starting row for formulas (row 1 is for headers)
        row_end = len(df_out) + 1  # Define the ending row based on the number of rows in the DataFrame

        # Formulas for each column (using explicit column letters)
        formulas = {
            'Speeder': '=IF({session_loi}{row}<({CompLOI}{row}/{speeder_multiplier}{row}), TRUE, FALSE)',
            'High_LOI': '=IF({session_loi}{row}>({CompLOI}{row}*{high_loi_multiplier}{row}), TRUE, FALSE)',
            'Poor_Conv_Rate': '=IF(AND({system_conversion_rate}{row}<({conversion_rate_threshold}{row}/100), {total_surveys_entered}{row}>{surveys_entered_threshold}{row}), TRUE, FALSE)',
            'High_Security': '=IF(AND({security_terms_rate}{row}>({security_terms_threshold}{row}/100), {total_surveys_entered}{row}>{surveys_entered_threshold}{row}), TRUE, FALSE)',
            'High_RR': '=IF(AND({negative_recs_rate}{row}>({negative_recs_rate_threshold}{row}/100), {total_surveys_entered}{row}>{surveys_entered_threshold}{row}), TRUE, FALSE)',
            'No_Enough_Data': '=IF(AND(NOT({first_entry_date}{row}={last_entry_date}{row}), {total_surveys_entered}{row}<={surveys_entered_threshold}{row}), TRUE, FALSE)',
            'New_User_Bot': '=IF({first_entry_date}{row}={last_entry_date}{row}, TRUE, FALSE)',
            'Flag_Count': '=COUNTIF({Speeder}{row}:{High_RR}{row}, TRUE)',
            'Tenure': '=DATEDIF({first_entry_date}{row}, TODAY(), "d")',
            'PrioFlag': (
                '=IF({High_RR}{row}, "High Reversal Rate", '
                'IF({High_Security}{row}, "High Security Terms Rate", '
                'IF({New_User_Bot}{row}, "New User, First survey, Bot suspect", '
                'IF({High_LOI}{row}, "High LOI, Distracted", '
                'IF({Speeder}{row}, "Low LOI, Speeder", '
                'IF({Poor_Conv_Rate}{row}, "Low Conversion Rate", '
                'IF({No_Enough_Data}{row}, "Recent User, No Enough Data", '
                '"No Flags")))))))'
            ),
            'Tenure_Group': (  # Added formula for Tenure_Group
                '=IF({Tenure}{row}<7, "Less than a week", '
                'IF({Tenure}{row}<30, "Less than a month", '
                'IF({Tenure}{row}<90, "Less than 3 months", '
                'IF({Tenure}{row}<180, "Less than 6 months", '
                'IF({Tenure}{row}<360, "Less than a year", "More than a year")))))'
            ),
            'entrydate_split': '=LEFT({entrydate}{row}, FIND("T", {entrydate}{row})-1)'  # Added formula for entrydate_split
        }

        # Write formulas for each row in each formula column
        for row in range(row_start, row_end + 1):
            for col_name, formula in formulas.items():
                if col_name in col_letters:
                    # Skip formulas for RID-dependent columns in PID-only mode
                    if is_pid_only_mode and col_name in ['Speeder', 'High_LOI', 'entrydate_split']:
                        continue  # Skip writing formula, leave cell empty
                    cell = f'{col_letters[col_name]}{row}'
                    # Remove any leading '=' or '@' from formula string
                    f = formula.format(
                        row=row,
                        **{k: v for k, v in col_letters.items()}
                    )
                    if f.startswith('='):
                        f = f[1:]
                    if f.startswith('@'):
                        f = f[1:]
                    # Debugging: Print the generated formula for PrioFlag
                    # if col_name == 'PrioFlag':
                    #     print(f"DEBUG: Generated formula for PrioFlag at row {row}: {f}")

                    worksheet.write_formula(cell, f)

        # After writing formulas, set entrydate_split column to Date format
        # if 'entrydate_split' in df_out.columns:
        #    entrydate_split_col_idx = df_out.columns.get_loc('entrydate_split')
        #    worksheet.set_column(entrydate_split_col_idx, entrydate_split_col_idx, 15, fmt_date)

        # Conditional formatting for background color scales and font colors using updated column letters
        worksheet.conditional_format('{col}2:{col}{end}'.format(col=col_letters['supplier_bu_id'], end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",
            'mid_color': '#FFFFA8',
            'max_color': '#FF7EFF',
        })  # supplier_bu_id

        worksheet.conditional_format('{col}2:{col}{end}'.format(col=col_letters['total_surveys_entered'], end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",
            'mid_color': '#FFFFA8',
            'max_color': '#FF7EFF',
        })  # total_surveys_entered

        worksheet.conditional_format('{col}2:{col}{end}'.format(col=col_letters['CompLOI'], end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",
            'mid_color': '#FFFFA8',
            'max_color': "#FF7EFF",
        })  # CompLOI

        # Add conditional formatting for Tenure column (same as CompLOI)
        worksheet.conditional_format('{col}2:{col}{end}'.format(col=col_letters['Tenure'], end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",
            'mid_color': '#FFFFA8',
            'max_color': "#FF7EFF",
        })  # Tenure

        # Conditional formatting for Flag_Count column: red to white, max value 5
        worksheet.conditional_format('{col}2:{col}{end}'.format(col=col_letters['Flag_Count'], end=len(df_out) + 1), {
            'type': '2_color_scale',
            'max_color': "#FF0000",  # Red
            'min_color': "#E5E0EC",  # White
            'max_type': 'num',
            'min_value': 0,
            'max_value': 7
        })  # Flag_Count

        # Conditional formatting for font colors
        worksheet.conditional_format('BP2:BP{end}'.format(end=len(df_out) + 1), {
            'type': 'formula',
            'criteria': '=$BP2=TRUE',
            'format': workbook.add_format({'font_color': "#666666"})  # Yellow for TRUE
        })  # No_Enough_Data

        worksheet.conditional_format('BR2:BR{end}'.format(end=len(df_out) + 1), {
            'type': 'formula',
            'criteria': '=$BR2="Recent User, No Enough Data"',
            'format': workbook.add_format({'font_color': "#666666"})  # Yellow for "Recent User, No Enough Data"
        })  # PrioFlag - Recent User, No Enough Data

        worksheet.conditional_format('BR2:BR{end}'.format(end=len(df_out) + 1), {
            'type': 'formula',
            'criteria': '=$BR2="No Flags"',
            'format': workbook.add_format({'font_color': '#026102'})  # Green for "No Flags"
        })  # PrioFlag - No Flags

        # Conditional formatting for 3-color scale font colors
        worksheet.conditional_format('U2:U{end}'.format(end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",  # Blue
            'mid_color': '#FFFFA8',  # Yellow
            'max_color': '#FF7EFF',  # Purple
        })  # supplier_bu_id

        worksheet.conditional_format('AW2:AW{end}'.format(end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",  # Blue
            'mid_color': '#FFFFA8',  # Yellow
            'max_color': '#FF7EFF',  # Purple
        })  # total_surveys_entered

        worksheet.conditional_format('BT2:BT{end}'.format(end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",  # Blue
            'mid_color': '#FFFFA8',  # Yellow
            'max_color': '#FF7EFF',  # Purple
        })  # entrydate_split

        worksheet.conditional_format('BB2:BB{end}'.format(end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",  # Blue
            'mid_color': '#FFFFA8',  # Yellow
            'max_color': "#FF7EFF",  # Purple
        })  # CompLOI

        worksheet.conditional_format('H2:H{end}'.format(end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",  # Blue
            'mid_color': '#FFFFA8',  # Yellow
            'max_color': '#FF7EFF',  # Purple
        })  # client_responsestatusid

        worksheet.conditional_format('J2:J{end}'.format(end=len(df_out) + 1), {
            'type': '3_color_scale',
            'min_color': "#7C7CFC",  # Blue
            'mid_color': '#FFFFA8',  # Yellow
            'max_color': '#FF7EFF',  # Purple
        })  # fulcrum_responsestatusid



        # After writing Combined Data, call pivot sheet functions
        import pivot_sheet1
        import pivot_sheet2
        import pivot_sheet3
        import pivot_sheet4
        import pivot_sheet5
        import pivot_sheet6
        if config is None:
            config = {
                'surveys_entered_threshold': surveys_entered_threshold,
                'conversion_rate_threshold': conversion_rate_threshold,
                'security_terms_threshold': security_terms_threshold,
                'negative_recs_rate_threshold': negative_recs_rate_threshold,
                'is_pid_only_mode': is_pid_only_mode,
                'debug': True
            }
        # Call each pivot sheet function
        pivot_sheet1.write_flag_pivot(workbook, df_out, config)
        pivot_sheet2.write_prioflag_pivot(workbook, df_out, config)        
        pivot_sheet3.write_multiflag_pivot(workbook, df_out, config)
        pivot_sheet4.write_entrydateflags_pivot(workbook, df_out, config)
        pivot_sheet5.write_entrydatesuppliers_pivot(workbook, df_out, config)
        pivot_sheet6.write_denylist_draft(workbook, df_out, config)
        # Explicitly call sheet reordering
        # pivot_sheets.reorder_pivot_sheets(workbook, config)