# controller.py - Main orchestrator and public API for report generation
import pandas as pd
from datetime import datetime
import os
import tempfile
import re

# Import the new modular components - FIXED PATHS
from data_processors import read_rid_file_from_stream, read_metrics_file_from_stream, apply_pid_observation_logic
from excel_generators import write_combined_data_xlsx

def generate_survey_report(
    rid_file_stream, metrics_file_stream, survey_loi_mapping, output_dir,
    conversion_rate_threshold=10,
    security_terms_threshold=30,
    speeder_multiplier=3,
    high_loi_multiplier=3,
    negative_recs_rate_threshold=15,
    surveys_entered_threshold=5,
    is_pid_only_mode=False,  # Added parameter
    is_average_loi_mode=False,
    average_loi_value=None
):
    """
    Processes the survey files and generates an Excel report with observations and a pivot table.
    Note: use_datetime_for_newuser is ignored - always uses date-only comparison now.
    survey_loi_mapping: Dictionary mapping survey IDs to their respective LOI values
    """
    try:
        # Validate survey_loi_mapping
        # Fix: Only require survey_loi_mapping if not in average LOI mode
        if not is_average_loi_mode:
            if not survey_loi_mapping or not isinstance(survey_loi_mapping, dict):
                raise ValueError("Survey LOI mapping is required and must be a dictionary of survey_id: loi_value pairs.")
            
            # Validate LOI values
            for survey_id, loi_value in survey_loi_mapping.items():
                if not (3 <= loi_value <= 100):
                    raise ValueError(f"Survey LOI for {survey_id} must be between 3 and 100, got {loi_value}.")

        # Read files with enhanced error handling
        try:
            rid_df = read_rid_file_from_stream(rid_file_stream)
        except Exception as e:
            raise ValueError(f"RID file error: {str(e)}")
            
        try:
            metrics_df = read_metrics_file_from_stream(metrics_file_stream)
        except Exception as e:
            raise ValueError(f"PID Metrics file error: {str(e)}")

        # Debug: Print shapes and columns of input DataFrames
        print("DEBUG: rid_df shape:", rid_df.shape)
        print("DEBUG: rid_df columns:", rid_df.columns.tolist())
        print("DEBUG: metrics_df shape:", metrics_df.shape)
        print("DEBUG: metrics_df columns:", metrics_df.columns.tolist())

        if 'pid' not in rid_df.columns or 'pid' not in metrics_df.columns:
            raise ValueError("Critical Error: 'pid' column not found in one or both input files. Please check that files contain 'pid' column.")
        
        # Enhanced merge validation - PIDs are already strings from file readers
        rid_pids = rid_df['pid']
        metrics_pids = metrics_df['pid']
        pid_intersection = set(rid_pids) & set(metrics_pids)
        print("DEBUG: Number of PIDs in RID file:", len(rid_pids))
        print("DEBUG: Number of PIDs in Metrics file:", len(metrics_pids))
        print("DEBUG: Number of matching PIDs:", len(pid_intersection))

        if len(pid_intersection) == 0:
            raise ValueError("No matching PIDs found between RID and Metrics files. Please ensure both files contain the same PIDs.")

        # PIDs are already converted to strings in the file readers
        merged_df = pd.merge(
            rid_df,
            metrics_df,
            on='pid',
            how='left',
            indicator=False
        )

        # Debug: Print merged DataFrame shape and columns after merge
        print("DEBUG: merged_df shape:", merged_df.shape)
        print("DEBUG: merged_df columns:", merged_df.columns.tolist())

        # Add this check immediately after merge
        if merged_df.empty:
            raise ValueError(
                "Merge resulted in an empty DataFrame. "
                "No matching rows found between RID and PID files after merge. "
                "Check that your files have overlapping PIDs and correct formats."
            )

        # DEBUG: Check for security terms rate
        print("DEBUG: merged_df columns after merge:", merged_df.columns.tolist())
        print("DEBUG: sample security terms rate:", merged_df['security terms rate'].head() if 'security terms rate' in merged_df.columns else "not found")

        # Remove scaling logic for percentage columns
        # Ensure all four columns are in 0-100 scale

        # Remove rounding for percent columns before writing to Excel
        percent_cols = [
            "system_conversion_rate",
            "Security_Terms_Rate", 
            "negative_recs_rate",
            "net recs rate"
        ]
        for col in percent_cols:
            if col in merged_df.columns:
                merged_df[col] = pd.to_numeric(merged_df[col], errors='coerce')
                # REMOVED: No rounding

        # REMOVED: Do not divide thresholds - use them directly as provided from UI
        # conversion_rate_threshold = conversion_rate_threshold / 100
        # security_terms_threshold = security_terms_threshold / 100  
        # negative_recs_rate_threshold = negative_recs_rate_threshold / 100

        # Apply observation logic and add check columns
        merged_df = apply_pid_observation_logic(
            merged_df,
            survey_loi_mapping=survey_loi_mapping,
            conversion_rate_threshold=conversion_rate_threshold,
            security_terms_threshold=security_terms_threshold,
            speeder_multiplier=speeder_multiplier,
            high_loi_multiplier=high_loi_multiplier,
            negative_recs_rate_threshold=negative_recs_rate_threshold,
            session_loi_checks=True,
            use_datetime_for_newuser=False,
            surveys_entered_threshold=surveys_entered_threshold,
            is_average_loi_mode=is_average_loi_mode,
            average_loi_value=average_loi_value
        )

        # --- Rename 'Observation' to 'PrioFlag' immediately after observation logic ---
        if 'Observation' in merged_df.columns:
            merged_df = merged_df.rename(columns={'Observation': 'PrioFlag'})

        # --- Move Security_Terms_Rate to between net recs rate and system_conversion_rate ---
        cols = list(merged_df.columns)
        if 'Security_Terms_Rate' in cols and 'net recs rate' in cols and 'system_conversion_rate' in cols:
            sec_val = merged_df['Security_Terms_Rate']
            merged_df.drop('Security_Terms_Rate', axis=1, inplace=True)
            cols = list(merged_df.columns)
            net_idx = cols.index('net recs rate')
            merged_df.insert(net_idx + 1, 'Security_Terms_Rate', sec_val)

        # --- Insert Diff Days column after last_entry_date using exact column names ---
        if 'first_entry_date' in merged_df.columns and 'last_entry_date' in merged_df.columns:
            # Convert to datetime if not already
            merged_df["first_entry_date_dt"] = pd.to_datetime(merged_df['first_entry_date'], errors='coerce')
            merged_df["last_entry_date_dt"] = pd.to_datetime(merged_df['last_entry_date'], errors='coerce')
            diff_days = (merged_df["last_entry_date_dt"] - merged_df["first_entry_date_dt"]).dt.days
            # Insert after last_entry_date
            cols = list(merged_df.columns)
            if 'last_entry_date' in cols:
                idx = cols.index('last_entry_date') + 1
                merged_df.insert(idx, "Diff Days", diff_days)
            else:
                merged_df["Diff Days"] = diff_days
            # Remove temp columns
            merged_df.drop(["first_entry_date_dt", "last_entry_date_dt"], axis=1, inplace=True)
        else:
            merged_df["Diff Days"] = None

        # --- Round columns to 2 decimal places using exact column names ---
        # REMOVED: All rounding logic for percentage columns
        # if "system_conversion_rate" in merged_df.columns:
        #     merged_df["system_conversion_rate"] = merged_df["system_conversion_rate"].round(2)
        # if 'net recs rate' in merged_df.columns:
        #     merged_df['net recs rate'] = merged_df['net recs rate'].round(2)
        # if "Security_Terms_Rate" in merged_df.columns:
        #     merged_df["Security_Terms_Rate"] = merged_df["Security_Terms_Rate"].round(2)

        # --- Add Flag_Count column at the end ---
        flag_columns = [
            "Poor_Conv_Rate", "New_User_Bot", "High_Security",
            "Speeder", "High_LOI", "High_RR"
        ]
        merged_df["Flag_Count"] = merged_df[flag_columns].sum(axis=1)

        # Remove blank columns (all values are NaN or empty) before writing to Excel
        merged_df = merged_df.dropna(axis=1, how='all')

        # Add user input columns at the end
        merged_df['conversion_rate_threshold'] = conversion_rate_threshold
        merged_df['security_terms_threshold'] = security_terms_threshold
        merged_df['speeder_multiplier'] = speeder_multiplier
        merged_df['high_loi_multiplier'] = high_loi_multiplier
        merged_df['negative_recs_rate_threshold'] = negative_recs_rate_threshold
        merged_df['surveys_entered_threshold'] = surveys_entered_threshold

        # Drop unwanted unnamed columns from output
        drop_cols = ['unnamed: 0', 'unnamed: 1', 'unnamed: 3', 'unnamed: 15']
        merged_df.drop(columns=[c for c in drop_cols if c in merged_df.columns], inplace=True, errors='ignore')

        # Remove columns not needed in output
        for col_to_remove in ['New_User_Less_History']:
            if col_to_remove in merged_df.columns:
                merged_df.drop(columns=[col_to_remove], inplace=True, errors='ignore')

        # --- Fixed column order for Combined Data sheet ---
        combined_data_columns = [
            'rid', 'buyer_account_id', 'buyer_account', 'buyer_bu', 'buyer_bu_id', 'survey_client',
            'client_responsestatusid', 'client_responsestatus', 'link_type_id', 'external_survey_name',
            'fulcrum_responsestatusid', 'fulcrum_responsestatus', 'internal_survey_name', 'marketplace_projectid',
            'marketplace_project', 'mid', 'parentsid', 'pid', 'respondentsid', 'entrydate', 'lastdate', 'id', 'name',
            'supplier_bu_id', 'link_type', 'supplierid', 'survey_country', 'survey_country_langauge', 'survey_ccpi',
            'survey_HASH_status', 'survey_SCCB_status', 'survey_https_status', 'project_manager', 'pm_email',
            'survey_qcpi', 'total_system_entrants', 'total_completes', 'total_negative_recs',
            'total_security_terms_on_marketplace_side', 'total_security_terms_on_client_side', 'total security terms',
            'first_entry_time', 'last_exit_time', 'net recs rate', 'supplier_bu', 'first_entry_date', 'last_entry_date',
            'diff days', 'total_surveys_entered', 'system_conversion_rate', 'security terms rate', 'negative_recs_rate',
            'surveyid', 'comploi', 'session_loi', 'speeder_multiplier', 'high_loi_multiplier',
            'surveys_entered_threshold', 'conversion_rate_threshold', 'security_terms_threshold',
            'negative_recs_rate_threshold', 'speeder', 'high_loi', 'poor_conv_rate', 'high_security', 'new_user_bot',
            'high_rr', 'no enough data', 'flag_count', 'prioflag'
        ]

        # --- preserve original column names, but build a normalized lookup for matching ---
        def _normalize_col_name(s):
            if s is None:
                return ''
            return re.sub(r'[^0-9a-z]', '', str(s).lower())

        # Build normalized->actual column map once
        normalized_to_actual = { _normalize_col_name(c): c for c in merged_df.columns }

        # DEBUG: print normalized->actual mapping and canonical->actual resolution
        print("DEBUG: Normalized -> actual column map (sample):")
        # print a limited sample to avoid huge logs - but show full mapping length
        for k, v in list(normalized_to_actual.items())[:50]:
            print(f"  {k!r} -> {v!r}")
        print(f"DEBUG: Total normalized columns: {len(normalized_to_actual)}")

        # Also show how canonical required columns will map to actual columns
        canonical_checks = [
            'rid', 'buyer_account_id', 'pid', 'system_conversion_rate',
            'security terms rate', 'total_surveys_entered', 'Diff Days',
            'CompLOI', 'Flag_Count', 'PrioFlag'
        ]
        print("DEBUG: Canonical -> actual mapping preview:")
        for canon in canonical_checks:
            mapped = normalized_to_actual.get(_normalize_col_name(canon))
            print(f"  Canonical: {canon!r}  ->  Actual: {mapped!r}")

        # --- Ensure calculated columns are present (empty/null if missing) ---
        # We will map canonical names to actual column names where possible; otherwise create columns with None.
        canonical_calc_cols = ['diff days', 'comploi', 'flag_count', 'prioflag']
        for canon in canonical_calc_cols:
            norm = _normalize_col_name(canon)
            if norm in normalized_to_actual:
                # ensure the actual column exists (it already does) — nothing to do
                pass
            else:
                # create the canonical column name in merged_df so downstream logic finds it if it expects canonical names
                # We create the canonical name as provided (use the canonical string) with empty values
                merged_df[canon] = None
                # update mapping
                normalized_to_actual[norm] = canon

        # --- Omit missing columns from output, log warnings ---
        # Use canonical combined_data_columns (the list you already have) and map to actual column names in merged_df
        output_actual_cols = []
        missing_cols = []
        for canon in combined_data_columns:
            norm = _normalize_col_name(canon)
            if norm in normalized_to_actual:
                output_actual_cols.append(normalized_to_actual[norm])
            else:
                missing_cols.append(canon)

        # Special handling for 'security terms rate' - ensure it's included if present in merged_df
        if 'security terms rate' in merged_df.columns and 'security terms rate' not in output_actual_cols:
            output_actual_cols.append('security terms rate')
            print("DEBUG: Added 'security terms rate' to output_actual_cols")

        if missing_cols:
            print(f"WARNING: The following columns are missing and will be omitted from output: {missing_cols}")

        # Build final ordered DataFrame using actual column names found
        # Ensure duplicates are removed while preserving order
        seen = set()
        final_output_cols = []
        for col in output_actual_cols:
            if col not in seen:
                final_output_cols.append(col)
                seen.add(col)

        # Ensure the calculated canonical columns exist at the end (use the actual column names from normalized_to_actual)
        for calc in ['diff days', 'comploi', 'flag_count', 'prioflag']:
            actual = normalized_to_actual.get(_normalize_col_name(calc))
            if actual and actual not in seen:
                final_output_cols.append(actual)
                seen.add(actual)

        # Slice merged_df to final column order (missing columns were created earlier as None where needed)
        merged_df = merged_df[final_output_cols]

        # --- Write to Excel using xlsxwriter ---
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"RID-PID_Report_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        print("DEBUG: Writing Combined Data sheet with xlsxwriter...")
        write_combined_data_xlsx(
            output_path,
            merged_df,
            surveys_entered_threshold=surveys_entered_threshold,
            conversion_rate_threshold=conversion_rate_threshold,
            security_terms_threshold=security_terms_threshold,
            negative_recs_rate_threshold=negative_recs_rate_threshold,
            is_pid_only_mode=is_pid_only_mode  # Added parameter
        )
        print("DEBUG: Combined Data sheet written successfully")
        return str(output_path)
        
    except ValueError:
        print("DEBUG: Caught ValueError, re-raising...")
        raise  # Re-raise ValueError as-is
    except Exception as e:
        print(f"DEBUG: Caught unexpected exception: {e}")
        import traceback
        traceback.print_exc()
        raise ValueError(f"Unexpected error during report generation: {str(e)}")

def generate_pid_only_report(
    metrics_file_stream,
    output_dir,
    conversion_rate_threshold=10,
    security_terms_threshold=30,
    negative_recs_rate_threshold=15,
    surveys_entered_threshold=5,
    is_pid_only_mode=False  # Added parameter
):
    """Processes only the PID Metrics file and generates an Excel report with observations (PID-only mode)."""
    try:
        # Enhanced file reading with error handling
        try:
            metrics_df = read_metrics_file_from_stream(metrics_file_stream)
        except Exception as e:
            raise ValueError(f"PID Metrics file error: {str(e)}")
            
        if metrics_df.empty:
            raise ValueError("PID Metrics file contains no data rows. Please check your Excel file.")
        
        # Pass None for speeder_multiplier and high_loi_multiplier, and session_loi_checks=False
        metrics_df = apply_pid_observation_logic(
            metrics_df,
            survey_loi_mapping=None,
            conversion_rate_threshold=conversion_rate_threshold,
            security_terms_threshold=security_terms_threshold,
            speeder_multiplier=None,
            high_loi_multiplier=None,
            negative_recs_rate_threshold=negative_recs_rate_threshold,
            session_loi_checks=False,
            use_datetime_for_newuser=True,
            surveys_entered_threshold=surveys_entered_threshold  # <-- Pass this argument
        )
        
        # --- Generate Excel File ---
        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_filename = f"pid_metrics_report_{timestamp}.xlsx"
        output_path = os.path.join(output_dir, output_filename)
        print("DEBUG: Writing PID Metrics Data sheet with xlsxwriter...")
        write_combined_data_xlsx(
            output_path,
            metrics_df,
            surveys_entered_threshold=surveys_entered_threshold,
            conversion_rate_threshold=conversion_rate_threshold,
            security_terms_threshold=security_terms_threshold,
            negative_recs_rate_threshold=negative_recs_rate_threshold,
            is_pid_only_mode=is_pid_only_mode  # Added parameter
        )
        print("DEBUG: PID Metrics Data sheet written successfully")
        return str(output_path)
    except ValueError:
        raise  # Re-raise ValueError as-is
    except Exception as e:
        print(f"DEBUG: PID-only unexpected error: {e}")
        raise ValueError(f"Unexpected error in PID-only processing: {str(e)}")