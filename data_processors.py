# data_processors.py - Core data processing and business logic
import pandas as pd

def read_rid_file_from_stream(file_stream):
    """Reads the RID lookup CSV file from a file stream."""
    try:
        df = pd.read_csv(file_stream)
        # Always normalize column names (convert to string first to handle integers)
        df.columns = [str(col) for col in df.columns]
        df.columns = df.columns.str.strip().str.lower()
        # Remove columns with blank headers
        df = df.loc[:, df.columns != '']
        
        # Validate file structure
        if df.empty:
            raise ValueError("RID file appears to be empty. Please check your CSV file.")
        
        # Check for required columns using exact names from specification
        required_cols = ['pid', 'supplier_bu', 'surveyid']
        missing_cols = [col for col in required_cols if col not in df.columns]
        if missing_cols:
            available_cols = ", ".join(df.columns[:10])
            raise ValueError(f"RID file missing required columns: {missing_cols}. Available columns: {available_cols}")
        
        # Convert PID to string immediately after reading and validation
        if 'pid' in df.columns:
            df['pid'] = df['pid'].apply(lambda x: str(x).strip() if pd.notnull(x) else '')
        
        return df
    except pd.errors.EmptyDataError:
        raise ValueError("RID file is empty or contains no valid data.")
    except pd.errors.ParserError as e:
        raise ValueError(f"RID file format error: Could not parse CSV file. Please ensure it's a valid CSV format. Details: {str(e)}")
    except UnicodeDecodeError:
        raise ValueError("RID file encoding error: Please ensure the CSV file is saved with UTF-8 encoding.")
    except Exception as e:
        raise ValueError(f"Error reading RID file: {str(e)}")

def read_metrics_file_from_stream(file_stream):
    """Reads the Marketplace Metrics Excel file from a file stream."""
    try:
        df = pd.read_excel(
            file_stream,
            sheet_name='Marketplace Metrics by PID',
            skiprows=5
        )
        # Always normalize column names (convert to string first to handle integers)
        df.columns = [str(col) for col in df.columns]
        df.columns = df.columns.str.strip().str.lower()
        # Remove columns with blank headers
        df = df.loc[:, df.columns != '']
        
        # Validate file structure
        if df.empty:
            raise ValueError("PID Metrics file appears to be empty after skipping header rows. Please check your Excel file structure.")
            
        # Check for required columns using exact names from specification
        required_cols = [
            'pid', 'total_system_entrants', 'total_surveys_entered', 'total_completes',
            'total_negative_recs', 'total_security_terms_on_marketplace_side',
            'total_security_terms_on_client_side', 'total security terms',
            'security terms rate', 'first_entry_date', 'last_entry_date',
            'first_entry_time', 'last_exit_time', 'negative_recs_rate',
            'net recs rate', 'system_conversion_rate'
        ]
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            available_cols = ", ".join(df.columns[:10])
            raise ValueError(f"PID Metrics file missing required columns: {', '.join(missing_cols)}. Available columns: {available_cols}")
        
        # Convert PID to string immediately after reading and validation
        if 'pid' in df.columns:
            df['pid'] = df['pid'].apply(lambda x: str(x).strip() if pd.notnull(x) else '')
        return df
    except FileNotFoundError:
        raise ValueError("Could not find the specified sheet 'Marketplace Metrics by PID' in the Excel file.")
    except ValueError as ve:
        if "Worksheet named" in str(ve):
            raise ValueError("Excel file must contain a sheet named 'Marketplace Metrics by PID'. Please check your file format.")
        raise ve  # Re-raise ValueError with original message
    except Exception as e:
        if "xlrd" in str(e) or "openpyxl" in str(e):
            raise ValueError("Excel file format error: Please ensure you're uploading a valid .xlsx file exported from SSRS.")
        raise ValueError(f"Error reading PID Metrics file: {str(e)}")

def apply_pid_observation_logic(
    df,
    survey_loi_mapping=None,
    conversion_rate_threshold=10,
    security_terms_threshold=30,
    speeder_multiplier=3,
    high_loi_multiplier=3,
    negative_recs_rate_threshold=15,
    session_loi_checks=True,
    use_datetime_for_newuser=True  # Kept for compatibility but ignored
):
    """
    Applies all observation logic to the DataFrame in-place.
    If session_loi_checks is False, skips Speeder and High LOI checks.
    Note: use_datetime_for_newuser is ignored - always uses date-only comparison now.
    survey_loi_mapping: Dictionary mapping survey IDs to their respective LOI values
    """
    import pandas as pd
    
    # Validate required columns exist using exact names from specification
    required_base_cols = ['system_conversion_rate', 'net recs rate', 'total_surveys_entered']
    missing_cols = [col for col in required_base_cols if col not in df.columns]
    if missing_cols:
        available_cols = ", ".join(df.columns[:10])
        raise ValueError(f"Missing required columns for analysis: {missing_cols}. Available columns: {available_cols}")
    
    # Validate date columns using exact names
    date_cols = ['first_entry_date', 'last_entry_date']
    missing_dates = [col for col in date_cols if col not in df.columns]
    if missing_dates:
        raise ValueError(f"Missing required date columns: {missing_dates}. Please ensure your PID Metrics file contains these columns.")
    
    # Validate security analysis columns using exact names
    security_cols = ['total security terms', 'security terms rate', 'total_system_entrants']
    missing_security = [col for col in security_cols if col not in df.columns]
    if missing_security:
        raise ValueError(f"Missing required security analysis columns: {missing_security}. Please ensure your PID Metrics file contains these columns.")
    
    # Check for session LOI column if needed
    if session_loi_checks and 'session_loi' not in df.columns:
        raise ValueError("Missing 'session_loi' column required for Speeder and High LOI analysis. Please ensure your files are merged correctly.")
    
    # Clean and prepare total_surveys_entered column (handle NULL/NaN/non-numeric as 0)
    df['total_surveys_entered'] = pd.to_numeric(df['total_surveys_entered'], errors='coerce').fillna(0)
    
    # Set default value
    df["Observation"] = "-n/a-"
    
    # Initialize check columns as False
    check_columns = [
        "Poor_Conv_Rate", "New_User_Bot", "High_Security",
        "Speeder", "High_LOI", "High_RR"
    ]
    
    for col in check_columns:
        df[col] = False
    
    # 1. Poor Conversion Rate (0-100 scale) - Only if total_surveys_entered >= 5
    mask_poor_conversion = (df["system_conversion_rate"] < conversion_rate_threshold) & \
                          (df['total_surveys_entered'] >= 5)
    df.loc[mask_poor_conversion.fillna(False), "Poor_Conv_Rate"] = True
    df.loc[mask_poor_conversion.fillna(False), "Observation"] = "Poor Conversion Rate"

    # 2. New User (bot?) - Always use date-only comparison now (no survey count condition)
    mask_new_user = (df['first_entry_date'] == df['last_entry_date']) & \
                   (df['first_entry_date'].notna()) & \
                   (df['last_entry_date'].notna())
    df.loc[mask_new_user.fillna(False), "New_User_Bot"] = True
    df.loc[mask_new_user.fillna(False), "Observation"] = "New User (bot?)"

    # 3. High Security Terms - Use pre-calculated rate, only if total_surveys_entered >= 5
    mask_high_security = (df['security terms rate'] > security_terms_threshold) & \
                        (df['total_surveys_entered'] >= 5)
    df.loc[mask_high_security.fillna(False), "High_Security"] = True
    df.loc[mask_high_security.fillna(False), "Observation"] = "High Security Terms"

    # 4. Speeder & 5. High LOI (no survey count condition) - Survey-specific LOI
    if session_loi_checks and survey_loi_mapping and speeder_multiplier and high_loi_multiplier:
        if 'surveyid' in df.columns:
            print(f"DEBUG: Using survey-specific LOI values for {len(survey_loi_mapping)} surveys")
            
            # Apply survey-specific Speeder and High LOI checks
            for idx, row in df.iterrows():
                survey_id = str(row['surveyid']).strip()
                if survey_id in survey_loi_mapping:
                    actual_loi = survey_loi_mapping[survey_id]
                    session_loi_val = row.get('session_loi')
                    
                    if pd.notna(session_loi_val):
                        # Speeder check
                        if session_loi_val < (actual_loi / speeder_multiplier):
                            df.loc[idx, "Speeder"] = True
                            df.loc[idx, "Observation"] = "Speeder"
                        
                        # High LOI check
                        elif session_loi_val > (actual_loi * high_loi_multiplier):
                            df.loc[idx, "High_LOI"] = True
                            df.loc[idx, "Observation"] = "High LOI, Distracted"
                else:
                    print(f"DEBUG: No LOI mapping found for survey ID: {survey_id}")
        else:
            print("DEBUG: No surveyid column found, skipping survey-specific LOI checks")
    else:
        print("DEBUG: Session LOI checks disabled or survey_loi_mapping not provided")

    # 6. High RR% - Use NET RECS RATE, only if total_surveys_entered >= 5
    mask_high_rr = (df['net recs rate'] > negative_recs_rate_threshold) & \
                   (df['total_surveys_entered'] >= 5)
    df.loc[mask_high_rr.fillna(False), "High_RR"] = True
    df.loc[mask_high_rr.fillna(False), "Observation"] = "High RR%"
    
    return df