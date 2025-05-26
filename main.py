# main.py
from flask import Flask, request, render_template, send_file, redirect, url_for, flash, session
import os
import tempfile # For handling temporary files on GAE
from werkzeug.utils import secure_filename # For secure filenames
import survey_processor # Your refactored logic
import pandas as pd # Added for server-side processing

app = Flask(__name__)
# IMPORTANT: Change this secret key in a production environment!
# You can generate one using: import secrets; secrets.token_hex(16)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "a_default_very_secret_key_for_dev")

# Register zip as a Jinja2 filter for template use
app.jinja_env.filters['zip'] = zip


# GAE Standard environment only allows writing to /tmp
# For more robust storage, especially for larger files or persistent needs, use Google Cloud Storage
# The UPLOAD_FOLDER will store the initially uploaded files temporarily.
# The OUTPUT_FOLDER will store the generated report temporarily before sending.
BASE_TEMP_DIR = tempfile.gettempdir() # Gets the appropriate temp directory (e.g., /tmp on Linux)
UPLOAD_FOLDER = os.path.join(BASE_TEMP_DIR, 'survey_uploads')
OUTPUT_FOLDER = os.path.join(BASE_TEMP_DIR, 'survey_outputs')

ALLOWED_EXTENSIONS_CSV = {'csv'}
ALLOWED_EXTENSIONS_XLSX = {'xlsx'}

# Ensure these directories exist when the app starts
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

def allowed_file(filename, allowed_extensions):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in allowed_extensions

@app.route('/', methods=['GET', 'POST'])
def index():
    top_surveyids_from_rid = []
    top_counts_from_rid = []
    
    if request.method == 'POST':
        if 'rid_file' not in request.files or 'metrics_file' not in request.files:
            flash('Both RID Lookup (CSV) and Marketplace Metrics (Excel) files are required!', 'error')
            return render_template('index.html', top_surveyids=top_surveyids_from_rid, top_counts=top_counts_from_rid)

        rid_file_storage = request.files.get('rid_file')
        metrics_file_storage = request.files.get('metrics_file')
        survey_loi_str = request.form.get('survey_loi', '10.0') # Default to 10.0 if not provided

        conversion_rate_threshold = float(request.form.get('conversion_rate_threshold', '10'))
        security_terms_threshold = float(request.form.get('security_terms_threshold', '30'))
        speeder_multiplier = float(request.form.get('speeder_multiplier', '3'))
        high_loi_multiplier = float(request.form.get('high_loi_multiplier', '4'))
        negative_recs_rate_threshold = float(request.form.get('negative_recs_rate_threshold', '15'))
        process_status_26_only = 'process_status_26_only' in request.form
        
        if rid_file_storage.filename == '' or metrics_file_storage.filename == '':
            flash('No selected file for one or both inputs. Please select both files.', 'error')
            # Try to get survey IDs if RID file was partially uploaded
            if rid_file_storage and rid_file_storage.filename:
                try:
                    rid_file_storage.stream.seek(0)
                    temp_df = pd.read_csv(rid_file_storage.stream)
                    rid_file_storage.stream.seek(0) # Reset stream
                    if 'surveyid' in temp_df.columns:
                        surveyid_counts = temp_df['surveyid'].value_counts()
                        top_surveyids_from_rid = list(surveyid_counts.index[:3])
                        top_counts_from_rid = list(surveyid_counts.values[:3])
                except Exception:
                    pass # Ignore if reading fails, already flashing error
            return render_template(
                'index.html',
                top_surveyids=top_surveyids_from_rid,
                top_counts=top_counts_from_rid
            )

        # Validate file extensions before saving (and before extensive parsing)
        if not (rid_file_storage and allowed_file(rid_file_storage.filename, ALLOWED_EXTENSIONS_CSV)):
            flash('Invalid RID Lookup file type. Must be a CSV file.', 'error')
            return render_template('index.html', top_surveyids=top_surveyids_from_rid, top_counts=top_counts_from_rid)

        if not (metrics_file_storage and allowed_file(metrics_file_storage.filename, ALLOWED_EXTENSIONS_XLSX)):
            flash('Invalid Marketplace Metrics file type. Must be an XLSX file.', 'error')
            # Try to get survey IDs if RID file was valid and uploaded
            if rid_file_storage and rid_file_storage.filename and allowed_file(rid_file_storage.filename, ALLOWED_EXTENSIONS_CSV):
                 try:
                    rid_file_storage.stream.seek(0)
                    temp_df = pd.read_csv(rid_file_storage.stream)
                    rid_file_storage.stream.seek(0)
                    if 'surveyid' in temp_df.columns:
                        surveyid_counts = temp_df['surveyid'].value_counts()
                        top_surveyids_from_rid = list(surveyid_counts.index[:3])
                        top_counts_from_rid = list(surveyid_counts.values[:3])
                 except Exception:
                    pass
            return render_template('index.html', top_surveyids=top_surveyids_from_rid, top_counts=top_counts_from_rid)
            
        try:
            actual_loi = float(survey_loi_str)
            if not (3 <= actual_loi <= 100):
                raise ValueError("Survey Actual LOI must be a number between 3 and 100.")
        except ValueError as e:
            flash(f"Invalid Survey Actual LOI: {e}", 'error')
            # Try to get survey IDs if RID file was uploaded
            if rid_file_storage and rid_file_storage.filename:
                try:
                    rid_file_storage.stream.seek(0)
                    temp_df = pd.read_csv(rid_file_storage.stream)
                    rid_file_storage.stream.seek(0)
                    if 'surveyid' in temp_df.columns:
                        surveyid_counts = temp_df['surveyid'].value_counts()
                        top_surveyids_from_rid = list(surveyid_counts.index[:3])
                        top_counts_from_rid = list(surveyid_counts.values[:3])
                except Exception:
                    pass
            return render_template(
                'index.html',
                top_surveyids=top_surveyids_from_rid,
                top_counts=top_counts_from_rid
            )

        rid_file_path = None
        metrics_file_path = None
        report_path_final = None

        try:
            # --- RID File Processing ---
            rid_file_storage.stream.seek(0)
            try:
                rid_df = pd.read_csv(rid_file_storage.stream)
            except Exception as e:
                flash(f"Error reading RID Lookup (CSV) file: {str(e)}", 'error')
                return render_template('index.html', top_surveyids=[], top_counts=[])
            rid_file_storage.stream.seek(0) # Reset stream for potential re-reads or survey_processor

            if 'surveyid' in rid_df.columns:
                surveyid_counts = rid_df['surveyid'].value_counts()
                top_surveyids_from_rid = list(surveyid_counts.index[:3])
                top_counts_from_rid = list(surveyid_counts.values[:3])
            else: # Ensure these exist even if surveyid column is missing
                top_surveyids_from_rid = []
                top_counts_from_rid = []
                flash("Warning: 'surveyid' column not found in RID file. Cannot display top Survey IDs for LOI context.", "warning")


            rid_pids = set(rid_df['pid'].unique()) if 'pid' in rid_df.columns else set()
            if not rid_pids:
                 flash("Warning: 'pid' column not found or empty in RID file.", "warning")


            has_status_26_rows = False
            status_26_pid_count = 0
            if 'client_responsestatusid' in rid_df.columns:
                status_26_mask = rid_df['client_responsestatusid'].astype(str) == '26'
                has_status_26_rows = status_26_mask.any()
                if has_status_26_rows:
                    status_26_pid_count = rid_df[status_26_mask]['pid'].nunique() if 'pid' in rid_df.columns else 0
            else:
                flash("Warning: 'client_responsestatusid' column not found in RID file. Status 26 check cannot be performed.", "warning")

            # --- Metrics File Processing ---
            metrics_file_storage.stream.seek(0)
            try:
                xls = pd.ExcelFile(metrics_file_storage.stream)
                sheet_name_to_read = None
                for name in xls.sheet_names:
                    if "Marketplace Metrics" in name:
                        sheet_name_to_read = name
                        break
                if not sheet_name_to_read:
                    sheet_name_to_read = xls.sheet_names[0] # Default to first sheet
                
                # Skip the first 6 rows (0-indexed)
                metrics_df = pd.read_excel(xls, sheet_name=sheet_name_to_read, skiprows=6)
            except Exception as e:
                flash(f"Error reading Marketplace Metrics (Excel) file: {str(e)}", 'error')
                return render_template('index.html', top_surveyids=top_surveyids_from_rid, top_counts=top_counts_from_rid)
            metrics_file_storage.stream.seek(0) # Reset stream

            metrics_pids = set(metrics_df['pid'].unique()) if 'pid' in metrics_df.columns else set()
            if not metrics_pids:
                flash("Warning: 'pid' column not found or empty in Metrics file (after skipping 6 header rows).", "warning")
                
            # --- Validations ---
            validation_failed = False
            if process_status_26_only and not has_status_26_rows:
                flash('No rows with status=26 found in the RID file. Uncheck "Process only status=26" or upload a valid file.', 'error')
                validation_failed = True

            # PID Comparison Warnings
            if rid_pids and metrics_pids: # Only compare if both PID sets are available
                pids_in_rid_not_in_metrics = rid_pids - metrics_pids
                if pids_in_rid_not_in_metrics:
                    flash(f"{len(pids_in_rid_not_in_metrics)} PIDs from RID file are not found in Metrics file.", 'warning')

                pids_in_metrics_not_in_rid = metrics_pids - rid_pids
                if pids_in_metrics_not_in_rid:
                    flash(f"{len(pids_in_metrics_not_in_rid)} PIDs from Metrics file are not found in RID file.", 'warning')
            
            if validation_failed:
                return render_template(
                    'index.html',
                    top_surveyids=top_surveyids_from_rid,
                    top_counts=top_counts_from_rid
                )

            # --- Save files temporarily (if all validations passed so far) ---
            # This is done after parsing and initial validation to avoid saving invalid/unreadable files
            rid_filename_secure = secure_filename(rid_file_storage.filename)
            rid_file_path = os.path.join(UPLOAD_FOLDER, rid_filename_secure)
            rid_file_storage.save(rid_file_path) # Save the original stream after reading from it

            metrics_filename_secure = secure_filename(metrics_file_storage.filename)
            metrics_file_path = os.path.join(UPLOAD_FOLDER, metrics_filename_secure)
            metrics_file_storage.save(metrics_file_path) # Save the original stream

            # --- Process files using survey_processor (pass streams again) ---
            rid_file_storage.stream.seek(0) # Ensure streams are reset before passing
            metrics_file_storage.stream.seek(0)
            
            report_path_final = survey_processor.generate_survey_report(
                rid_file_storage.stream, metrics_file_storage.stream, actual_loi, OUTPUT_FOLDER,
                conversion_rate_threshold=conversion_rate_threshold,
                security_terms_threshold=security_terms_threshold,
                speeder_multiplier=speeder_multiplier,
                high_loi_multiplier=high_loi_multiplier,
                negative_recs_rate_threshold=negative_recs_rate_threshold,
                process_status_26_only=process_status_26_only
            )
            
            flash('Processing complete! Your download should start automatically.', 'success')
            session['just_processed'] = True
            # Note: top_surveyids_from_rid and top_counts_from_rid are already set from above
            return send_file(
                report_path_final,
                as_attachment=True,
                download_name=os.path.basename(report_path_final)
            )
        except ValueError as ve: # Handles LOI validation from survey_processor or other ValueErrors
            app.logger.error(f"Validation Error: {ve}")
            flash(f"Error: {str(ve)}", 'error')
            # top_surveyids_from_rid and top_counts_from_rid might be populated from earlier try block
            return render_template(
                'index.html',
                top_surveyids=top_surveyids_from_rid,
                top_counts=top_counts_from_rid
            )
        except Exception as e:
            app.logger.error(f"An unexpected error occurred during processing: {e}", exc_info=True)
            flash(f"An unexpected error occurred: {str(e)}. Please check logs or try again.", 'error')
            # top_surveyids_from_rid and top_counts_from_rid might be populated from earlier try block
            return render_template(
                'index.html',
                top_surveyids=top_surveyids_from_rid,
                top_counts=top_counts_from_rid
            )
        finally:
            # Clean up uploaded files from UPLOAD_FOLDER
            if rid_file_path and os.path.exists(rid_file_path):
                try:
                    os.remove(rid_file_path)
                except Exception as e_clean: # pragma: no cover
                    app.logger.error(f"Error cleaning up RID file {rid_file_path}: {e_clean}")
            
            if metrics_file_path and os.path.exists(metrics_file_path):
                try:
                    os.remove(metrics_file_path)
                except Exception as e_clean: # pragma: no cover
                     app.logger.error(f"Error cleaning up Metrics file {metrics_file_path}: {e_clean}")
            
            # Output file cleanup is handled by send_file or OS for /tmp

    # GET request or if POST processing leads here without sending a file
    # Flashed messages are automatically available to the template context if not consumed
    
    # Handle session pop for 'just_processed'
    if session.pop('just_processed', None):
        # Messages will be displayed by the template normally.
        # If the intention was to clear messages specifically on a refresh after processing,
        # that behavior will change. Now, any flashed messages (like 'Processing complete!')
        # will be shown on the next render, which is generally desirable.
        # If truly no messages should be shown after 'just_processed', then main.py should not flash them in the first place
        # for that specific scenario, or a different mechanism is needed.
        # For now, the goal is to fix the UnboundLocalError and ensure messages *can* be shown.
        return render_template('index.html', top_surveyids=[], top_counts=[])

    # For a regular GET request, or if POST failed and redirected to GET,
    # or if 'just_processed' was not set.
    # top_surveyids_from_rid and top_counts_from_rid will be empty if it's an initial GET.
    # If it's after a POST that failed and rendered template directly, they might have values.
    # However, the standard pattern is POST -> flash -> redirect -> GET for success,
    # and POST -> flash -> render for errors.
    # The 'messages' variable is not explicitly passed anymore.
    # The template will call get_flashed_messages() directly.
    return render_template(
        'index.html',
        top_surveyids=top_surveyids_from_rid, 
        top_counts=top_counts_from_rid
    )

if __name__ == "__main__":
    app.run(debug=True)