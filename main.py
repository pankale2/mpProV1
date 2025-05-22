# main.py
from flask import Flask, request, render_template, send_file, redirect, url_for, flash, session
import os
import tempfile # For handling temporary files on GAE
from werkzeug.utils import secure_filename # For secure filenames
import survey_processor # Your refactored logic

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

def get_top_surveyids_from_file(file_storage):
    import pandas as pd
    try:
        file_storage.stream.seek(0)
        df = pd.read_csv(file_storage.stream)
        file_storage.stream.seek(0)
        if 'surveyid' in df.columns:
            surveyid_counts = df['surveyid'].value_counts()
            top_surveyids = list(surveyid_counts.index[:3])
            top_counts = list(surveyid_counts.values[:3])
            return top_surveyids, top_counts
    except Exception:
        pass
    return [], []

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'rid_file' not in request.files or 'metrics_file' not in request.files:
            flash('Both RID Lookup (CSV) and Marketplace Metrics (Excel) files are required!', 'error')
            return render_template('index.html', messages=[], top_surveyids=[], top_counts=[])

        rid_file_storage = request.files.get('rid_file')
        metrics_file_storage = request.files.get('metrics_file')
        survey_loi_str = request.form.get('survey_loi', '10.0') # Default to 10.0 if not provided

        # New: Get thresholds from form, with defaults
        conversion_rate_threshold = float(request.form.get('conversion_rate_threshold', '10'))
        security_terms_threshold = float(request.form.get('security_terms_threshold', '30'))
        speeder_multiplier = float(request.form.get('speeder_multiplier', '3'))
        high_loi_multiplier = float(request.form.get('high_loi_multiplier', '4'))
        negative_recs_rate_threshold = float(request.form.get('negative_recs_rate_threshold', '15'))
        process_status_26_only = 'process_status_26_only' in request.form
        
        if rid_file_storage.filename == '' or metrics_file_storage.filename == '':
            top_surveyids, top_counts = get_top_surveyids_from_file(rid_file_storage) if rid_file_storage and rid_file_storage.filename else ([], [])
            flash('No selected file for one or both inputs. Please select both files.', 'error')
            return render_template(
                'index.html',
                messages=[],
                top_surveyids=top_surveyids,
                top_counts=top_counts
            )

        try:
            actual_loi = float(survey_loi_str)
            if not (3 <= actual_loi <= 100):
                # This validation is also in survey_processor, but good to have early feedback
                raise ValueError("Survey Actual LOI must be a number between 3 and 100.")
        except ValueError as e:
            top_surveyids, top_counts = get_top_surveyids_from_file(rid_file_storage) if rid_file_storage and rid_file_storage.filename else ([], [])
            flash(f"Invalid Survey Actual LOI: {e}", 'error')
            return render_template(
                'index.html',
                messages=[],
                top_surveyids=top_surveyids,
                top_counts=top_counts
            )

        # Secure filenames and save uploaded files temporarily
        rid_file_path = None
        metrics_file_path = None
        report_path_final = None

        try:
            if rid_file_storage and allowed_file(rid_file_storage.filename, ALLOWED_EXTENSIONS_CSV):
                rid_filename_secure = secure_filename(rid_file_storage.filename)
                rid_file_path = os.path.join(UPLOAD_FOLDER, rid_filename_secure)
                rid_file_storage.save(rid_file_path)
            else:
                flash('Invalid RID Lookup file type. Please upload a CSV file.', 'error')
                return redirect(request.url)

            if metrics_file_storage and allowed_file(metrics_file_storage.filename, ALLOWED_EXTENSIONS_XLSX):
                metrics_filename_secure = secure_filename(metrics_file_storage.filename)
                metrics_file_path = os.path.join(UPLOAD_FOLDER, metrics_filename_secure)
                metrics_file_storage.save(metrics_file_path)
            else:
                flash('Invalid Marketplace Metrics file type. Please upload an XLSX file.', 'error')
                return redirect(request.url)

            # --- Process files using streams/paths ---
            rid_file_storage.stream.seek(0)
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

            # Read surveyids for UI (top 3 by count)
            rid_file_storage.stream.seek(0)
            import pandas as pd
            try:
                rid_df_preview = pd.read_csv(rid_file_storage.stream)
            except Exception as e:
                rid_df_preview = pd.DataFrame()
            rid_file_storage.stream.seek(0)
            surveyid_counts = (
                rid_df_preview['surveyid'].value_counts()
                if 'surveyid' in rid_df_preview.columns else pd.Series(dtype=int)
            )
            top_surveyids = list(surveyid_counts.index[:3])
            top_counts = list(surveyid_counts.values[:3])

            # Check for status=26 rows if checkbox is enabled
            if process_status_26_only:
                rid_file_storage.stream.seek(0)
                try:
                    rid_df_check = pd.read_csv(rid_file_storage.stream)
                except Exception as e:
                    rid_df_check = pd.DataFrame()
                rid_file_storage.stream.seek(0)
                if 'client_responsestatusid' in rid_df_check.columns:
                    has_26 = (rid_df_check['client_responsestatusid'].astype(str) == '26').any()
                    if not has_26:
                        flash('No rows with status=26 found in the uploaded RID file. Uncheck "Process only status=26" to proceed with all rows.', 'error')
                        from flask import get_flashed_messages
                        return render_template(
                            'index.html',
                            messages=get_flashed_messages(with_categories=True),
                            top_surveyids=top_surveyids,
                            top_counts=top_counts
                        )

            flash('Processing complete! Your download should start automatically.', 'success')
            session['just_processed'] = True
            return send_file(
                report_path_final,
                as_attachment=True,
                download_name=os.path.basename(report_path_final)
            )
        except ValueError as ve:
            app.logger.error(f"Validation Error: {ve}")
            top_surveyids, top_counts = get_top_surveyids_from_file(rid_file_storage) if rid_file_storage and rid_file_storage.filename else ([], [])
            flash(f"Error: {str(ve)}", 'error')
            return render_template(
                'index.html',
                messages=[],
                top_surveyids=top_surveyids,
                top_counts=top_counts
            )
        except Exception as e:
            app.logger.error(f"An unexpected error occurred during processing: {e}", exc_info=True)
            top_surveyids, top_counts = get_top_surveyids_from_file(rid_file_storage) if rid_file_storage and rid_file_storage.filename else ([], [])
            flash(f"An unexpected error occurred: {str(e)}. Please check logs or try again.", 'error')
            return render_template(
                'index.html',
                messages=[],
                top_surveyids=top_surveyids,
                top_counts=top_counts
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
            
            # The generated report in OUTPUT_FOLDER will be sent by send_file.
            # GAE's /tmp directory is ephemeral, so explicit cleanup of output might not be strictly necessary
            # but can be done if you want to be thorough or if send_file doesn't clean up.
            # For this example, we assume the OS or GAE environment handles /tmp cleanup.

        return redirect(request.url)

    # Only show messages if not a refresh after processing
    if session.pop('just_processed', None):
        from flask import get_flashed_messages
        get_flashed_messages()
        return render_template('index.html', messages=[], top_surveyids=[], top_counts=[])

    # GET: no file uploaded yet, so pass empty lists
    return render_template('index.html', top_surveyids=[], top_counts=[])

if __name__ == "__main__":
    app.run(debug=True)