# AI Coding Instructions for MPproV1

## Project Overview
Flask-based survey data processor that analyzes RID lookup (CSV) and PID metrics (Excel) files to generate flagged observations in multi-sheet Excel reports. Supports dual deployment: standalone executable via PyInstaller and Google App Engine.

## Architecture & Data Flow

### Core Components
- **`main.py`**: Flask app with single route handling file uploads, form validation, and report generation
- **`survey_processor.py`**: Pure business logic for data processing, Excel generation, and observation flagging
- **`run.py`**: Local development server with auto-browser opening
- **`templates/index.html`**: Single-page form interface with dynamic UI sections

### Processing Modes
1. **RID+PID Mode**: Merges RID lookup CSV with PID metrics Excel on `pid` column
2. **PID-Only Mode**: Processes only PID metrics file (checkbox or auto-detected)

### Key Data Processing Pipeline
```
CSV/Excel Upload → Stream Processing → Pandas Merge → Observation Logic → Multi-Sheet Excel Output
```

## Critical Business Logic

### Observation Flagging System (survey_processor.py:680-720)
Six flag types applied in sequence, with later flags overwriting `Observation` column:
1. **Poor_Conv_Rate**: `system_conversion_rate < threshold%`
2. **New_User_Bot**: `first_entry_date_time == last_entry_date_time` (datetime mode) or date-only comparison
3. **High_Security**: `sum_f_and_g_column/total_system_entrants > threshold%`
4. **Speeder**: `session_loi < actual_loi/multiplier` (RID+PID mode only)
5. **High_LOI**: `session_loi > actual_loi*multiplier` (RID+PID mode only)
6. **High_RR**: `negative_recs_rate > threshold%`

### Excel Output Structure
- **Combined Data**: Merged dataset with calculated columns
- **Flags Pivot (Priority)**: Suppliers × observations with conditional formatting
- **Flags Pivot (Multi)**: Multi-flag analysis by supplier
- **Pivot EntryDate × Supplier/Flags**: Time-series analysis
- **DenyList_Draft**: Filtered flagged records for review

## Development Patterns

### File Processing Convention
Always use stream-based processing: `file_storage.stream.seek(0)` before each operation. Files are temporarily saved to `UPLOAD_FOLDER` but processed via streams for GAE compatibility.

### Column Handling Pattern
```python
# Always normalize column names
df.columns = df.columns.str.strip().str.lower()
# Check existence before processing
if 'column_name' in df.columns:
    # Process column
```

### Excel Formatting Approach
Consistent pattern across all pivot sheets:
1. Calculate exclusion columns (index, -n/a-, totals)
2. Apply conditional formatting (red color scale)
3. Style -n/a- columns in dark green (`Font(color="006400")`)
4. Set supplier_bu column width to 200px

### Error Handling Strategy
- Form validation with flash messages
- Stream processing with try/catch and cleanup
- Graceful degradation (skip missing columns/sheets)

## Key Conventions

### File Structure
- `requirements.txt`: Minimal dependencies (Flask, pandas, openpyxl, gunicorn)
- `app.yaml`: GAE configuration with F1 instances, auto-scaling 0-2
- `RIDPIDProcessor.spec`: PyInstaller config including templates/static
- `Example inputs/`: Sample CSV/Excel files for testing

### Configuration Patterns
- Thresholds as form inputs with sensible defaults
- Boolean flags via checkbox presence in `request.form`
- Environment-specific temp directories (`tempfile.gettempdir()`)

### Sheet Reordering Logic
Always move "Flags Pivot (Multi)" after "Flags Pivot (Priority)" using openpyxl's `_sheets` manipulation.

## Local Development
```bash
python run.py  # Auto-opens browser to localhost:5000
```

## Building Executable
```bash
pyinstaller RIDPIDProcessor.spec  # Creates dist/RIDPIDProcessor.exe
```

## Testing Approach
Use files in `Example inputs/` directory. Test both processing modes with various threshold combinations. No automated test suite exists.

## Common Pitfalls
- Column name mismatches (case-sensitive)
- Stream position not reset between operations
- Missing null checks in boolean mask operations
- Excel formula references when moving sheets
