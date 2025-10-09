# AI Coding Instructions for MPproV1

## Project Overview
Flask-based survey data processor that analyzes RID lookup (CSV) and PID metrics (Excel) files to generate flagged observations in multi-sheet Excel reports. Supports dual deployment: standalone executable via PyInstaller and Google App Engine.

## Architecture & Data Flow

### Core Components
- **`main.py`**: Flask app with routes for file uploads, form validation, and report generation. Integrates with the `controller.py` module for processing.
- **`controller.py`**: Orchestrates the data processing pipeline, including reading files, applying observation logic, and generating Excel reports.
- **`data_processors.py`**: Handles data ingestion and validation for RID and PID files. Implements core business logic for observation flagging.
- **`excel_generators.py`**: Responsible for creating and formatting Excel reports, including pivot tables and conditional formatting with comprehensive column mapping and Excel formulas.
- **`pivot_sheet1.py`**: Creates "Counts" sheet with TRANSPOSE formulas referencing pivot data.
- **`pivot_sheet2.py`**: Creates "Flags Pivot (Priority)" sheet with supplier × observations analysis.
- **`pivot_sheet3.py`**: Creates "Flags Pivot (Multi)" sheet with multi-flag analysis by supplier.
- **`pivot_sheet4.py`**: Creates "Pivot EntryDate × Flags" time-series analysis.
- **`pivot_sheet5.py`**: Creates "Pivot EntryDate × Supplier" time-series analysis.
- **`pivot_sheet51.py`**: Creates "TenureSupplier Pivot" sheet with supplier × tenure groups analysis.
- **`pivot_sheet6.py`**: Creates "DenyList_Draft" sheet for filtered flagged records.
- **`run.py`**: Development server that opens the application in a browser and manages session-specific flags.
- **`templates/index.html`**: Single-page form interface with dynamic UI sections, drag-and-drop file uploads, and LOI mode toggle.
- **`static/js/app.js`**: Handles frontend logic, including form validation, dynamic UI updates, LOI modes, drag-and-drop functionality, PID boxes display, and AJAX-based communication.
- **`static/js/file-manager.js`**: Manages all file operations, drag-drop zones, PID boxes with label wrapper, RID processing textboxes, and file previews.
- **`static/js/form-manager.js`**: Handles form state management, LOI input rendering, validation, and dark mode toggling.
- **`static/css/main.css`**: Comprehensive styling with dark mode support, PID boxes styling with wrapper/label, and responsive design.

### Processing Modes
**Default Mode**: RID+PID Mode - Merges RID lookup CSV with PID metrics Excel on `pid` column. All 6 observation checks are applied including session-based checks.

*Note: PID-only mode exists in the backend but the mode slider is hidden by default in the UI. The application runs in RID+PID mode requiring both files.*

### LOI Input Modes
1. **Survey-specific LOI**: Individual LOI values for each survey ID with marketplace links and RID counts.
2. **Average LOI**: Single LOI value applied to all surveys with aggregated marketplace link to top survey.

### Key Data Processing Pipeline
```
CSV/Excel Upload → Stream Processing → Pandas Merge → Observation Logic → Excel Formulas → Multi-Sheet Excel Output
```

## Critical Business Logic

### Observation Flagging System (data_processors.py)
Six flag types applied in sequence, with later flags overwriting `Observation` column:
1. **Poor_Conv_Rate**: `system_conversion_rate < threshold%` AND `total_surveys_entered > surveys_entered_threshold`
2. **New_User_Bot**: `first_entry_date == last_entry_date` AND `total_surveys_entered < 5` (date-only comparison)
3. **High_Security**: `security_terms_rate > threshold%` AND `total_surveys_entered > surveys_entered_threshold`
4. **Speeder**: `session_loi < actual_loi/multiplier` (RID+PID mode only)
5. **High_LOI**: `session_loi > actual_loi*multiplier` (RID+PID mode only)
6. **High_RR**: `net_recs_rate > threshold%` AND `total_surveys_entered > surveys_entered_threshold`

**Survey Count Condition**: Three flags (Poor_Conv_Rate, High_Security, High_RR) only apply when `total_surveys_entered > surveys_entered_threshold`. NULL/NaN/non-numeric values in this column are treated as 0.

### Excel Output Structure
- **Combined Data**: Merged dataset with calculated columns and Excel formulas for dynamic calculations, including user input thresholds repeated for all rows.
- **Counts**: Count tables with TRANSPOSE formulas referencing pivot data, including PrioFlag, MultiFlag, and general count tables.
- **Flags Pivot (Priority)**: Suppliers × observations with conditional formatting (priority-based single observation).
- **Flags Pivot (Multi)**: Multi-flag analysis by supplier showing all applicable flags.
- **Pivot EntryDate × Flags**: Time-series analysis of flags by date.
- **Pivot EntryDate × Supplier**: Time-series analysis of entries by supplier and date.
- **TenureSupplier Pivot**: Suppliers × tenure groups analysis with optimized conditional formatting.
- **DenyList_Draft**: Filtered flagged records for review with proper formatting.

### Excel Formula Architecture
Instead of Python-side calculations, most dynamic columns use Excel formulas:
- **Flag Columns**: Use IF/AND formulas referencing threshold columns with explicit column letter mapping
- **PrioFlag**: Nested IF statements determining priority observation
- **Tenure**: DATEDIF formula calculating days since first entry
- **Flag_Count**: COUNTIF formula counting TRUE values in flag columns
- **Tenure_Group**: IF statements categorizing tenure into groups
- **entrydate_split**: LEFT/FIND formula extracting date portion from datetime
- **Counts Sheet**: TRANSPOSE/XLOOKUP formulas referencing pivot sheets

## Frontend Logic

### Dynamic Form Handling
- **File Inputs**: Handles RID and PID file uploads with validation and dynamic UI updates using drag-and-drop zones.
- **Survey LOI Inputs**: Dynamically generates input fields for survey LOI values based on uploaded RID data. Validates inputs to ensure values are within the range of 3-100.
- **LOI Mode Toggle**: Switches between survey-specific and average LOI input modes with dynamic rendering.
- **Mode Toggle**: The mode slider toggles between RID+PID and PID-only modes, dynamically updating the form's required fields and visibility of sections.

### PID Boxes Feature
- **Automatic Display**: When RID file is uploaded, unique PIDs are extracted and displayed in scrollable boxes above the PID drop zone.
- **Label & Wrapper Structure**: PIDs are displayed within a `.pid-boxes-wrapper` containing a descriptive label ("PIDs pulled from the uploaded RID sheet:") and the container with PID boxes.
- **Batch Handling**: Large PID lists (>2000) are split into multiple batches for better performance. First batch always shows descriptive title regardless of total count.
- **Auto-Scroll Behavior**: PID boxes automatically scroll to the right on load (50ms delay) to show the last PIDs, improving SSRS copy workflow.
- **Click-to-Select**: Users can click on PID boxes to select all text for copying to SSRS.
- **Dynamic Visibility**: PID boxes (including label wrapper) are hidden when PID file is uploaded and shown again when PID file is removed.
- **SSRS Integration**: PIDs are formatted with semicolon separation for direct SSRS input.
- **Consistent Styling**: Matches file preview styling with fade animations and dark mode support.

### LOI Mode Functionality
- **Survey-specific Mode**: Shows individual inputs for each survey ID with marketplace links and RID counts.
- **Average LOI Mode**: Shows single input field with aggregated marketplace link to top survey.
- **Dynamic Rendering**: Re-renders LOI inputs when mode changes while preserving validation state.

### Dark Mode
- **Toggle Logic**: Implements a dark mode toggle that updates the UI theme and persists the preference using `localStorage`.
- **Styling**: Adjusts background and text colors for dark and light modes. Hover effects on buttons change to #cd25f1 in dark mode.
- **Complete Coverage**: All UI elements including PID boxes, drag-drop zones, and form elements support dark mode.

### UI Optimizations
- **Compact Design**: Reduced font sizes (body to 0.85rem, h1 to 1.25rem) and spacing for minimal scrolling.
- **Element Sizing**: Adjusted input padding, button sizes, and form group gaps to fit more content.
- **Title Area**: Full-width title area with integrated mode slider and info icon.
- **Responsive Layout**: Two-column layout with vertical divider, maintained across screen sizes.

### Drag-and-Drop File Upload
- **Dual Zones**: Separate drop zones for RID (CSV) and PID (Excel) files with visual feedback.
- **File Validation**: Automatic file type validation with error messages.
- **File Previews**: Dynamic file preview with statistics (file size, record counts, unique counts).
- **Processing Integration**: Seamless integration with file processing pipeline.

### EXE Mode Detection
- Detects if the application is running as a standalone executable (EXE) based on the user agent or protocol. Displays a shutdown button in EXE mode.

### AJAX-Based Communication
- **Form Submission**: Prevents default form submission and uses `fetch` to send form data to the backend. Handles responses for file downloads or error messages.
- **File Download**: Processes server responses to initiate file downloads dynamically.

### Error Handling
- **Frontend Validation**: Validates user inputs and displays error messages dynamically at the bottom of the form.
- **Debugging**: Includes debug logs for key events, such as button clicks and file uploads.
- **Error File Generation**: Backend generates downloadable error files when processing fails.

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
1. Calculate exclusion columns (index, -n/a-, totals).
2. Apply conditional formatting (red color scale).
3. Style -n/a- columns in dark green (`Font(color="006400")`).
4. Set supplier_bu column width to 200px.
5. Use TRANSPOSE formulas for count tables referencing pivot data.

### Formula-Based Excel Calculations
- **Dynamic Columns**: Use Excel formulas instead of Python calculations for better performance and transparency.
- **Column Letters Mapping**: Comprehensive mapping of column names to Excel column letters (A-BT range).
- **Array Formulas**: Use `write_formula()` for dynamic array formulas (TRANSPOSE/XLOOKUP) instead of `write_array_formula()`.
- **Conditional Logic**: Complex nested IF statements for PrioFlag and Tenure_Group calculations.

### Column Width and Formatting Standards
- **Comprehensive Mapping**: Complete column width mapping in `excel_generators.py` covering all 72+ columns.
- **Background Colors**: Threshold columns (#D8D8D8), flag columns (#E5E0EC).
- **Conditional Formatting**: 3-color scales on numeric columns, 2-color scales for flag counts.
- **Format Types**: Text format for time columns, date format for date columns, percentage format for rates.

### Counts Sheet Architecture
- **PrioFlag Counts**: TRANSPOSE formulas referencing PrioFlag Pivot sheet.
- **MultiFlag Counts**: TRANSPOSE formulas referencing MultiFlag Pivot sheet.
- **General Counts**: COUNTIF formulas for various data columns with background colors and data bars.
- **Consistent Formatting**: All tables use borders, background colors matching source pivot sheets, and blue data bars.

### Error Handling Strategy
- Form validation with flash messages positioned at bottom of form.
- Stream processing with try/catch and cleanup.
- Graceful degradation (skip missing columns/sheets).
- Generate error files for download when processing fails.
- Comprehensive error messages with specific guidance.

## Key Conventions

### File Structure
- `requirements.txt`: Minimal dependencies (Flask, pandas, openpyxl, gunicorn).
- `app.yaml`: GAE configuration with F1 instances, auto-scaling 0-2.
- `RIDPIDProcessor.spec`: Simplified PyInstaller config without size optimizations.

### Configuration Patterns
- Thresholds as form inputs with sensible defaults.
- Boolean flags via checkbox presence in `request.form`.
- Environment-specific temp directories (`tempfile.gettempdir()`).
- LOI mode detection via slider state.
- Config dictionaries passed between modules for debugging and user feedback.

### CSS Architecture
- **Color Scheme**: Primary (#2a0b57), Secondary (#cd25f1), Accent (#6502f5).
- **Dark Mode Variables**: Consistent color mapping for light/dark themes.
- **Component Classes**: Modular CSS for drag-drop zones, PID boxes (wrapper, label, container), file previews.
- **PID Boxes Structure**: `.pid-boxes-wrapper` contains `.pid-boxes-label` and `.pid-boxes-container`, all with animation support.
- **Responsive Design**: Flexible layouts that maintain functionality across screen sizes.
- **Code Hygiene**: Empty CSS rulesets are removed to maintain clean, error-free stylesheets.

### JavaScript Architecture
- **Global State Management**: Central variables for ridData, metricsData, surveyLoiValid.
- **Event-Driven Updates**: Dynamic UI updates based on file uploads and mode changes.
- **Validation Pipeline**: Multi-stage validation with real-time feedback.
- **localStorage Integration**: Persistent form state across sessions.

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
Test both processing modes with various threshold combinations and LOI modes. Test PID boxes functionality with large datasets. No automated test suite exists.

## Common Pitfalls
- Column name mismatches (case-sensitive).
- Stream position not reset between operations.
- Missing null checks in boolean mask operations.
- Excel formula references when moving sheets.
- TRANSPOSE formula compatibility - use `write_formula()` not `write_array_formula()`.
- LOI mode state management in frontend validation.
- PID boxes visibility state when switching between files.
- Dark mode CSS inheritance and specificity issues.

## Recent Architecture Changes
- **PID Boxes Enhancement**: Added label wrapper structure (`.pid-boxes-wrapper` → `.pid-boxes-label` + `.pid-boxes-container`) for better UI organization and consistent animations.
- **Auto-Scroll Feature**: PID boxes auto-scroll to right on load (50ms delay) to show last PIDs first, respecting user's manual scroll after initial display.
- **First Batch Title**: When PIDs exceed 2000, first batch maintains descriptive title "PIDs list (for SSRS input, ';' separated):" while subsequent batches show range-based titles.
- **TenureSupplier Pivot Sheet**: New `pivot_sheet51.py` creates supplier × tenure groups analysis with optimized conditional formatting (6 calls instead of inefficient 8+).
- **Conditional Formatting Optimization**: Standardized efficient formatting patterns across all pivot sheets with range-based applications.
- **Enhanced UI**: Title area redesign, improved drag-drop zones, comprehensive dark mode.
- **Formula Architecture**: Complete migration to Excel formulas for dynamic calculations.
- **Column Mapping**: Comprehensive 72-column mapping with proper formatting.
- **Error Handling**: Enhanced error messages and downloadable error files.
- **Configuration System**: Config dictionaries for debugging and user feedback control.
- **CSS Code Quality**: Removed empty rulesets for cleaner, error-free stylesheets.

## When I request any changes, please do not make changes directly to the script. First confirm your understanding, any queries/confusion, provide the plan for approval.