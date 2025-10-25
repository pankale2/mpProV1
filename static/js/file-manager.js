// File Manager Module - Handles all file operations, drag-drop, PID boxes, RID processing
window.FileManager = (function() {
    'use strict';

    // Private variables
    let ridData = null;
    let metricsData = null;
    let userHasScrolled = false; // Track if user has manually scrolled
    let lastContentHeight = 0; // Track content height changes
    
    // DOM elements
    const rawRidsInput = document.getElementById('raw-rids-input');
    const commaRidsOutput = document.getElementById('comma-rids-output');
    const ridProcessingContainer = document.getElementById('rid-processing-container');
    const rawRidsCount = document.getElementById('raw-rids-count');
    const commaRidsCount = document.getElementById('comma-rids-count');

    // Private functions
    function scrollToBottom(element) {
        if (element) {
            element.scrollTop = element.scrollHeight;
        }
    }

    function isScrolledToBottom(element) {
        if (!element) return true;
        return element.scrollTop >= (element.scrollHeight - element.clientHeight - 5); // 5px tolerance
    }

    function setupScrollTracking(element) {
        if (!element) return;
        
        element.addEventListener('scroll', function() {
            // Only mark as user-scrolled if they scroll away from the bottom
            if (!isScrolledToBottom(element)) {
                userHasScrolled = true;
            } else {
                // If they scroll back to bottom, reset the flag
                userHasScrolled = false;
            }
        });
    }

    function updateLineCounts(rawLines = 0, outputLines = 0) {
        if (rawRidsCount) {
            rawRidsCount.textContent = `Lines: ${rawLines}/6000`;
            rawRidsCount.classList.toggle('over-limit', rawLines > 6000);
        }
        
        if (commaRidsCount) {
            commaRidsCount.textContent = `Lines: ${outputLines}`;
        }
    }

    function processRawRids() {
        if (!rawRidsInput || !commaRidsOutput) return;
        
        const rawText = rawRidsInput.value;
        if (!rawText.trim()) {
            commaRidsOutput.value = '';
            updateLineCounts(0, 0);
            // Reset scroll tracking for empty content
            userHasScrolled = false;
            lastContentHeight = 0;
            return;
        }
        
        // Split by newlines and filter out empty lines
        const lines = rawText.split('\n')
            .map(line => line.trim())
            .filter(line => line.length > 0);
        
        // Update line count
        const totalRawLines = rawText.split('\n').length;
        updateLineCounts(totalRawLines, lines.length);
        
        // Check line limit
        if (lines.length > 6000) {
            window.App.addJsMessage('Too many lines! Maximum 6000 lines accepted. Please reduce the input.', 'error');
            return;
        }
        
        // Each RID on separate line with comma (except last)
        let commaText = '';
        for (let i = 0; i < lines.length; i++) {
            if (i === lines.length - 1) {
                // Last line: no comma
                commaText += lines[i];
            } else {
                // All other lines: add comma
                commaText += lines[i] + ',\n';
            }
        }
        
        // Check if content has actually changed (new content added)
        const previousContent = commaRidsOutput.value;
        const newContentHeight = commaText.split('\n').length;
        const hasNewContent = newContentHeight > lastContentHeight;
        
        commaRidsOutput.value = commaText;
        
        // Auto-scroll logic: only scroll if new content was added AND user hasn't manually scrolled
        if (hasNewContent && !userHasScrolled) {
            scrollToBottom(commaRidsOutput);
        }
        
        // Update tracking variables
        lastContentHeight = newContentHeight;
        
        // Clear any previous error messages
        window.App.clearJsMessages();
    }

    function selectRidText() {
        if (!commaRidsOutput) return;
        
        commaRidsOutput.select();
        commaRidsOutput.setSelectionRange(0, 99999); // For mobile devices
        
        // Try to copy to clipboard
        try {
            document.execCommand('copy');
        } catch (err) {
            console.log('Copy to clipboard failed:', err);
        }
    }

    function showRidProcessing() {
        if (ridProcessingContainer) {
            ridProcessingContainer.style.display = 'block';
            ridProcessingContainer.classList.remove('fade-slide-out');
            ridProcessingContainer.classList.add('fade-slide-in');
        }
    }

    function hideRidProcessing() {
        if (ridProcessingContainer) {
            ridProcessingContainer.classList.remove('fade-slide-in');
            ridProcessingContainer.classList.add('fade-slide-out');
            setTimeout(() => {
                if (ridProcessingContainer.classList.contains('fade-slide-out')) {
                    ridProcessingContainer.style.display = 'none';
                }
            }, 300);
        }
    }

    // Drag and drop functions
    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    function setupDropZone(dropZone, fileInput, expectedType, processCallback) {
        // Prevent default drag behaviors
        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
            document.body.addEventListener(eventName, preventDefaults, false);
        });

        // Highlight drop zone when item is dragged over it
        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.add('drag-over'), false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.classList.remove('drag-over'), false);
        });

        // Handle dropped files
        dropZone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            handleFiles(files, fileInput, expectedType, processCallback);
        }, false);

        // Handle click to browse
        dropZone.addEventListener('click', () => {
            fileInput.click();
        });

        // Handle file input change
        fileInput.addEventListener('change', (e) => {
            const files = e.target.files;
            handleFiles(files, fileInput, expectedType, processCallback);
        });
    }

    function handleFiles(files, fileInput, expectedType, processCallback) {
        if (files.length === 0) return;
        
        const file = files[0];
        const fileExtension = file.name.split('.').pop().toLowerCase();
        
        // Validate file type
        if (fileExtension !== expectedType) {
            const typeText = expectedType === 'csv' ? 'CSV' : 'Excel (.xlsx)';
            window.App.addJsMessage(`Invalid file type. Please select a ${typeText} file.`, 'error');
            return;
        }
        
        // Update the hidden file input
        const dt = new DataTransfer();
        dt.items.add(file);
        fileInput.files = dt.files;
        
        // Process the file
        processCallback(file, fileInput);
    }

    // PID handling functions
    function findPidColumn(data) {
        if (!data || data.length === 0) return null;
        const keys = Object.keys(data[0]);
        return keys.find(k => k.trim().toLowerCase() === 'pid');
    }

    function extractUniquePids(data, pidKey) {
        const uniquePids = [...new Set(
            data.map(row => (row[pidKey] || '').toString().trim())
                .filter(pid => pid && pid !== '' && pid !== 'nan' && pid !== 'null')
        )];
        return uniquePids.sort(); // Sort for consistent display
    }

    function createPidContainer() {
        // Create wrapper to hold both label and container
        const wrapper = document.createElement('div');
        wrapper.className = 'pid-boxes-wrapper';
        
        // Create label
        const label = document.createElement('label');
        label.className = 'pid-boxes-label';
        label.textContent = 'PIDs pulled from uploaded RID sheet, ready for SSRS input, \';\' separated:';
        
        // Create container
        const container = document.createElement('div');
        container.className = 'pid-boxes-container';
        
        // Append label and container to wrapper
        wrapper.appendChild(label);
        wrapper.appendChild(container);
        
        return { wrapper, container };
    }

    function createPidBoxElement(title, pids) {
        const pidBox = document.createElement('div');
        pidBox.className = 'pid-box';
        
        const pidsString = pids.join(';');
        
        pidBox.innerHTML = `
            <div class="pid-box-content">
                <span class="pid-box-title">${title}</span>
                <div class="pid-box-scroll" onclick="selectAllText(this)" title="Click to select all">${pidsString}</div>
            </div>
        `;
        
        return pidBox;
    }

    function displayPidBoxes(allPids) {
        removePidBoxes();
        
        const batchSize = 2000;
        const pidDropZone = document.getElementById('pid-drop-zone');
        if (!pidDropZone) return;
        
        const { wrapper, container } = createPidContainer();
        
        if (allPids.length <= batchSize) {
            const pidBox = createPidBoxElement(`PIDs (${allPids.length}):`, allPids);
            container.appendChild(pidBox);
        } else {
            const totalBatches = Math.ceil(allPids.length / batchSize);
            
            for (let i = 0; i < totalBatches; i++) {
                const start = i * batchSize + 1;
                const end = Math.min((i + 1) * batchSize, allPids.length);
                const batchPids = allPids.slice(i * batchSize, end);
                
                // Use descriptive title for first batch, range title for others
                const title = (i === 0) 
                    ? `PIDs ${start}-${end}:`
                    : `PIDs ${start}-${end}:`;
                
                const pidBox = createPidBoxElement(title, batchPids);
                container.appendChild(pidBox);
            }
        }
        
        pidDropZone.parentNode.insertBefore(wrapper, pidDropZone);
        
        // Animate in wrapper (which contains both label and boxes)
        wrapper.classList.add('fade-slide-in');
        
        // Auto-scroll all PID boxes to the right after DOM rendering
        setTimeout(() => {
            const pidScrollBoxes = container.querySelectorAll('.pid-box-scroll');
            pidScrollBoxes.forEach(box => {
                box.scrollLeft = box.scrollWidth;
            });
        }, 50);
    }

    function removePidBoxes() {
        const existingWrapper = document.querySelector('.pid-boxes-wrapper');
        if (existingWrapper) {
            existingWrapper.classList.remove('fade-slide-in');
            existingWrapper.classList.add('fade-slide-out');
            setTimeout(() => {
                if (existingWrapper.parentNode) {
                    existingWrapper.remove();
                }
            }, 300);
        }
    }

    function hidePidBoxes() {
        const existingWrapper = document.querySelector('.pid-boxes-wrapper');
        if (existingWrapper) {
            existingWrapper.classList.remove('fade-slide-in');
            existingWrapper.classList.add('fade-slide-out');
            setTimeout(() => {
                if (existingWrapper.classList.contains('fade-slide-out')) {
                    existingWrapper.style.display = 'none';
                }
            }, 300);
        }
    }

    function showPidBoxes() {
        const existingWrapper = document.querySelector('.pid-boxes-wrapper');
        if (existingWrapper) {
            existingWrapper.style.display = 'block';
            existingWrapper.classList.remove('fade-slide-out');
            existingWrapper.classList.add('fade-slide-in');
        }
    }

    function extractAndDisplayPids(data) {
        const pidKey = findPidColumn(data);
        if (!pidKey) return; // No PIDs found, don't display anything
        
        const uniquePids = extractUniquePids(data, pidKey);
        if (uniquePids.length === 0) return; // No PIDs found, don't display anything
        
        displayPidBoxes(uniquePids);
    }

    // File processing functions
    function processRidFile(file, fileInput) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const text = e.target.result;
                
                // For CSV files, use XLSX with proper CSV parsing
                const workbook = XLSX.read(text, {type: 'string', raw: false});
                
                // Check if workbook was created successfully
                if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
                    throw new Error('Invalid CSV structure - no data sheets found');
                }
                
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                
                // Check if sheet exists and has data
                if (!sheet) {
                    throw new Error('Invalid CSV structure - no data found');
                }
                
                const json = XLSX.utils.sheet_to_json(sheet, {defval: '', raw: false});
                
                // Validate that we have data
                if (!json || json.length === 0) {
                    throw new Error('CSV file appears to be empty or has no valid data rows');
                }

                // Find surveyid column (case-insensitive)
                let surveyidKey = null;
                if (json.length > 0) {
                    const keys = Object.keys(json[0]);
                    surveyidKey = keys.find(k => k.trim().toLowerCase() === 'surveyid');
                }

                let surveyIds = {};
                if (surveyidKey) {
                    json.forEach(row => {
                        const sid = (row[surveyidKey] || '').toString().trim();
                        if (sid && sid !== '' && sid !== 'undefined' && sid !== 'null') {
                            surveyIds[sid] = (surveyIds[sid] || 0) + 1;
                        }
                    });
                }

                // Extract unique PID count using existing functions
                const pidKey = findPidColumn(json);
                const uniquePids = pidKey ? extractUniquePids(json, pidKey) : [];
                const pidCount = uniquePids.length;

                ridData = json;

                // Show file preview with PID count
                showFilePreview('rid', file, json.length, Object.keys(surveyIds).length, pidCount);

                // Extract and display PIDs
                extractAndDisplayPids(json);

                // Check if PID file is already uploaded when RID file is processed
                // If PID file exists, hide PID boxes immediately
                if (metricsData) {
                    hidePidBoxes();
                }

                // Hide RID processing area when file is uploaded
                hideRidProcessing();

                // Always render LOI inputs (with or without survey data)
                window.FormManager.renderLoiInputs(surveyIds);

                // Always show LOI group
                window.FormManager.showLoiGroup();
                window.FormManager.updateProcessBtnState();
                window.App.clearJsMessages();
            } catch (error) {
                console.error('Error parsing RID file:', error);
                let errorMessage = 'Error processing RID file. ';
                
                // Provide more specific error messages
                if (error.message.includes('Invalid CSV structure')) {
                    errorMessage += error.message;
                } else if (error.message.includes('empty')) {
                    errorMessage += 'The file appears to be empty or contains no valid data.';
                } else {
                    errorMessage += 'Please ensure it\'s a valid CSV format with proper headers (pid, surveyid, supplier_bu, etc.).';
                }
                
                window.App.addJsMessage(errorMessage, 'error');
                removeFile('rid');
            }
        };
        reader.readAsText(file);
    }

    function processPidFile(file, fileInput) {
        try {
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    // For PID file, read and parse the Excel file to count PIDs
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, {type: 'array'});
                    
                    console.log('Available sheets:', workbook.SheetNames); // Debug log
                    
                    // Look for the "Marketplace Metrics by PID" sheet first
                    let sheetName = null;
                    
                    // Try exact match first
                    if (workbook.SheetNames.includes('Marketplace Metrics by PID')) {
                        sheetName = 'Marketplace Metrics by PID';
                    } else {
                        // Try case-insensitive search for sheets containing key terms
                        sheetName = workbook.SheetNames.find(name => {
                            const lowerName = name.toLowerCase();
                            return lowerName.includes('marketplace') || 
                                   lowerName.includes('metrics') || 
                                   lowerName.includes('pid');
                        });
                        
                        // If still not found, use the first sheet
                        if (!sheetName && workbook.SheetNames.length > 0) {
                            sheetName = workbook.SheetNames[0];
                            console.log('Using first available sheet:', sheetName); // Debug log
                        }
                    }
                    
                    if (!sheetName) {
                        throw new Error('No valid sheets found in the Excel file');
                    }
                    
                    console.log('Using sheet:', sheetName); // Debug log
                    
                    const sheet = workbook.Sheets[sheetName];
                    if (!sheet) {
                        throw new Error(`Could not access sheet "${sheetName}" in the Excel file`);
                    }
                    
                    // Try to convert sheet to JSON with different approaches
                    let json = null;
                    let totalRecords = 0;
                    
                    try {
                        // First try with skipping 5 rows (standard format)
                        json = XLSX.utils.sheet_to_json(sheet, {
                            defval: '', 
                            raw: false,
                            range: 5
                        });
                        
                        if (!json || json.length === 0) {
                            // If no data with skip, try without skipping
                            json = XLSX.utils.sheet_to_json(sheet, {
                                defval: '', 
                                raw: false
                            });
                        }
                    } catch (parseError) {
                        console.error('Error parsing sheet data:', parseError);
                        // Fallback: try basic parsing
                        json = XLSX.utils.sheet_to_json(sheet, {defval: ''});
                    }
                    
                    if (!json || json.length === 0) {
                        throw new Error('Excel file contains no readable data. Please check the file format.');
                    }
                    
                    console.log('Parsed records:', json.length); // Debug log
                    console.log('Sample columns:', Object.keys(json[0] || {})); // Debug log
                    
                    totalRecords = json.length;
                    
                    // Find PID column (case-insensitive, try multiple variations)
                    const possiblePidColumns = ['pid', 'PID', 'Pid', 'p_id', 'P_ID'];
                    let pidKey = null;
                    
                    if (json.length > 0) {
                        const availableKeys = Object.keys(json[0]);
                        
                        // Try exact matches first
                        pidKey = possiblePidColumns.find(col => availableKeys.includes(col));
                        
                        // If not found, try case-insensitive search
                        if (!pidKey) {
                            pidKey = availableKeys.find(key => 
                                possiblePidColumns.some(pid => 
                                    key.toString().toLowerCase().trim() === pid.toLowerCase()
                                )
                            );
                        }
                    }
                    
                    let pidCount = 0;
                    if (pidKey && json.length > 0) {
                        console.log('Using PID column:', pidKey); // Debug log
                        
                        // Count unique PIDs
                        const uniquePids = [...new Set(
                            json.map(row => {
                                const pidValue = row[pidKey];
                                return pidValue ? pidValue.toString().trim() : '';
                            })
                            .filter(pid => pid && pid !== '' && pid !== 'nan' && pid !== 'null' && pid !== 'undefined')
                        )];
                        pidCount = uniquePids.length;
                        console.log('Unique PIDs found:', pidCount); // Debug log
                    } else {
                        console.log('No PID column found or no data'); // Debug log
                    }
                    
                    metricsData = true;
                    
                    // Show file preview with actual counts
                    showFilePreview('pid', file, totalRecords, pidCount);
                    
                    // Hide PID boxes when PID file is uploaded
                    hidePidBoxes();
                    
                    window.FormManager.updateProcessBtnState();
                    window.App.clearJsMessages();
                } catch (error) {
                    console.error('Error processing PID file:', error);
                    let errorMessage = 'Error processing PID file. ';
                    
                    if (error.message.includes('No valid sheets found')) {
                        errorMessage += 'The Excel file appears to be empty or corrupted.';
                    } else if (error.message.includes('no readable data')) {
                        errorMessage += 'The file contains no readable data. Please ensure it\'s a valid Excel file with data.';
                    } else if (error.message.includes('Could not access sheet')) {
                        errorMessage += 'Could not read the sheet data. Please try re-exporting the file from SSRS.';
                    } else {
                        errorMessage += 'Please ensure it\'s a valid Excel (.xlsx) file exported from SSRS with the correct data structure.';
                    }
                    
                    window.App.addJsMessage(errorMessage, 'error');
                    removeFile('pid');
                }
            };
            reader.readAsArrayBuffer(file);
        } catch (error) {
            console.error('Error reading PID file:', error);
            window.App.addJsMessage('Error reading PID file. Please ensure it\'s a valid Excel (.xlsx) format.', 'error');
            removeFile('pid');
        }
    }

    function showFilePreview(type, file, totalRecords, uniqueCount, pidCount = 0) {
        const previewId = type === 'rid' ? 'rid-file-preview' : 'pid-file-preview';
        const dropZoneId = type === 'rid' ? 'rid-drop-zone' : 'pid-drop-zone';
        
        const preview = document.getElementById(previewId);
        const dropZone = document.getElementById(dropZoneId);
        
        if (preview && dropZone) {
            // Animate out drop zone
            dropZone.classList.remove('fade-slide-in');
            dropZone.classList.add('fade-slide-out');
            
            setTimeout(() => {
                dropZone.style.display = 'none';
                preview.style.display = 'block';
                
                // Animate in preview
                preview.classList.remove('fade-slide-out');
                preview.classList.add('fade-slide-in');
            }, 300);
            
            // Update preview content
            const fileName = preview.querySelector('.file-name');
            const fileStats = preview.querySelector('.file-stats');
            
            if (fileName) {
                fileName.textContent = file.name;
            }
            
            if (fileStats) {
                const sizeText = formatFileSize(file.size);
                const validTotalRecords = Math.max(0, totalRecords || 0);
                const validUniqueCount = Math.max(0, uniqueCount || 0);
                const validPidCount = Math.max(0, pidCount || 0);
                
                if (type === 'rid') {
                    fileStats.innerHTML = `${sizeText} • ${validTotalRecords.toLocaleString()} RIDs • ${validPidCount.toLocaleString()} Unique PIDs • ${validUniqueCount.toLocaleString()} Survey IDs`;
                } else {
                    fileStats.innerHTML = `${sizeText} • ${validTotalRecords.toLocaleString()} PIDs`;
                }
            }
        }
    }

    function formatFileSize(bytes) {
        if (bytes === 0) return '0 Bytes';
        const k = 1024;
        const sizes = ['Bytes', 'KB', 'MB', 'GB'];
        const i = Math.floor(Math.log(bytes) / Math.log(k));
        return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
    }

    function removeFile(type) {
        const previewId = type === 'rid' ? 'rid-file-preview' : 'pid-file-preview';
        const dropZoneId = type === 'rid' ? 'rid-drop-zone' : 'pid-drop-zone';
        const fileInputId = type === 'rid' ? 'rid_file' : 'metrics_file';
        
        const preview = document.getElementById(previewId);
        const dropZone = document.getElementById(dropZoneId);
        const fileInput = document.getElementById(fileInputId);
        
        if (preview && dropZone && fileInput) {
            fileInput.value = '';
            
            // Animate out preview
            preview.classList.remove('fade-slide-in');
            preview.classList.add('fade-slide-out');
            
            setTimeout(() => {
                preview.style.display = 'none';
                dropZone.style.display = 'flex';
                
                // Animate in drop zone
                dropZone.classList.remove('fade-slide-out');
                dropZone.classList.add('fade-slide-in');
            }, 300);
            
            if (type === 'rid') {
                ridData = null;
                window.FormManager.resetRidData();
                removePidBoxes();
                showRidProcessing();
                window.FormManager.showLoiGroup();
            } else {
                metricsData = null;
                showPidBoxes();
            }
            
            window.FormManager.updateProcessBtnState();
        }
    }

    function initializeDragAndDrop() {
        const ridDropZone = document.getElementById('rid-drop-zone');
        const pidDropZone = document.getElementById('pid-drop-zone');
        const ridFileInput = document.getElementById('rid_file');
        const metricsFileInput = document.getElementById('metrics_file');

        // RID Drop Zone
        if (ridDropZone) {
            setupDropZone(ridDropZone, ridFileInput, 'csv', processRidFile);
        }

        // PID Drop Zone  
        if (pidDropZone) {
            setupDropZone(pidDropZone, metricsFileInput, 'xlsx', processPidFile);
        }
    }

    function initializeRidProcessing() {
        const rawInput = document.getElementById('raw-rids-input');
        const commaOutput = document.getElementById('comma-rids-output');
        const rawCountSpan = document.getElementById('raw-rids-count');
        const commaCountSpan = document.getElementById('comma-rids-count');
        
        if (!rawInput || !commaOutput) return;

        // Ensure initial state is clean (fixes first-launch stray value issue)
        rawInput.value = '';
        commaOutput.value = '';
        rawInput.placeholder = rawInput.placeholder || 'Paste RIDs here, one per line, max 6000 lines accepted...';
        commaOutput.placeholder = commaOutput.placeholder || 'Auto-filled with RIDs and comma...';
        userHasScrolled = false;
        lastContentHeight = 0;
        updateLineCounts(0, 0);
        
        // Process RIDs on input
        rawInput.addEventListener('input', function() {
            const rawText = this.value;
            
            // Split by lines, filter empty, trim, and convert to lowercase
            const rids = rawText.split(/\r?\n/)
                .filter(line => line.trim() !== '')
                .map(line => line.trim().toLowerCase());
            
            // Join with commas and newlines for output (one RID per line with comma)
            let commaText = '';
            for (let i = 0; i < rids.length; i++) {
                if (i === rids.length - 1) {
                    // Last line: no comma
                    commaText += rids[i];
                } else {
                    // All other lines: add comma
                    commaText += rids[i] + ',\n';
                }
            }
            
            // Update output
            commaOutput.value = commaText;
            
            // Update counters
            const lineCount = rids.length;
            rawCountSpan.textContent = `Lines: ${lineCount}/6000`;
            commaCountSpan.textContent = `Lines: ${lineCount}`;
            
            // Highlight if over limit
            if (lineCount > 6000) {
                rawCountSpan.classList.add('over-limit');
            } else {
                rawCountSpan.classList.remove('over-limit');
            }
        });
        
        // Click-to-select functionality for comma output
        commaOutput.addEventListener('click', function() {
            this.select();
        });
    }

    // Public API
    return {
        // Data access
        getRidData: () => ridData,
        getMetricsData: () => metricsData,
        setRidData: (data) => { ridData = data; },
        setMetricsData: (data) => { metricsData = data; },
        
        // File operations
        processRidFile,
        processPidFile,
        removeFile,
        showFilePreview,
        formatFileSize,
        
        // RID processing
        processRawRids,
        selectRidText,
        showRidProcessing,
        hideRidProcessing,
        updateLineCounts,
        
        // PID boxes
        displayPidBoxes,
        removePidBoxes,
        hidePidBoxes,
        showPidBoxes,
        extractAndDisplayPids,
        
        // Initialization
        initializeDragAndDrop,
        initializeRidProcessing
    };
})();

// Global function for compatibility
function selectAllText(element) {
    const range = document.createRange();
    range.selectNodeContents(element);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
}

// Global function for removeFile compatibility
function removeFile(type) {
    window.FileManager.removeFile(type);
}
