// Main App Module - Coordinates all functionality and handles initialization
window.App = (function() {
    'use strict';

    // Message handling
    let jsMsgDiv = document.getElementById('js-msg-div');
    if (!jsMsgDiv) {
        jsMsgDiv = document.createElement('div');
        jsMsgDiv.id = 'js-msg-div';
        jsMsgDiv.style.marginTop = '15px';
        // Insert before spinner overlay instead of before form
        const container = document.querySelector('.container');
        const spinnerOverlay = document.getElementById('processing-overlay');
        container.insertBefore(jsMsgDiv, spinnerOverlay);
    }

    // Accumulate messages in an array
    let jsMessages = [];

    function showJsMessages() {
        if (jsMessages.length === 0) {
            jsMsgDiv.innerHTML = '';
            return;
        }
        jsMsgDiv.innerHTML = `<div class="flash-messages">${jsMessages.map(m => `<li class="${m.type}">${m.text}</li>`).join('')}</div>`;
    }

    function addJsMessage(msg, type='error') {
        jsMessages.push({text: msg, type: type});
        showJsMessages();
    }

    function clearJsMessages() {
        jsMessages = [];
        showJsMessages();
    }

    // EXE mode detection
    function checkEXEMode() {
        // Check if running as EXE (window title or other indicators)
        const isEXE = window.navigator.userAgent.includes('Electron') || 
                     window.location.protocol === 'file:' ||
                     window.location.hostname === '127.0.0.1';
        
        if (isEXE) {
            document.getElementById('shutdown-btn').style.display = 'block';
        }
    }

    // Form submission handling
    function setupFormSubmission() {
        document.getElementById('main-form').addEventListener('submit', function(e) {
            e.preventDefault();
            document.getElementById('processing-overlay').style.display = 'flex';
            const formData = new FormData(this);

            fetch('/', {
                method: 'POST',
                body: formData
            })
            .then(response => {
                if (response.ok) {
                    // SUCCESS: Status is 200-299. This should be a file download.
                    const disposition = response.headers.get('Content-Disposition');
                    if (disposition && disposition.includes('attachment')) {
                        return response.blob().then(blob => {
                            const filenameMatch = disposition.match(/filename="?([^"]+)"?/);
                            const filename = filenameMatch ? filenameMatch[1] : 'report.xlsx';
                            const url = window.URL.createObjectURL(blob);
                            const a = document.createElement('a');
                            a.style.display = 'none';
                            a.href = url;
                            a.download = filename;
                            document.body.appendChild(a);
                            a.click();
                            window.URL.revokeObjectURL(url);
                            a.remove();
                            
                            document.getElementById('processing-overlay').style.display = 'none';
                            // Updated to place success message at bottom
                            let flashContainer = document.querySelector('.flash-messages');
                            if (!flashContainer) {
                                flashContainer = document.createElement('ul');
                                flashContainer.className = 'flash-messages';
                                flashContainer.style.marginTop = '15px';
                                // Insert before spinner overlay
                                const container = document.querySelector('.container');
                                const spinnerOverlay = document.getElementById('processing-overlay');
                                container.insertBefore(flashContainer, spinnerOverlay);
                            }
                            flashContainer.innerHTML = `<li class="success">Processing complete! Your download has started.</li>`;
                        });
                    } else {
                        // Unexpected 200 OK without a file. Treat as an error and reload.
                        alert('An unexpected response was received from the server. The page will now reload.');
                        window.location.reload();
                    }
                } else {
                    // ERROR: Status is not OK (e.g., 400 or 500).
                    // The backend has flashed the error message to the session.
                    // Reload the page to display it.
                    window.location.reload();
                }
            })
            .catch((error) => {
                console.error('Network or script error:', error);
                alert('A critical network error occurred. Please check your connection and try again.');
                document.getElementById('processing-overlay').style.display = 'none';
            });
        });
    }

    // File input event handlers setup
    function setupFileInputHandlers() {
        const ridFileInput = document.getElementById('rid_file');
        const metricsFileInput = document.getElementById('metrics_file');

        // Update how ridData is processed
        ridFileInput.addEventListener('change', async function(e) {
            if (this.files.length > 0) {
                try {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        const text = e.target.result;
                        const workbook = XLSX.read(text, {type: 'string'});
                        const sheetName = workbook.SheetNames[0];
                        const sheet = workbook.Sheets[sheetName];
                        const json = XLSX.utils.sheet_to_json(sheet, {defval: '', raw: false});

                        // Find surveyid column (case-insensitive)
                        let surveyidKey = null;
                        if (json.length > 0) {
                            const keys = Object.keys(json[0]);
                            surveyidKey = keys.find(k => k.trim().toLowerCase() === 'surveyid');
                        }

                        let surveyIds = {};
                        if (surveyidKey) {
                            json.forEach(row => {
                                const sid = (row[surveyidKey] || '').trim();
                                if (sid && sid !== '') {
                                    surveyIds[sid] = (surveyIds[sid] || 0) + 1;
                                }
                            });
                        }

                        window.FileManager.setRidData(json);

                        // Create survey LOI inputs if surveys found
                        if (Object.keys(surveyIds).length > 0) {
                            window.FormManager.renderLoiInputs(surveyIds);
                        }

                        window.FormManager.showLoiGroup();
                        window.FormManager.updateProcessBtnState();
                    };
                    reader.readAsText(this.files[0]);
                } catch (error) {
                    console.error('Error parsing RID file:', error);
                    window.FileManager.setRidData(null);
                    window.FormManager.resetRidData();
                    window.FormManager.showLoiGroup();
                    window.FormManager.updateProcessBtnState();
                }
            } else {
                window.FileManager.setRidData(null);
                window.FormManager.resetRidData();
                window.FormManager.showLoiGroup();
                window.FormManager.updateProcessBtnState();
            }
        });

        // Add metrics file handler
        metricsFileInput.addEventListener('change', function(e) {
            if (this.files.length > 0) {
                try {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        window.FileManager.setMetricsData(true); // Just need to know it's uploaded
                        window.FormManager.updateProcessBtnState();
                    };
                    reader.readAsArrayBuffer(this.files[0]); // Use ArrayBuffer for Excel files
                } catch (error) {
                    console.error('Error reading Metrics file:', error);
                    window.FileManager.setMetricsData(null);
                    window.FormManager.updateProcessBtnState();
                }
            } else {
                window.FileManager.setMetricsData(null);
                window.FormManager.updateProcessBtnState();
            }
        });
    }

    // Main initialization function
    function initialize() {
        console.log('DOM loaded, initializing...'); // Debug log
        
        // Initialize mode
        const modeSlider = document.getElementById('mode_slider');
        modeSlider.checked = false;  // Default to RID+PID mode
        
        // Set LOI mode slider to checked (Average LOI mode by default)
        const loiModeSlider = document.getElementById('loi_mode_slider');
        if (loiModeSlider) {
            loiModeSlider.checked = true;
        }
        
        // Always show survey LOI container and initialize with default Average LOI input
        const surveyLoiContainer = document.getElementById('survey-loi-container');
        surveyLoiContainer.style.display = 'block';
        window.FormManager.renderLoiInputs({}); // Initialize with empty survey data
        
        // Restore form state
        window.FormManager.restoreFormState();
        
        // Restore dark mode preference
        const darkPref = localStorage.getItem('dark_mode');
        window.FormManager.setDarkMode(darkPref === '1');
        
        // Check EXE mode
        checkEXEMode();
        
        // Setup all event listeners
        window.FormManager.setupEventListeners();
        setupFormSubmission();
        setupFileInputHandlers();
        
        // Initialize modules
        window.FileManager.initializeDragAndDrop();
        window.FileManager.initializeRidProcessing();
        
        // Update process button state
        window.FormManager.updateProcessBtnState();
        
        console.log('Initialization complete'); // Debug log
    }

    // Public API
    return {
        // Message handling
        addJsMessage,
        clearJsMessages,
        showJsMessages,
        
        // Initialization
        initialize
    };
})(); // Fixed: Added proper closing for IIFE

// Global functions for compatibility
function shutdownApp() {
    if (confirm('Are you sure you want to exit the application?')) {
        fetch('/shutdown', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            }
        }).then(() => {
            alert('Application is shutting down. You can close this browser window.');
            window.close();
        }).catch(() => {
            alert('Application is shutting down. You can close this browser window.');
            window.close();
        });
    }
}

// Initialize everything when DOM is loaded
document.addEventListener('DOMContentLoaded', window.App.initialize);