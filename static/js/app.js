// Global variables
let allSurveyIds = [];
let surveyLoiValid = {};

const ridFileInput = document.getElementById('rid_file');
const metricsFileInput = document.getElementById('metrics_file');
const surveyLoiContainer = document.getElementById('survey-loi-container');
const surveyLoiInputsDiv = document.getElementById('survey-loi-inputs');
const processBtn = document.getElementById('process-btn');
const modeSlider = document.getElementById('mode_slider');
const darkModeToggle = document.getElementById('dark-mode-toggle');
const infoIcon = document.getElementById('info-icon');
const infoPopup = document.getElementById('info-popup');
const infoPopupClose = document.getElementById('info-popup-close');
const loiModeSlider = document.getElementById('loi_mode_slider');
const marketplaceLinkContainer = document.getElementById('marketplace-link-container');
const marketplaceHyperlink = document.getElementById('marketplace-hyperlink');
const marketplaceLabel = document.getElementById('marketplace-label');

function isPidOnlyMode() { return modeSlider.checked; }

function isAverageLoiMode() {
  return loiModeSlider && loiModeSlider.checked;
}

// Message area for JS alerts - updated to place at bottom
let jsMsgDiv = document.getElementById('js-msg-div');
if (!jsMsgDiv) {
  jsMsgDiv = document.createElement('div');
  jsMsgDiv.id = 'js-msg-div';
  jsMsgDiv.style.marginTop = '15px'; // Changed from marginBottom
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

// Store parsed data for cross-file checks
let ridData = null;
let metricsData = null;

function showLoiGroup() {
  // Always show the container - no conditional logic
  surveyLoiContainer.style.display = 'block';
}

function validateSurveyLoi(surveyId, value) {
    const numValue = parseFloat(value);
    const isValid = !isNaN(numValue) && numValue >= 3 && numValue <= 100;
    surveyLoiValid[surveyId] = isValid;
    return isValid;
}

// Helper to get surveyid with highest RID count
function getTopSurveyId(surveyIdsObj) {
  let maxCount = -1;
  let topSurveyId = null;
  for (const [sid, count] of Object.entries(surveyIdsObj)) {
    if (count > maxCount) {
      maxCount = count;
      topSurveyId = sid;
    }
  }
  return topSurveyId;
}

// Render LOI input(s) based on mode
function renderLoiInputs(surveyData) {
  surveyLoiInputsDiv.innerHTML = '';
  surveyLoiValid = {};

  if (isAverageLoiMode()) {
    // Average LOI mode: show only one input
    const inputGroup = document.createElement('div');
    inputGroup.className = 'form-group';

    const label = document.createElement('label');
    label.innerHTML = `Survey LOI (to compare with the session_loi of each RID): `;

    const input = document.createElement('input');
    input.type = 'number';
    input.name = 'average_loi';
    input.id = 'average_loi';
    input.min = '3';
    input.max = '100';
    input.step = '0.1';
    input.required = false; // Not required initially
    input.placeholder = 'Enter LOI (3-100)';

    input.addEventListener('input', function() {
      const numValue = parseFloat(this.value);
      const isValid = !isNaN(numValue) && numValue >= 3 && numValue <= 100;
      input.style.borderColor = isValid ? '' : '#e53935';
      input.style.background = isValid ? '' : '#fff5f5';
      surveyLoiValid['average_loi'] = isValid;
      updateProcessBtnState();
    });

    inputGroup.appendChild(label);
    inputGroup.appendChild(input);
    surveyLoiInputsDiv.appendChild(inputGroup);

    // Always set allSurveyIds for Average LOI mode
    if (surveyData && Object.keys(surveyData).length > 0) {
      allSurveyIds = Object.keys(surveyData);
      // Show "Marketplace" as hyperlink when survey data is available
      marketplaceLabel.style.display = 'none';
      marketplaceLinkContainer.style.display = 'inline';
      // Set hyperlink to top surveyid
      const topSurveyId = getTopSurveyId(surveyData);
      if (topSurveyId) {
        marketplaceHyperlink.href = `https://www.samplicio.us/fulcrum/next/surveys/${topSurveyId}/reports`;
      } else {
        marketplaceHyperlink.href = '#';
      }
    } else {
      allSurveyIds = [];
      // Show static "Get from Marketplace" text when no survey data
      marketplaceLabel.style.display = 'inline';
      marketplaceLinkContainer.style.display = 'none';
    }
  } else {
    // Survey-wise LOI mode: show all inputs
    marketplaceLabel.style.display = 'inline';
    marketplaceLinkContainer.style.display = 'none';

    if (surveyData && Object.keys(surveyData).length > 0) {
      allSurveyIds = Object.keys(surveyData);
      
      // Sort surveys by occurrence count (descending)
      const sortedSurveys = Object.entries(surveyData)
        .sort(([,a], [,b]) => b - a);

      sortedSurveys.forEach(([surveyId, count]) => {
        const inputGroup = document.createElement('div');
        inputGroup.className = 'form-group';

        const label = document.createElement('label');
        label.innerHTML = `Survey ID: <a href="https://www.samplicio.us/fulcrum/next/surveys/${surveyId}/reports" target="_blank" style="color: #cd25f1; font-weight: bold;">${surveyId}</a> <span style="color: #666;">(${count} RIDs)</span>:`;

        const input = document.createElement('input');
        input.type = 'number';
        input.name = `survey_loi_${surveyId}`;
        input.id = `survey_loi_${surveyId}`;
        input.min = '3';
        input.max = '100';
        input.step = '0.1';
        input.required = false; // Will be validated via updateProcessBtnState
        input.placeholder = 'Enter LOI (3-100)';

        input.addEventListener('input', function() {
          const isValid = validateSurveyLoi(surveyId, this.value);
          this.style.borderColor = isValid ? '' : '#e53935';
          this.style.background = isValid ? '' : '#fff5f5';
          updateProcessBtnState();
        });

        inputGroup.appendChild(label);
        inputGroup.appendChild(input);
        surveyLoiInputsDiv.appendChild(inputGroup);

        // Initialize validation state
        surveyLoiValid[surveyId] = false;
      });
    } else {
      allSurveyIds = [];
      // Show a message that RID file is needed for survey-specific mode
      const messageDiv = document.createElement('div');
      messageDiv.className = 'form-group';
      messageDiv.innerHTML = '<p style="color: #666; font-style: italic;">Upload RID file to see survey-specific LOI inputs</p>';
      surveyLoiInputsDiv.appendChild(messageDiv);
    }
  }
}

// Update createSurveyLoiInputs to use renderLoiInputs
function createSurveyLoiInputs(surveyData) {
  renderLoiInputs(surveyData);
}

// Add LOI mode slider event
if (loiModeSlider) {
  loiModeSlider.addEventListener('change', function() {
    // Always re-render LOI inputs on mode change, regardless of RID data
    let surveyIds = {};
    if (ridData) {
      ridData.forEach(row => {
        const keys = Object.keys(row);
        const surveyidKey = keys.find(k => k.trim().toLowerCase() === 'surveyid');
        if (surveyidKey) {
          const sid = (row[surveyidKey] || '').trim();
          if (sid && sid !== '') {
            surveyIds[sid] = (surveyIds[sid] || 0) + 1;
          }
        }
      });
    }
    renderLoiInputs(surveyIds); // Pass survey data or empty object
    updateProcessBtnState();
  });
}

// Update updateProcessBtnState for new LOI container behavior
function updateProcessBtnState() {
  let disabled = false;
  let tooltip = "";

  // Validate Surveys entered Threshold - get the input element dynamically
  const surveysEnteredThresholdInput = document.getElementById('surveys_entered_threshold');
  if (surveysEnteredThresholdInput) {
    const surveysThreshold = parseInt(surveysEnteredThresholdInput.value, 10);
    if (isNaN(surveysThreshold) || surveysThreshold < 4 || surveysThreshold > 20) {
      disabled = true;
      tooltip = "Please enter a valid Surveys entered Threshold (4-20).";
    }
  }

  if (isPidOnlyMode()) {
    if (!metricsData) {
      disabled = true;
      tooltip = "Please upload the PID Metrics file.";
    }
  } else {
    // RID+PID mode checks
    if (!ridData) {
      disabled = true;
      tooltip = "Please upload the RID file.";
    } else if (!metricsData) {
      disabled = true;
      tooltip = "Please upload the PID Metrics file.";
    } else {
      // LOI validation logic based on mode and data availability
      if (isAverageLoiMode()) {
        // Average LOI mode: only validate if RID file has survey data
        if (allSurveyIds.length > 0) {
          const avgLoiInput = document.getElementById('average_loi');
          const isValid = avgLoiInput && !isNaN(parseFloat(avgLoiInput.value)) && parseFloat(avgLoiInput.value) >= 3 && parseFloat(avgLoiInput.value) <= 100;
          if (!isValid) {
            disabled = true;
            tooltip = "Please enter a valid Average LOI value (3-100).";
          }
        }
      } else {
        // Survey-specific LOI mode: validate all survey inputs if data exists
        if (allSurveyIds.length > 0) {
          const invalidSurveys = allSurveyIds.filter(sid => !surveyLoiValid[sid]);
          if (invalidSurveys.length > 0) {
            disabled = true;
            tooltip = `Please enter valid LOI values for all surveys. Missing/invalid: ${invalidSurveys.join(', ')}`;
          }
        }
      }
    }
  }

  // Update button state
  processBtn.disabled = disabled;
  processBtn.title = disabled ? tooltip : "";
  processBtn.style.cursor = disabled ? "not-allowed" : "pointer";
}

// File parsing function
function parseRIDFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const csv = e.target.result;
        const lines = csv.split('\n');
        if (lines.length > 1) {
          resolve(lines);
        } else {
          reject('No data rows found in RID file');
        }
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsText(file);
  });
}

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

        ridData = json;

        // Create survey LOI inputs if surveys found
        if (Object.keys(surveyIds).length > 0) {
          allSurveyIds = Object.keys(surveyIds); // <-- Ensure allSurveyIds is set
          renderLoiInputs(surveyIds);
        } else {
          allSurveyIds = [];
        }

        showLoiGroup(); // <-- Call after allSurveyIds is set and inputs rendered
        updateProcessBtnState();
      };
      reader.readAsText(this.files[0]);
    } catch (error) {
      console.error('Error parsing RID file:', error);
      ridData = null;
      allSurveyIds = [];
      surveyLoiValid = {};
      showLoiGroup();
      updateProcessBtnState();
    }
  } else {
    ridData = null;
    allSurveyIds = [];
    surveyLoiValid = {};
    surveyLoiInputsDiv.innerHTML = '';
    showLoiGroup();
    updateProcessBtnState();
  }
});

// Add metrics file handler after ridFileInput handler:
metricsFileInput.addEventListener('change', function(e) {
    if (this.files.length > 0) {
        try {
            const reader = new FileReader();
            reader.onload = function(e) {
                metricsData = true; // Just need to know it's uploaded
                updateProcessBtnState();
            };
            reader.readAsArrayBuffer(this.files[0]); // Use ArrayBuffer for Excel files
        } catch (error) {
            console.error('Error reading Metrics file:', error);
            metricsData = null;
            updateProcessBtnState();
        }
    } else {
        metricsData = null;
        updateProcessBtnState();
    }
});

// Update mode slider logic to handle Current session checks subgroup
modeSlider.addEventListener('change', function() {
  const ridGroup = ridFileInput.closest('.form-group');
  const loiContainer = document.getElementById('survey-loi-container');
  const speederGroup = document.getElementById('speeder-group');
  const highLoiGroup = document.getElementById('high-loi-group');
  
  // Find the Current session checks subgroup by looking for the parent .form-subgroup of speeder-group
  const currentSessionSubgroup = speederGroup ? speederGroup.closest('.form-subgroup') : null;
  const currentSessionHeader = document.getElementById('current-session-header');
  
  if (isPidOnlyMode()) {
    ridFileInput.removeAttribute('required');
    ridFileInput.disabled = true;
    ridGroup.style.display = 'none';
    ridFileInput.value = '';
    loiContainer.style.display = 'none';
    // Clear survey LOI inputs
    surveyLoiInputsDiv.innerHTML = '';
    allSurveyIds = [];
    surveyLoiValid = {};
    // Hide the entire Current session checks subgroup
    if (currentSessionSubgroup) {
      currentSessionSubgroup.style.display = 'none';
    }
    if (currentSessionHeader) {
      currentSessionHeader.style.display = 'none';
    }
  } else {
    ridFileInput.setAttribute('required', 'required');
    ridFileInput.disabled = false;
    ridGroup.style.display = '';
    // Show the entire Current session checks subgroup
    if (currentSessionSubgroup) {
      currentSessionSubgroup.style.display = '';
    }
    if (currentSessionHeader) {
      currentSessionHeader.style.display = '';
    }
  }
  showLoiGroup();
  updateProcessBtnState();
});

// --- Save and restore form values using localStorage ---
function saveFormState() {
  const form = document.getElementById('main-form');
  const elements = form.querySelectorAll('input, select, textarea');
  elements.forEach(el => {
    if (el.name === 'survey_loi') return; // Do not save survey_loi
    if (el.id === 'mode_slider') {
      // skip saving mode_slider
    } else if (el.type === 'checkbox') {
      localStorage.setItem('form_' + el.name, el.checked ? '1' : '0');
    } else if (el.type === 'file') {
      // Do not save file inputs
    } else {
      localStorage.setItem('form_' + el.name, el.value);
    }
  });
  
  // Save surveys_entered_threshold separately
  const surveysEnteredThresholdInput = document.getElementById('surveys_entered_threshold');
  if (surveysEnteredThresholdInput) {
    localStorage.setItem('surveys_entered_threshold', surveysEnteredThresholdInput.value);
  }
}

function restoreFormState() {
  const form = document.getElementById('main-form');
  const elements = form.querySelectorAll('input, select, textarea');
  elements.forEach(el => {
    if (el.name === 'survey_loi') return; // Do not restore survey_loi
    if (el.id === 'mode_slider') {
      // skip restoring mode_slider
    } else if (el.type === 'checkbox') {
      const val = localStorage.getItem('form_' + el.name);
      if (val !== null) el.checked = val === '1';
    } else if (el.type === 'file') {
      // Do not restore file inputs
    } else {
      const val = localStorage.getItem('form_' + el.name);
      if (val !== null) el.value = val;
    }
  });
  
  // Restore surveys_entered_threshold separately
  const savedThreshold = localStorage.getItem('surveys_entered_threshold');
  const surveysEnteredThresholdInput = document.getElementById('surveys_entered_threshold');
  if (savedThreshold !== null && surveysEnteredThresholdInput) {
    surveysEnteredThresholdInput.value = savedThreshold;
  }
}

// Info popup handlers
infoIcon.addEventListener('click', function(e) {
  console.log('Info icon clicked'); // Debug log
  e.preventDefault();
  infoPopup.classList.add('visible');
});

infoIcon.addEventListener('keypress', function(e) {
  if (e.key === 'Enter' || e.key === ' ') {
    e.preventDefault();
    infoPopup.classList.add('visible');
  }
});

infoPopupClose.addEventListener('click', function() {
  infoPopup.classList.remove('visible');
});

// Close popup when clicking outside
document.addEventListener('click', function(e) {
  if (!infoPopup.contains(e.target) && !infoIcon.contains(e.target)) {
    infoPopup.classList.remove('visible');
  }
});

// Close on Escape key
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape' && infoPopup.classList.contains('visible')) {
    infoPopup.classList.remove('visible');
  }
});

// Dark mode toggle logic - Fixed to ensure proper event handling
function setDarkMode(enabled) {
  document.body.classList.toggle('dark-mode', enabled);
  darkModeToggle.textContent = enabled ? '☀️ Light Mode' : '🌙 Dark Mode';
  darkModeToggle.style.background = enabled ? '#23263a' : '#f4f7fa';
  darkModeToggle.style.color = enabled ? '#e0e0e0' : '#2a0b57';
  localStorage.setItem('dark_mode', enabled ? '1' : '0');
}

// Dark mode toggle event
darkModeToggle.addEventListener('click', function(e) {
  console.log('Dark mode toggle clicked'); // Debug log
  e.preventDefault();
  e.stopPropagation();
  const isCurrentlyDark = document.body.classList.contains('dark-mode');
  setDarkMode(!isCurrentlyDark);
});

// Show shutdown button only in EXE mode
function checkEXEMode() {
  // Check if running as EXE (window title or other indicators)
  const isEXE = window.navigator.userAgent.includes('Electron') || 
               window.location.protocol === 'file:' ||
               window.location.hostname === '127.0.0.1';
  
  if (isEXE) {
    document.getElementById('shutdown-btn').style.display = 'block';
  }
}

// Shutdown function
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

// Prevent default form submission and use AJAX to handle responses.
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

// Save on change
document.getElementById('main-form').addEventListener('input', saveFormState);
document.getElementById('main-form').addEventListener('change', saveFormState);

// Add drag and drop functionality
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

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function handleFiles(files, fileInput, expectedType, processCallback) {
    if (files.length === 0) return;
    
    const file = files[0];
    const fileExtension = file.name.split('.').pop().toLowerCase();
    
    // Validate file type
    if (fileExtension !== expectedType) {
        const typeText = expectedType === 'csv' ? 'CSV' : 'Excel (.xlsx)';
        addJsMessage(`Invalid file type. Please select a ${typeText} file.`, 'error');
        return;
    }
    
    // Update the hidden file input
    const dt = new DataTransfer();
    dt.items.add(file);
    fileInput.files = dt.files;
    
    // Process the file
    processCallback(file, fileInput);
}

function processRidFile(file, fileInput) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
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

            ridData = json;

            // Show file preview
            showFilePreview('rid', file, json.length, Object.keys(surveyIds).length);

            // Always render LOI inputs (with or without survey data)
            renderLoiInputs(surveyIds);

            // Always show LOI group
            showLoiGroup();
            updateProcessBtnState();
            clearJsMessages();
        } catch (error) {
            console.error('Error parsing RID file:', error);
            addJsMessage('Error processing RID file. Please ensure it\'s a valid CSV format.', 'error');
            removeFile('rid');
        }
    };
    reader.readAsText(file);
}

function processPidFile(file, fileInput) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            metricsData = true;
            
            // Try to read Excel file to get row count
            const workbook = XLSX.read(e.target.result, {type: 'arraybuffer'});
            let rowCount = 0;
            let uniquePids = 0;
            
            if (workbook.SheetNames.includes('Marketplace Metrics by PID')) {
                const sheet = workbook.Sheets['Marketplace Metrics by PID'];
                const json = XLSX.utils.sheet_to_json(sheet, {range: 5}); // Skip first 5 rows
                rowCount = json.length;
                
                // Count unique PIDs
                const pids = new Set();
                json.forEach(row => {
                    const pidKey = Object.keys(row).find(k => k.toString().toLowerCase().includes('pid'));
                    if (pidKey && row[pidKey]) {
                        pids.add(row[pidKey].toString());
                    }
                });
                uniquePids = pids.size;
            }
            
            // Show file preview
            showFilePreview('pid', file, rowCount, uniquePids);
            updateProcessBtnState();
            clearJsMessages();
        } catch (error) {
            console.error('Error reading Metrics file:', error);
            addJsMessage('Error processing PID file. Please ensure it\'s a valid Excel format.', 'error');
            removeFile('pid');
        }
    };
    reader.readAsArrayBuffer(file);
}

function showFilePreview(type, file, totalRecords, uniqueCount) {
    const previewId = type === 'rid' ? 'rid-file-preview' : 'pid-file-preview';
    const dropZoneId = type === 'rid' ? 'rid-drop-zone' : 'pid-drop-zone';
    
    const preview = document.getElementById(previewId);
    const dropZone = document.getElementById(dropZoneId);
    
    if (preview && dropZone) {
        // Hide drop zone and show preview
        dropZone.style.display = 'none';
        preview.style.display = 'block';
        
        // Update preview content
        const fileName = preview.querySelector('.file-name');
        const fileStats = preview.querySelector('.file-stats');
        
        fileName.textContent = file.name;
        
        const sizeText = formatFileSize(file.size);
        const recordType = type === 'rid' ? 'RIDs' : 'PIDs';
        const uniqueText = type === 'rid' ? 'Unique Survey IDs' : 'Unique PIDs';
        
        fileStats.innerHTML = `${sizeText} • ${totalRecords} records • ${uniqueCount} ${uniqueText}`;
    }
}

function removeFile(type) {
    const previewId = type === 'rid' ? 'rid-file-preview' : 'pid-file-preview';
    const dropZoneId = type === 'rid' ? 'rid-drop-zone' : 'pid-drop-zone';
    const fileInputId = type === 'rid' ? 'rid_file' : 'metrics_file';
    
    const preview = document.getElementById(previewId);
    const dropZone = document.getElementById(dropZoneId);
    const fileInput = document.getElementById(fileInputId);
    
    if (preview && dropZone && fileInput) {
        // Clear file input
        fileInput.value = '';
        
        // Show drop zone and hide preview
        preview.style.display = 'none';
        dropZone.style.display = 'flex';
        
        // Reset data
        if (type === 'rid') {
            ridData = null;
            allSurveyIds = [];
            surveyLoiValid = {};
            surveyLoiInputsDiv.innerHTML = '';
            showLoiGroup();
        } else {
            metricsData = null;
        }
        
        updateProcessBtnState();
    }
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

// Update mode slider logic to show/hide RID drop zone
modeSlider.addEventListener('change', function() {
    const ridDropZone = document.getElementById('rid-drop-zone');
    const ridPreview = document.getElementById('rid-file-preview');
    const loiContainer = document.getElementById('survey-loi-container');
    const speederGroup = document.getElementById('speeder-group');
    const highLoiGroup = document.getElementById('high-loi-group');
    
    // Find the Current session checks subgroup by looking for the parent .form-subgroup of speeder-group
    const currentSessionSubgroup = speederGroup ? speederGroup.closest('.form-subgroup') : null;
    const currentSessionHeader = document.getElementById('current-session-header');
    
    if (isPidOnlyMode()) {
        ridFileInput.removeAttribute('required');
        ridFileInput.disabled = true;
        if (ridDropZone) ridDropZone.style.display = 'none';
        if (ridPreview) ridPreview.style.display = 'none';
        ridFileInput.value = '';
        loiContainer.style.display = 'none';
        // Clear survey LOI inputs
        surveyLoiInputsDiv.innerHTML = '';
        allSurveyIds = [];
        surveyLoiValid = {};
        // Hide the entire Current session checks subgroup
        if (currentSessionSubgroup) {
            currentSessionSubgroup.style.display = 'none';
        }
        if (currentSessionHeader) {
            currentSessionHeader.style.display = 'none';
        }
    } else {
        ridFileInput.setAttribute('required', 'required');
        ridFileInput.disabled = false;
        if (ridDropZone) ridDropZone.style.display = 'flex';
        // Show the entire Current session checks subgroup
        if (currentSessionSubgroup) {
            currentSessionSubgroup.style.display = '';
        }
        if (currentSessionHeader) {
            currentSessionHeader.style.display = '';
        }
    }
    showLoiGroup();
    updateProcessBtnState();
});

// Initialize everything when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
  console.log('DOM loaded, initializing...'); // Debug log
  
  // Initialize mode
  modeSlider.checked = false;  // Default to RID+PID mode
  
  // Set LOI mode slider to checked (Average LOI mode by default)
  if (loiModeSlider) {
    loiModeSlider.checked = true;
  }
  
  // Always show survey LOI container and initialize with default Average LOI input
  surveyLoiContainer.style.display = 'block';
  renderLoiInputs({}); // Initialize with empty survey data
  
  // Restore form state
  restoreFormState();
  
  // Restore dark mode preference
  const darkPref = localStorage.getItem('dark_mode');
  setDarkMode(darkPref === '1');
  
  // Check EXE mode
  checkEXEMode();
  
  // Update process button state
  updateProcessBtnState();
  
  // Initialize drag and drop
  initializeDragAndDrop();
  
  console.log('Initialization complete'); // Debug log
});