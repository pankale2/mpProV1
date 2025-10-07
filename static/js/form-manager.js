// Form Manager Module - Handles form validation, LOI management, UI components
window.FormManager = (function() {
    'use strict';

    // Private variables
    let allSurveyIds = [];
    let surveyLoiValid = {};

    // DOM elements
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

    // Private functions
    function isPidOnlyMode() { return modeSlider.checked; }

    function isAverageLoiMode() {
        return loiModeSlider && loiModeSlider.checked;
    }

    function validateSurveyLoi(surveyId, value) {
        const numValue = parseFloat(value);
        const isValid = !isNaN(numValue) && numValue >= 3 && numValue <= 100;
        surveyLoiValid[surveyId] = isValid;
        return isValid;
    }

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

    function showLoiGroup() {
        // Always show the container - no conditional logic
        surveyLoiContainer.style.display = 'block';
    }

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
            if (!window.FileManager.getMetricsData()) {
                disabled = true;
                tooltip = "Please upload the PID Metrics file.";
            }
        } else {
            // RID+PID mode checks
            if (!window.FileManager.getRidData()) {
                disabled = true;
                tooltip = "Please upload the RID file.";
            } else if (!window.FileManager.getMetricsData()) {
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

    // Dark mode functions
    function setDarkMode(enabled) {
        document.body.classList.toggle('dark-mode', enabled);
        darkModeToggle.textContent = enabled ? '☀️ Light Mode' : '🌙 Dark Mode';
        darkModeToggle.style.background = enabled ? '#23263a' : '#f4f7fa';
        darkModeToggle.style.color = enabled ? '#e0e0e0' : '#2a0b57';
        localStorage.setItem('dark_mode', enabled ? '1' : '0');
    }

    // Form state persistence
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

    function setupEventListeners() {
        // LOI mode slider event
        if (loiModeSlider) {
            loiModeSlider.addEventListener('change', function() {
                // Always re-render LOI inputs on mode change, regardless of RID data
                let surveyIds = {};
                if (window.FileManager.getRidData()) {
                    window.FileManager.getRidData().forEach(row => {
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

        // Mode slider logic
        modeSlider.addEventListener('change', function() {
            const ridFileInput = document.getElementById('rid_file');
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
                const ridFileInput = document.getElementById('rid_file');
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
                const ridFileInput = document.getElementById('rid_file');
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

        // Dark mode toggle event
        darkModeToggle.addEventListener('click', function(e) {
            console.log('Dark mode toggle clicked'); // Debug log
            e.preventDefault();
            e.stopPropagation();
            const isCurrentlyDark = document.body.classList.contains('dark-mode');
            setDarkMode(!isCurrentlyDark);
        });

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

        // Form state saving
        const form = document.getElementById('main-form');
        form.addEventListener('input', saveFormState);
        form.addEventListener('change', saveFormState);
    }

    // Public API
    return {
        // Data access
        getAllSurveyIds: () => allSurveyIds,
        getSurveyLoiValid: () => surveyLoiValid,
        resetRidData: () => {
            allSurveyIds = [];
            surveyLoiValid = {};
            surveyLoiInputsDiv.innerHTML = '';
        },
        
        // Form management
        renderLoiInputs,
        updateProcessBtnState,
        showLoiGroup,
        isPidOnlyMode,
        isAverageLoiMode,
        
        // Form state
        saveFormState,
        restoreFormState,
        
        // Dark mode
        setDarkMode,
        
        // Initialization
        setupEventListeners
    };
})();
