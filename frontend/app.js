// DOM Elements
const promptInput = document.getElementById('prompt-input');
const runBtn = document.getElementById('run-btn');
const agentStatusBadge = document.getElementById('agent-status');
const consoleLogs = document.getElementById('console-logs');
const dataPreviewBody = document.getElementById('data-preview-body');
const openExcelBtn = document.getElementById('open-excel-btn');
const openSheetsLink = document.getElementById('open-sheets-link');
const actionsPanel = document.getElementById('actions-panel');

// State trackers
let pollingInterval = null;
let displayedLogCount = 0;

// API Endpoints
const API_RUN = '/api/run';
const API_STATE = '/api/state';
const API_OPEN_EXCEL = '/api/excel/open';

// Set up event listeners
runBtn.addEventListener('click', startAgentWorkflow);
openExcelBtn.addEventListener('click', openExcelLocally);

// Allow pressing Enter in the input box to trigger the agent
promptInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !runBtn.disabled) startAgentWorkflow();
});

// Initialize UI — keep button enabled on first load
resetUI();
runBtn.disabled = false;  // Always start with button enabled

/**
 * Initiates the agent workflow by posting the prompt to the backend
 */
async function startAgentWorkflow() {
    const promptText = promptInput.value.trim();
    if (!promptText) return;

    try {
        resetUI();
        runBtn.disabled = true; // Disable ONLY while the agent is actively running
        setAgentStatus('planning');
        appendLogLine('[System] Dispatching user request to agent...', 'system-line');

        const response = await fetch(API_RUN, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ prompt: promptText })
        });

        if (!response.ok) {
            const errData = await response.json();
            throw new Error(errData.detail || 'Failed to start agent');
        }

        appendLogLine('[System] Agent started. Commencing execution monitoring...', 'system-line');
        
        // Start polling for state updates every 500ms
        pollingInterval = setInterval(pollAgentState, 500);
        
    } catch (error) {
        appendLogLine(`[Error] Failed to initialize agent: ${error.message}`, 'error-line');
        setAgentStatus('failed');
        runBtn.disabled = false;
    }
}

/**
 * Periodically polls the backend for agent state, logs, and data
 */
async function pollAgentState() {
    try {
        const response = await fetch(API_STATE);
        if (!response.ok) throw new Error('Failed to retrieve state');
        
        const state = await response.json();
        
        // 1. Update Agent Status Badge
        setAgentStatus(state.status);
        
        // 2. Update Console Logs
        updateLogs(state.logs);
        
        // 3. Update Stepper Statuses
        updateStepper(state.steps);
        
        // 4. Update CSV Data Table Preview
        if (state.csv_data && state.csv_data.length > 0) {
            updateDataPreview(state.csv_data);
        }
        
        // 5. Update Spreadsheet action buttons
        updateSpreadsheetLinks(state.excel_path, state.google_sheets_url);
        
        // 6. Handle Completion / Failures
        if (state.status === 'completed' || state.status === 'failed') {
            clearInterval(pollingInterval);
            pollingInterval = null;
            runBtn.disabled = false;

            if (state.status === 'completed') {
                appendLogLine('[System] Workflow completed successfully. Spreadsheets are now accessible.', 'system-line');
            } else {
                const errMsg = state.error || 'Unknown error occurred.';
                appendLogLine(`[System] Workflow failed: ${errMsg}`, 'error-line');
            }
        }

        // Always show action panel if we have results (regardless of status)
        if (state.excel_path || state.google_sheets_url) {
            actionsPanel.style.display = 'block';
        }
        
    } catch (error) {
        console.error('Polling error:', error);
        appendLogLine(`[System Error] Lost communication with backend: ${error.message}`, 'error-line');
    }
}

/**
 * Triggers Excel opening command on the backend
 */
async function openExcelLocally() {
    try {
        appendLogLine('[System] Requesting host to launch Microsoft Excel...', 'system-line');
        const response = await fetch(API_OPEN_EXCEL, { method: 'POST' });
        if (!response.ok) {
            const errData = await response.json();
            throw new Error(errData.detail || 'Could not open workbook');
        }
        const data = await response.json();
        appendLogLine(`[System] Excel launched successfully: ${data.path || ''}`, 'system-line');
    } catch (error) {
        appendLogLine(`[Error] Failed to open Excel: ${error.message}`, 'error-line');
        alert(`Could not open Excel automatically.\n\nOpen it manually at:\nC:\\Users\\HP\\Downloads\\assignment_genai\\employees.xlsx`);
    }
}

/**
 * Resets the UI components to prepare for a fresh agent run
 */
function resetUI() {
    runBtn.disabled = false; // Button stays enabled; only disabled during active runs
    displayedLogCount = 0;
    consoleLogs.innerHTML = '';
    actionsPanel.style.display = 'none';
    
    openExcelBtn.disabled = true;
    openSheetsLink.href = '#';
    openSheetsLink.classList.add('disabled');
    
    // Clear data table
    dataPreviewBody.innerHTML = `
        <tr>
            <td colspan="5" class="empty-state">No data generated yet. Click "Run Agent" to start.</td>
        </tr>
    `;
    
    // Reset Stepper
    const stepIds = ['csv', 'excel', 'sheets', 'verify'];
    stepIds.forEach(id => {
        const stepEl = document.getElementById(`step-${id}`);
        const descEl = document.getElementById(`desc-${id}`);
        
        stepEl.className = 'step';
        descEl.textContent = 'Pending agent initialization.';
    });
}

/**
 * Sets the main agent status indicator badge
 */
function setAgentStatus(status) {
    agentStatusBadge.textContent = status;
    agentStatusBadge.className = `status-badge ${status}`;
}

/**
 * Compares incoming logs with already rendered logs and appends new ones
 */
function updateLogs(logs) {
    if (!logs || logs.length <= displayedLogCount) return;
    
    for (let i = displayedLogCount; i < logs.length; i++) {
        const log = logs[i];
        let className = 'agent-line';
        
        if (log.startsWith('[System]') || log.startsWith('Received prompt:')) {
            className = 'system-line';
        } else if (log.startsWith('Success:') || log.startsWith('Tool generate_') || log.startsWith('Tool import_')) {
            className = 'tool-line';
        } else if (log.startsWith('Error:') || log.startsWith('Fatal') || log.startsWith('Unhandled')) {
            className = 'error-line';
        }
        
        appendLogLine(log, className);
    }
    
    displayedLogCount = logs.length;
}

/**
 * Appends a log line to the scrolling console box
 */
function appendLogLine(text, className) {
    const line = document.createElement('div');
    line.className = `log-line ${className}`;
    line.textContent = text;
    consoleLogs.appendChild(line);
    
    // Scroll to bottom
    consoleLogs.scrollTop = consoleLogs.scrollHeight;
}

/**
 * Updates the state visualizer stepper in the center of the UI
 */
function updateStepper(steps) {
    if (!steps) return;
    
    steps.forEach(step => {
        const stepEl = document.getElementById(`step-${step.id}`);
        const descEl = document.getElementById(`desc-${step.id}`);
        
        if (stepEl && descEl) {
            // Map status values to classes: pending, running, completed, failed
            let classSuffix = '';
            if (step.status === 'running') classSuffix = 'active';
            else if (step.status === 'completed') classSuffix = 'success';
            else if (step.status === 'failed') classSuffix = 'failed';
            
            stepEl.className = `step ${classSuffix}`;
            
            // Set message
            if (step.status === 'pending') {
                descEl.textContent = 'Awaiting execution...';
            } else {
                descEl.textContent = step.message || 'Executing step...';
            }
        }
    });
}

/**
 * Populates the interactive data table preview with employee data rows
 */
function updateDataPreview(data) {
    // Only update if the table is currently empty or has the placeholder
    if (dataPreviewBody.querySelector('.empty-state') || dataPreviewBody.children.length !== data.length) {
        dataPreviewBody.innerHTML = '';
        
        const currencyFormatter = new Intl.NumberFormat('en-US', {
            style: 'currency',
            currency: 'USD',
            maximumFractionDigits: 0
        });
        
        data.forEach(emp => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td><strong>${emp['Employee ID']}</strong></td>
                <td>${emp['Name']}</td>
                <td>${emp['Department']}</td>
                <td>${emp['Email']}</td>
                <td style="text-align: right;">${currencyFormatter.format(emp['Salary'])}</td>
            `;
            dataPreviewBody.appendChild(tr);
        });
    }
}

/**
 * Activates spreadsheet links when filepaths or urls become available
 */
function updateSpreadsheetLinks(excelPath, sheetsUrl) {
    if (excelPath) {
        openExcelBtn.disabled = false;
        openExcelBtn.title = excelPath;
    }
    
    if (sheetsUrl) {
        openSheetsLink.href = sheetsUrl;
        openSheetsLink.target = '_blank';       // open in new browser tab
        openSheetsLink.rel = 'noopener noreferrer';
        openSheetsLink.classList.remove('disabled');
    }
}
