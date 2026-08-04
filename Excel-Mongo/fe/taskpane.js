/**
 * File: taskpane.js
 * Description: Main logic for the Excel taskpane Add-in. Handles UI interactions, Excel API calls, and backend communication.
 * Dependencies: Office.js
 */
// Local dev uses fe/config.js. Production builds generate this value from ADDIN_API_BASE.
let API_BASE = (window.EXCEL_MONGO_CONFIG && window.EXCEL_MONGO_CONFIG.apiBase) || window.location.origin;

let currentToken = "";

// WebSocket Manager
let wsSocket = null;
const wsPendingRequests = new Map();
const wsStreamHandlers = new Map();

function connectWebSocket() {
    return new Promise((resolve, reject) => {
        if (wsSocket && wsSocket.readyState === WebSocket.OPEN) {
            return resolve();
        }
        
        let wsProtocol = window.location.protocol === 'https:' || API_BASE.startsWith('https') ? 'wss://' : 'ws://';
        let wsHost = API_BASE.replace('http://', '').replace('https://', '');
        let wsUrl = `${wsProtocol}${wsHost}/ws?token=${encodeURIComponent(currentToken)}&base_url=${encodeURIComponent(currentFinnotoBase)}`;
        
        wsSocket = new WebSocket(wsUrl);
        
        wsSocket.onopen = () => {
            console.log("WebSocket connected");
            resolve();
        };
        
        wsSocket.onerror = (err) => {
            console.error("WebSocket error", err);
            reject(err);
        };
        
        wsSocket.onclose = () => {
            console.log("WebSocket closed");
            wsSocket = null;
        };
        
        wsSocket.onmessage = (event) => {
            const data = JSON.parse(event.data);
            const reqId = data.requestId;
            
            if (["chunk", "sheet_start", "sheet_complete", "progress"].includes(data.status) && wsStreamHandlers.has(reqId)) {
                wsStreamHandlers.get(reqId)(data.data, data.status);
            } else if (wsPendingRequests.has(reqId)) {
                const { res, rej } = wsPendingRequests.get(reqId);
                wsPendingRequests.delete(reqId);
                wsStreamHandlers.delete(reqId);
                
                if (data.status === "error") {
                    rej(new Error(data.error || "Unknown error"));
                } else {
                    res(data.data);
                }
            }
        };
    });
}

function sendWsRequest(action, payload, onChunk = null) {
    return new Promise(async (resolve, reject) => {
        try {
            await connectWebSocket();
        } catch (e) {
            return reject(e);
        }
        
        const reqId = crypto.randomUUID();
        wsPendingRequests.set(reqId, { res: resolve, rej: reject });
        if (onChunk) {
            wsStreamHandlers.set(reqId, onChunk);
        }
        
        wsSocket.send(JSON.stringify({
            requestId: reqId,
            action: action,
            payload: payload
        }));
    });
}

class WsResponse {
    constructor(data, ok = true, status = 200) {
        this._data = data;
        this.ok = ok;
        this.status = status;
    }
    async json() { return this._data; }
}

function getAuthHeaders(existingHeaders = {}) {
    return {
        "Authorization": `Bearer ${currentToken}`,
        ...existingHeaders
    };
}

// Fetch with a timeout so we fail fast instead of hanging
async function fetchWithTimeout(url, options = {}, timeoutMs = 60000) {
    if (url.endsWith('/health') || url.endsWith('/create_collection') || url.endsWith('/reports/list')) {
        const controller = new AbortController();
        const timer = setTimeout(() => controller.abort(), timeoutMs);
        try {
            const res = await fetch(url, { ...options, signal: controller.signal });
            return res;
        } finally {
            clearTimeout(timer);
        }
    }
    
    const urlObj = new URL(url);
    const action = urlObj.pathname.split('/').pop();
    let payload = options.body ? JSON.parse(options.body) : {};
    
    for (const [k, v] of urlObj.searchParams.entries()) {
        payload[k] = v;
    }
    
    try {
        const data = await sendWsRequest(action, payload);
        return new WsResponse(data);
    } catch(err) {
        return new WsResponse({detail: err.message}, false, 500);
    }
}

// DOM Elements
const connectionStatus = document.getElementById("connection-status");
const collectionSelect = document.getElementById("collection-select");
const btnRefreshCollections = document.getElementById("btn-refresh-collections");
const btnImportSchema = document.getElementById("btn-import-schema");
const schemaBar = document.getElementById("schema-bar");
const schemaBarFields = document.getElementById("schema-bar-fields");
const queryInput = document.getElementById("query-input");
const limitInput = document.getElementById("limit-input");
const jsonValidity = document.getElementById("json-validity");
const btnRun = document.getElementById("btn-run");
const btnText = btnRun ? btnRun.querySelector(".btn-text") : null;
const btnSpinner = btnRun ? btnRun.querySelector(".btn-spinner") : null;
const feedbackPanel = document.getElementById("feedback-panel");
const feedbackMessage = document.getElementById("feedback-message");
const statFetched = document.getElementById("stat-fetched");
const statInserted = document.getElementById("stat-inserted");
const statUpdated = document.getElementById("stat-updated");

// State
let isApiOnline = false;
let currentSchema = []; // field names currently in the sheet's row 1

// Sync Control State
let isSyncCancelled = false;
let activeAbortController = null;

// ── Query Builder State ──
let guiLogic = 'and';       // 'and' | 'or'
let conditionCounter = 0;   // unique ID for each row

const GUI_OPERATORS = [
    { label: '=  equals',         value: 'eq',      hasValue: true  },
    { label: '\u2260  not equals',  value: 'ne',      hasValue: true  },
    { label: '>  greater than',   value: 'gt',      hasValue: true  },
    { label: '\u2265  greater or =', value: 'gte',   hasValue: true  },
    { label: '<  less than',      value: 'lt',      hasValue: true  },
    { label: '\u2264  less or =',    value: 'lte',   hasValue: true  },
    { label: '~  contains',       value: 'regex',   hasValue: true  },
    { label: '^  starts with',    value: 'starts',  hasValue: true  },
    { label: '\u2208  in (a,b,c)', value: 'in',      hasValue: true  },
    { label: '\u2209  not in',     value: 'nin',     hasValue: true  },
    { label: '\u2713  exists',     value: 'exists',  hasValue: false },
    { label: '\u2717  not exists', value: 'nexists', hasValue: false },
];

// Initialize Add-in
const FINNOTO_BASE_PRESETS = {
    eapi: "https://eapi.finnoto.dev",
    abdebug: "https://abdebug.finnoto.dev",
    arcapi: "https://arcapi.finnoto.dev",
};

const FINNOTO_PRESET_LABELS = {
    eapi: "eapi.finnoto.dev",
    abdebug: "abdebug.finnoto.dev",
    arcapi: "arcapi.finnoto.dev",
};

let _selectedEnvKey = "eapi";

Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        // Automatically open the taskpane next time this document is opened
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
        Office.context.document.settings.saveAsync();

        // Restore credentials
        currentToken = Office.context.document.settings.get('FinnotoToken') || '';
        currentFinnotoBase = Office.context.document.settings.get('FinnotoBase') || 'https://eapi.finnoto.dev';

        // Restore selected env preset
        const savedEnv = Office.context.document.settings.get('FinnotoEnv') || 'eapi';
        _selectedEnvKey = savedEnv;
        if (FINNOTO_BASE_PRESETS[savedEnv]) {
            currentFinnotoBase = FINNOTO_BASE_PRESETS[savedEnv];
        }

        const savedApiBase = Office.context.document.settings.get('ApiBase');
        if (savedApiBase) {
            API_BASE = savedApiBase;
            document.getElementById("api-base-input").value = API_BASE;
            document.getElementById("use-custom-backend-checkbox").checked = true;
            document.getElementById("custom-backend-container").classList.remove("hidden");
        } else {
            document.getElementById("api-base-input").value = "";
        }

        // ── Environment preset buttons ──
        document.querySelectorAll('.env-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.env-btn').forEach(b => b.classList.remove('env-active'));
                btn.classList.add('env-active');
                _selectedEnvKey = btn.dataset.env;
                currentFinnotoBase = FINNOTO_BASE_PRESETS[_selectedEnvKey];
            });
        });
        
        document.getElementById('use-custom-backend-checkbox').addEventListener('change', (e) => {
            if (e.target.checked) {
                document.getElementById('custom-backend-container').classList.remove('hidden');
            } else {
                document.getElementById('custom-backend-container').classList.add('hidden');
            }
        });

        document.getElementById('custom-env-select').addEventListener('change', (e) => {
            if (e.target.value) {
                document.getElementById('api-base-input').value = e.target.value;
            }
        });
        
        if (currentToken) {
            document.getElementById('login-view').classList.add('hidden');
            document.getElementById('app-view').classList.remove('hidden');
            await initApp();
        } else {
            document.getElementById('login-view').classList.remove('hidden');
            document.getElementById('app-view').classList.add('hidden');
        }
    }
});

// App Initialization
async function initApp() {
    // Wire new report UI
    document.getElementById('btn-refresh-reports').addEventListener('click', loadReports);
    document.getElementById('report-search').addEventListener('input', filterReportDropdown);
    document.getElementById('report-select').addEventListener('change', onReportChange);

    // Date preset chips
    document.querySelectorAll('.date-preset-btn').forEach(btn => {
        btn.addEventListener('click', () => onDatePreset(btn.dataset.preset));
    });

    // Sheet select-all / none
    document.getElementById('btn-select-all-sheets').addEventListener('click', () => {
        document.querySelectorAll('#sheet-checkboxes input[type=checkbox]').forEach(cb => cb.checked = true);
    });
    document.getElementById('btn-deselect-all-sheets').addEventListener('click', () => {
        document.querySelectorAll('#sheet-checkboxes input[type=checkbox]').forEach(cb => cb.checked = false);
    });

    // Custom filter add
    document.getElementById('btn-add-filter').addEventListener('click', addFilterRow);

    // Main fetch button
    btnRun.addEventListener('click', runReportFetch);

    // Boot
    await checkApiHealth();
    if (isApiOnline) await loadReports();

    // Register workbook activation event and load current state
    await Excel.run(async (context) => {
        context.workbook.worksheets.onActivated.add(onWorksheetActivated);
        await context.sync();
    }).catch(console.error);
    await loadSheetState();
    
    if (schemaBar) await refreshSchemaBar();
}

async function saveSheetState() {
    const col = (collectionSelect && collectionSelect.value) || "";
    const query = (queryInput && queryInput.value) || "";
    
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.customProperties.add("syncCollection", col);
        sheet.customProperties.add("syncQuery", query);
        await context.sync();
    }).catch(console.error);
}

async function loadSheetState() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const colProp = sheet.customProperties.getItemOrNullObject("syncCollection");
        const queryProp = sheet.customProperties.getItemOrNullObject("syncQuery");
        colProp.load("value");
        queryProp.load("value");
        await context.sync();
        
        if (!colProp.isNullObject && colProp.value && collectionSelect) {
            collectionSelect.value = colProp.value;
        }
        if (!queryProp.isNullObject && queryProp.value && queryInput) {
            queryInput.value = queryProp.value;
            validateQuerySyntax();
        }
    }).catch(console.error);
}

async function onWorksheetActivated(event) {
    await loadSheetState();
    if (schemaBar) await refreshSchemaBar();
}

async function ensureTargetSheet(collection) {
    if (!collection) return;
    await Excel.run(async (context) => {
        const activeSheet = context.workbook.worksheets.getActiveWorksheet();
        activeSheet.load("name");
        
        const colProp = activeSheet.customProperties.getItemOrNullObject("syncCollection");
        colProp.load("value");
        
        await context.sync();
        
        let activeCollection = null;
        if (!colProp.isNullObject && colProp.value) {
            activeCollection = colProp.value;
        }

        let safeName = collection.substring(0, 31).replace(/[\\/?*\[\]:]/g, '_');

        if (!activeCollection || activeCollection === collection) {
            try {
                if (activeSheet.name !== safeName) {
                    activeSheet.name = safeName;
                    await context.sync();
                }
            } catch (e) {}
            activeSheet.customProperties.add("syncCollection", collection);
            await context.sync();
        } else {
            let existingSheet = context.workbook.worksheets.getItemOrNullObject(safeName);
            await context.sync();
            
            if (!existingSheet.isNullObject) {
                existingSheet.activate();
                existingSheet.customProperties.add("syncCollection", collection);
                await context.sync();
            } else {
                let newSheet = context.workbook.worksheets.add();
                try {
                    newSheet.name = safeName;
                    await context.sync();
                } catch(e) {}
                newSheet.activate();
                newSheet.customProperties.add("syncCollection", collection);
                await context.sync();
            }
        }
    }).catch(console.error);
}

// Check if the Finnoto Proxy backend is reachable
async function checkApiHealth() {
    try {
        updateStatus("checking", "Checking API...");
        const res = await fetchWithTimeout(`${API_BASE}/health`);
        const data = await res.json();
        
        if (data.status === "ok") {
            try {
                await connectWebSocket();
                isApiOnline = true;
                updateStatus("online", "Connected to Finnoto");
                if (btnRun) btnRun.disabled = false;
            } catch (wsErr) {
                throw new Error("WebSocket connection failed");
            }
        } else {
            throw new Error("Health check returned status " + data.status);
        }
    } catch (err) {
        isApiOnline = false;
        updateStatus("offline", "API Offline");
        if (btnRun) btnRun.disabled = true;
        const msg = err.name === "AbortError"
            ? "Connection timed out. Make sure ./run.sh is running in your terminal."
            : "Backend API is offline. Run ./run.sh in your terminal first.";
        showError(msg);
    }
}

// ─────────────────────────────────────────────────────────────
// REPORT LOADING & FILTERING
// ─────────────────────────────────────────────────────────────

let _allReports = [];       // raw { id, name, sheets:[{id,name}] } list
let _allReportsBase = null; // base URL the current report list came from
let _selectedPreset = null; // currently active date preset key
let currentFinnotoBase = "https://eapi.finnoto.dev";

async function loadReports() {
    const select = document.getElementById('report-select');
    select.disabled = true;
    select.innerHTML = '<option value="">Loading reports...</option>';
    document.getElementById('sheet-selector-section').classList.add('hidden');

    try {
        const res = await fetchWithAuth(`${API_BASE}/reports/list`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ page: 1 })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || 'Failed to load reports');

        _allReports = data.reports || [];
        _allReportsBase = data._base_url || currentFinnotoBase;
        if (_allReports.length === 0) {
            const tried = (data._attempted || [])
                .map(a => a.base_url.replace('https://', ''))
                .join(', ');
            select.innerHTML = `<option value="">No reports found on ${tried || currentFinnotoBase}</option>`;
            select.disabled = true;
            btnRun.disabled = true;
            return;
        }
        renderReportOptions(_allReports);
    } catch (err) {
        showError('Failed to load Finnoto reports: ' + err.message);
        select.innerHTML = '<option value="">Error loading reports</option>';
    }
}

function renderReportOptions(reports) {
    const select = document.getElementById('report-select');
    select.innerHTML = '<option value="">— select a report —</option>';
    reports.forEach(r => {
        const opt = document.createElement('option');
        opt.value = r.id;
        const dateStr = r.created_at ? new Date(r.created_at).toLocaleDateString() : '';
        const statusTag = r.url ? '' : ' [processing]';
        opt.textContent = `${r.name} — ${dateStr}${statusTag}`;
        opt.dataset.reportUrl = r.url || '';
        opt.dataset.reportName = r.name || '';
        select.appendChild(opt);
    });
    select.disabled = reports.length === 0;
    btnRun.disabled = true;
}

function filterReportDropdown() {
    const q = document.getElementById('report-search').value.toLowerCase().trim();
    const filtered = q ? _allReports.filter(r => (r.name || '').toLowerCase().includes(q)) : _allReports;
    renderReportOptions(filtered);
}

function onReportChange() {
    const select = document.getElementById('report-select');
    const selected = select.options[select.selectedIndex];
    const sheetSection = document.getElementById('sheet-selector-section');
    const sheetCheckboxes = document.getElementById('sheet-checkboxes');

    if (!selected || !selected.value) {
        sheetSection.classList.add('hidden');
        btnRun.disabled = true;
        return;
    }

    const reportId = parseInt(selected.value);
    const report = _allReports.find(r => r.id === reportId);
    const sheets = (report && report.sheets) || [];

    if (sheets.length === 0) {
        sheetSection.classList.add('hidden');
        btnRun.disabled = false;
        return;
    }

    // Populate sheet checkboxes
    sheetCheckboxes.innerHTML = '';
    sheets.forEach(s => {
        const label = document.createElement('label');
        label.className = 'flex items-center gap-3 cursor-pointer p-2 hover:bg-surface-container-lowest transition-colors';
        label.innerHTML = `
            <input type="checkbox" class="sheet-cb w-4 h-4 border-2 border-black accent-[#FFD300]" value="${s.id}" checked />
            <span class="font-data-field text-[13px]">${s.name}</span>
        `;
        sheetCheckboxes.appendChild(label);
    });

    sheetSection.classList.remove('hidden');
    btnRun.disabled = false;
}

// ─────────────────────────────────────────────────────────────
// DATE FILTER HELPERS
// ─────────────────────────────────────────────────────────────

function onDatePreset(preset) {
    document.querySelectorAll('.date-preset-btn').forEach(b => b.classList.remove('active'));
    const btn = document.querySelector(`.date-preset-btn[data-preset="${preset}"]`);

    if (_selectedPreset === preset) {
        // Toggle off
        _selectedPreset = null;
        document.getElementById('date-custom-range').classList.add('hidden');
        return;
    }
    _selectedPreset = preset;
    if (btn) btn.classList.add('active');
    document.getElementById('date-custom-range').classList.toggle('hidden', preset !== 'custom');
}

function buildDatePayload() {
    if (!_selectedPreset) return undefined;
    if (_selectedPreset === 'custom') {
        const min = document.getElementById('date-min').value;
        const max = document.getElementById('date-max').value;
        if (!min && !max) return undefined;
        return { range: { min, max } };
    }
    return { [_selectedPreset]: true };
}

// ─────────────────────────────────────────────────────────────
// CUSTOM FILTERS (key → value rows)
// ─────────────────────────────────────────────────────────────

function addFilterRow(key = '', value = '') {
    const container = document.getElementById('filters-rows');
    const row = document.createElement('div');
    row.className = 'filter-row flex gap-2 items-center';
    row.innerHTML = `
        <input type="text" placeholder="key" value="${key}" class="filter-key flex-1 bg-surface-container-lowest border-border-width border-black p-2 font-data-field text-[12px] text-on-background neo-shadow-sm focus:outline-none" />
        <span class="font-label-mono text-[12px] opacity-40">→</span>
        <input type="text" placeholder="value" value="${value}" class="filter-val flex-1 bg-surface-container-lowest border-border-width border-black p-2 font-data-field text-[12px] text-on-background neo-shadow-sm focus:outline-none" />
        <button class="remove-filter-btn p-1 opacity-50 hover:opacity-100" title="Remove">
            <span class="material-symbols-outlined text-[16px]">close</span>
        </button>`;
    row.querySelector('.remove-filter-btn').onclick = () => row.remove();
    container.appendChild(row);
}

function buildFiltersPayload() {
    const rows = document.querySelectorAll('.filter-row');
    if (rows.length === 0) return undefined;
    const out = {};
    rows.forEach(row => {
        const k = row.querySelector('.filter-key').value.trim();
        const v = row.querySelector('.filter-val').value.trim();
        if (k) {
            // Auto-coerce numbers
            out[k] = isNaN(v) || v === '' ? v : Number(v);
        }
    });
    return Object.keys(out).length > 0 ? out : undefined;
}

// ─────────────────────────────────────────────────────────────
// FETCH AUTH HELPER  (adds Authorization + X-Finnoto-Base)
// ─────────────────────────────────────────────────────────────

function fetchWithAuth(url, options = {}) {
    const headers = {
        'Authorization': `Bearer ${currentToken}`,
        'X-Finnoto-Base': currentFinnotoBase,
        ...(options.headers || {})
    };
    return fetchWithTimeout(url, { ...options, headers });
}

// ─────────────────────────────────────────────────────────────
// MAIN: runReportFetch
// ─────────────────────────────────────────────────────────────

function getSelectedSheetIds() {
    const cbs = document.querySelectorAll('#sheet-checkboxes input[type=checkbox]:checked');
    return Array.from(cbs).map(cb => parseInt(cb.value)).filter(id => !isNaN(id));
}

async function runReportFetch() {
    const select = document.getElementById('report-select');
    const reportId = parseInt(select.value);
    if (!reportId) {
        showError('Please select a report first.');
        return;
    }

    const sheetIds = getSelectedSheetIds();
    // Find the selected report in _allReports to get URL (if it came from search fallback)
    const selectedReport = _allReports.find(r => r.id === reportId);
    const reportUrl = selectedReport && selectedReport.url ? selectedReport.url : null;
    const hasSheets = sheetIds.length > 0;

    // If no sheets are available (search fallback report) and no URL, we can't do anything useful
    if (!hasSheets && !reportUrl) {
        showError('This report has no sheet definitions available. Try refreshing the report list or contacting your Finnoto admin.');
        return;
    }

    const datePayload    = buildDatePayload();
    const filtersPayload = buildFiltersPayload();

    setLoading(true);
    hideFeedback();
    isSyncCancelled = false;
    activeAbortController = new AbortController();

    const actionContainer = document.getElementById('action-buttons-container');
    const btnStopContainer = document.getElementById('btn-stop-container');
    const btnStop = document.getElementById('btn-stop');
    if (actionContainer)   actionContainer.classList.add('hidden');
    if (btnStopContainer)  btnStopContainer.classList.remove('hidden');
    if (btnStop) {
        btnStop.disabled = false;
        btnStop.querySelector('.btn-text').textContent = 'Stop Fetch';
        btnStop.onclick = () => {
            isSyncCancelled = true;
            if (activeAbortController) activeAbortController.abort();
            btnStop.querySelector('.btn-text').textContent = 'Stopping...';
            btnStop.disabled = true;
        };
    }

    updateProgress('Connecting to Finnoto...', 0);

    let totalRows = 0;
    let totalCols = 0;
    let sheetsCompleted = 0;

    const sheetState = {};

    try {
        await Excel.run(async (ctx) => {
            ctx.runtime.enableEvents = false;
            ctx.workbook.application.calculationMode = Excel.CalculationMode.manual;
            await ctx.sync();
        }).catch(() => {});

        // Build the WS payload — sheet IDs path (primary) or URL fallback path
        const wsPayload = hasSheets
            ? {
                business_report_id: reportId,
                sheet_ids: sheetIds,
                date: datePayload,
                filters: filtersPayload,
                base_url: _allReportsBase,
              }
            : {
                // No sheets — use generate/download fallback
                business_report_id: reportId,
                sheet_ids: [],
                url: reportUrl,
                date: datePayload,
                filters: filtersPayload,
                _fallback: true,
                base_url: _allReportsBase,
              };

        await sendWsRequest('report_sheet_stream', wsPayload, async (chunk, status) => {
            if (isSyncCancelled) return;

            if (status === "progress" && chunk.message) {
                updateProgress(chunk.message, null);
                return;
            }

            const sheetName = chunk._sheet_name || 'Sheet';

            if (status === "sheet_start" && chunk._columns) {
                const cols = chunk._columns;
                sheetState[sheetName] = { columns: cols, startRow: 1 };
                totalCols = Math.max(totalCols, cols.length);

                await ensureTargetSheet(sheetName);
                await Excel.run(async (ctx) => {
                    const ws = ctx.workbook.worksheets.getItem(sheetName);
                    ws.getUsedRange().clear();
                    const headerRange = ws.getRangeByIndexes(0, 0, 1, cols.length);
                    headerRange.values = [cols];
                    headerRange.format.font.bold = true;
                    headerRange.format.fill.color = '#FFD300';
                    headerRange.format.font.color = 'black';
                    await ctx.sync();
                });
                updateProgress(`Writing sheet: ${sheetName}`, 50);
                return;
            }

            if (status === "chunk" && chunk._rows && sheetState[sheetName]) {
                const state = sheetState[sheetName];
                const rows = chunk._rows;
                if (!rows.length) return;

                await Excel.run(async (ctx) => {
                    const ws = ctx.workbook.worksheets.getItem(sheetName);
                    const range = ws.getRangeByIndexes(state.startRow, 0, rows.length, state.columns.length);
                    range.values = rows;
                    await ctx.sync();
                });
                state.startRow += rows.length;
                totalRows += rows.length;
                updateProgress(`${sheetName}: ${state.startRow - 1} rows`, null);
            }

            if (status === "sheet_complete") {
                sheetsCompleted++;
                try {
                    await Excel.run(async (ctx) => {
                        const ws = ctx.workbook.worksheets.getItem(sheetName);
                        ws.getUsedRange().format.autofitColumns();
                        await ctx.sync();
                    });
                } catch (_) {}
            }
        });

        if (totalRows > 0) {
            showSuccess(
                isSyncCancelled
                    ? `Fetch stopped. Loaded ${totalRows} rows across ${sheetsCompleted} sheet(s).`
                    : `Done! ${totalRows} rows across ${sheetsCompleted} sheet(s) from Finnoto.`,
                totalRows, sheetsCompleted, totalCols
            );
        } else {
            showSuccess('No data returned from report sheets.', 0, 0, 0);
        }

    } catch (err) {
        if (err.name === 'AbortError' || isSyncCancelled) {
            showSuccess(`Fetch stopped. ${totalRows} rows loaded.`, totalRows, sheetsCompleted, totalCols);
        } else {
            showError('Report fetch failed: ' + (err.message || String(err)));
        }
    } finally {
        await Excel.run(async (ctx) => {
            ctx.runtime.enableEvents = true;
            ctx.workbook.application.calculationMode = Excel.CalculationMode.automatic;
            await ctx.sync();
        }).catch(() => {});
        setLoading(false);
        hideProgress();
        if (actionContainer)  actionContainer.classList.remove('hidden');
        if (btnStopContainer) btnStopContainer.classList.add('hidden');
    }
}

// Called when the user switches collection — re-read schema from sheet
function onCollectionChange() {
    // Schema bar stays showing whatever is in the sheet.
    // User must click Import Schema or run a Fetch to update it.
    refreshSchemaBar();
}

// Real-time JSON validation
function validateQuerySyntax() {
    if (!queryInput) return true; // legacy MongoDB UI not present — always valid
    const val = queryInput.value.trim();
    if (!val) {
        if (jsonValidity) { jsonValidity.className = "validity-indicator empty"; jsonValidity.textContent = "Empty (Fetch All)"; }
        if (btnRun) btnRun.disabled = !isApiOnline;
        if (document.getElementById('btn-delete')) document.getElementById('btn-delete').disabled = !isApiOnline;
        return true;
    }
    
    try {
        const parsed = JSON.parse(val);
        if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
            throw new Error("Must be a JSON object");
        }
        if (jsonValidity) { jsonValidity.className = "validity-indicator valid"; jsonValidity.textContent = "Valid Filter JSON"; }
        if (btnRun) btnRun.disabled = !isApiOnline;
        if (document.getElementById('btn-delete')) document.getElementById('btn-delete').disabled = !isApiOnline;
        return true;
    } catch (err) {
        if (jsonValidity) { jsonValidity.className = "validity-indicator invalid"; jsonValidity.textContent = "Invalid JSON"; }
        if (btnRun) btnRun.disabled = true;
        if (document.getElementById('btn-delete')) document.getElementById('btn-delete').disabled = true;
        return false;
    }
}

// Core Workflow: Sync Excel adjustments then Fetch matching data
async function runSyncAndFetch() {
    let queriesToFetch = [];
    let isMultiMode = false;
    
    if (document.getElementById('tab-builder').classList.contains('active')) {
        const built = buildGuiQuery();
        if (Array.isArray(built)) {
            isMultiMode = true;
            queriesToFetch = built;
        } else {
            if (!validateQuerySyntax()) {
                showError("Query filter is not valid JSON.");
                return;
            }
            queriesToFetch = [{ collection: collectionSelect.value, filters: built, limit: parseInt(limitInput.value) || 0 }];
        }
    } else {
        if (!validateQuerySyntax()) {
            showError("Query filter is not valid JSON.");
            return;
        }
        queriesToFetch = [{ collection: collectionSelect.value, filters: JSON.parse(queryInput.value.trim()), limit: parseInt(limitInput.value) || 0 }];
    }

    // Ensure all target sheets exist
    for (const query of queriesToFetch) {
        if (!query.collection) {
            showError("One of your queries is missing a collection.");
            return;
        }
        await ensureTargetSheet(query.collection);
    }

    setLoading(true);
    hideFeedback();
    
    isSyncCancelled = false;
    activeAbortController = new AbortController();

    const actionContainer = document.getElementById("action-buttons-container");
    const btnStopContainer = document.getElementById("btn-stop-container");
    const btnStop = document.getElementById("btn-stop");
    
    if (actionContainer) actionContainer.classList.add("hidden");
    if (btnStopContainer) btnStopContainer.classList.remove("hidden");
    if (btnStop) {
        btnStop.disabled = false;
        btnStop.querySelector(".btn-text").textContent = "Stop Fetch";
        btnStop.onclick = () => {
            isSyncCancelled = true;
            if (activeAbortController) activeAbortController.abort();
            btnStop.querySelector(".btn-text").textContent = "Stopping...";
            btnStop.disabled = true;
        };
    }
    
    let expectedTotal = 1000000;
    let inserted = 0;
    let updated = 0;
    let totalFetched = 0;

    if (!isMultiMode && queriesToFetch.length > 0) {
        try {
            const countRes = await fetchWithTimeout(`${API_BASE}/schema?collection=${encodeURIComponent(queriesToFetch[0].collection)}`);
            if (countRes.ok) {
                const countData = await countRes.json();
                expectedTotal = countData.total_count || 1000000;
            }
        } catch (e) {}
    }

    // Finnoto is read-only — skip write-back phase, go straight to fetch
    updateProgress("Connecting to Finnoto...", 0);

    try {
        // Suspend auto-calculation and events for extreme performance boost
        await Excel.run(async (context) => {
            context.runtime.enableEvents = false;
            context.workbook.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
        }).catch(() => {});

        updateProgress("Connecting to fetch stream...", 0);
        let fetchAction = isMultiMode ? `multi_stream_fetch` : `stream_fetch`;
        let fetchBody = isMultiMode ? { queries: queriesToFetch } : queriesToFetch[0];

        const BATCH_SIZE = 5000;
        let collState = {};
        if (!isMultiMode && queriesToFetch.length > 0) {
            collState[queriesToFetch[0].collection] = { currentBatch: [], startRow: 1, headersKnown: false, headers: ["_id"] };
        }

        // Wait for the stream fetch to complete while handling chunks
        await sendWsRequest(fetchAction, fetchBody, async (doc) => {
            if (isSyncCancelled) return;
            
            const docCol = doc._collection || (queriesToFetch[0] ? queriesToFetch[0].collection : "");
            if (!collState[docCol]) {
                collState[docCol] = { currentBatch: [], startRow: 1, headersKnown: false, headers: ["_id"] };
            }
            
            if (doc._collection) {
                delete doc._collection;
            }
            
            collState[docCol].currentBatch.push(doc);
            totalFetched++;
            
            // Process batch if large enough
            let state = collState[docCol];
            if (state.currentBatch.length >= BATCH_SIZE) {
                const chunkToProcess = state.currentBatch.slice(0, BATCH_SIZE);
                state.currentBatch = state.currentBatch.slice(BATCH_SIZE);
                
                if (!state.headersKnown) {
                    const hs = new Set(["_id"]);
                    chunkToProcess.forEach(r => Object.keys(r).forEach(k => hs.add(k)));
                    state.headers = Array.from(hs);
                    state.headersKnown = true;
                    if (isMultiMode) await ensureTargetSheet(docCol);
                }
                
                updateProgress(`Rendering ${docCol}...`, null);
                state.startRow = await appendDataToSheet(docCol, chunkToProcess, state.headers, state.startRow);
            }
        });

        for (const colName in collState) {
            let state = collState[colName];
            if (state.currentBatch.length > 0) {
                if (!state.headersKnown) {
                    const hs = new Set(["_id"]);
                    state.currentBatch.forEach(r => Object.keys(r).forEach(k => hs.add(k)));
                    state.headers = Array.from(hs);
                    if (isMultiMode) await ensureTargetSheet(colName);
                }
                updateProgress(`Rendering final ${colName} records...`, null);
                state.startRow = await appendDataToSheet(colName, state.currentBatch, state.headers, state.startRow);
            }
            
            if (state.startRow === 1 && !isMultiMode) {
                await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    try { sheet.getUsedRange().clear(); await context.sync(); } catch(e){}
                });
            }
        }

        await refreshSchemaBar();

        showSuccess(
            isSyncCancelled
                ? `Fetch stopped by user. Loaded ${totalFetched} records.`
                : (totalFetched > 0 
                    ? `Fetch complete! Loaded ${totalFetched} records from Finnoto.`
                    : `Fetch complete. No records matched the query.`),
            totalFetched, 0, 0
        );
    } catch (error) {
        if (error.name === "AbortError" || isSyncCancelled) {
            showSuccess(`Sync stopped by user. Loaded ${totalFetched} records.`, totalFetched, inserted, updated);
        } else {
            showError("Execution failed: " + error.message);
        }
    } finally {
        if (actionContainer) actionContainer.classList.remove("hidden");
        const stopContainer = document.getElementById("btn-stop-container");
        if (stopContainer) stopContainer.classList.add("hidden");
        await Excel.run(async (context) => {
            context.runtime.enableEvents = true;
            context.workbook.application.calculationMode = Excel.CalculationMode.automatic;
            await context.sync();
        }).catch(() => {});
        setLoading(false);
        hideProgress();
    }
}


// Read sheet values using Office.js
async function getSheetData() {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowCount", "columnCount"]);
        
        try {
            await context.sync();
            return {
                values: usedRange.values,
                rowCount: usedRange.rowCount,
                columnCount: usedRange.columnCount
            };
        } catch (error) {
            // Sheet is empty or lacks formatted range
            return null;
        }
    });
}

// Convert 2D Excel sheet data to inserts & updates
function parseSheetData(sheetData) {
    const payload = { inserts: [], updates: [] };
    if (!sheetData || !sheetData.values || sheetData.values.length === 0) {
        return payload;
    }

    const rows = sheetData.values;
    const headers = rows[0];
    const dataRows = rows.slice(1);

    const idIndex = headers.indexOf("_id");

    dataRows.forEach((row, rowIndex) => {
        // Skip empty rows
        if (row.every(cell => cell === "" || cell === null || cell === undefined)) {
            return;
        }

        const doc = {};
        let hasData = false;

        for (let i = 0; i < headers.length; i++) {
            const key = headers[i];
            if (!key) continue; // Skip column without header title
            
            const cellVal = row[i];
            
            if (key === "_id") {
                if (cellVal !== "" && cellVal !== null && cellVal !== undefined) {
                    doc["_id"] = String(cellVal).trim();
                }
            } else {
                if (cellVal !== "" && cellVal !== null && cellVal !== undefined) {
                    doc[key] = cellVal;
                    hasData = true;
                }
            }
        }

        // Only sync if there is some useful cell data
        if (hasData || doc["_id"]) {
            doc._rowIndex = rowIndex + 1;
            if (doc["_id"]) {
                payload.updates.push(doc);
            } else {
                payload.inserts.push(doc);
            }
        }
    });

    return payload;
}

// Overwrite the current active worksheet incrementally
async function appendDataToSheet(baseSheetName, records, headers, globalStartRow) {
    if (records.length === 0) return globalStartRow;

    const EXCEL_LIMIT = 1000000;
    
    const sheetIndex = Math.floor((globalStartRow - 1) / EXCEL_LIMIT) + 1;
    let localStartRow = ((globalStartRow - 1) % EXCEL_LIMIT) + 1;

    await Excel.run(async (context) => {
        let sheet;
        
        if (sheetIndex === 1) {
            sheet = context.workbook.worksheets.getActiveWorksheet();
        } else {
            const sheetName = `${baseSheetName}_${sheetIndex}`;
            sheet = context.workbook.worksheets.getItemOrNullObject(sheetName);
            await context.sync();
            
            if (sheet.isNullObject) {
                sheet = context.workbook.worksheets.add(sheetName);
                sheet.activate();
                await context.sync();
            }
        }

        if (localStartRow === 1) {
            try {
                sheet.getUsedRange().clear();
                await context.sync();
            } catch(e) {}
            
            const hr = sheet.getRangeByIndexes(0, 0, 1, headers.length);
            hr.values = [headers];
            hr.format.fill.color = "#5BAD7F";
            hr.format.font.color = "#FFFFFF";
            hr.format.font.bold = true;
            
            const vIndex = headers.indexOf("__v");
            if (vIndex !== -1) {
                sheet.getRangeByIndexes(0, vIndex, 1, 1).columnHidden = true;
            }
            await context.sync();
        }

        const values = [];
        records.forEach(rec => {
            const row = [];
            headers.forEach(h => {
                let cell = rec[h];
                if (cell !== null && typeof cell === "object") cell = JSON.stringify(cell);
                else if (cell === undefined) cell = "";
                row.push(cell);
            });
            values.push(row);
        });

        const range = sheet.getRangeByIndexes(localStartRow, 0, values.length, headers.length);
        range.values = values;
        
        const vIndex = headers.indexOf("__v");
        if (vIndex !== -1) {
            sheet.getRangeByIndexes(localStartRow, vIndex, values.length, 1).columnHidden = true;
        }

        if (localStartRow === 1) {
            range.format.autofitColumns();
        }

        context.workbook.application.suspendScreenUpdatingUntilNextSync();
        await context.sync();
    });

    return globalStartRow + records.length;
}

// --- Schema Functions ---

/**
 * Reads row 1 of the active sheet and returns the field headers found.
 * Returns [] if the sheet is empty or has no headers.
 */
async function detectSheetSchema() {
    return await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const usedRange = sheet.getUsedRange();
        usedRange.load(["values", "rowCount"]);
        try {
            await context.sync();
            if (!usedRange.values || usedRange.values.length === 0) return [];
            // Row 1 is the header row
            const headers = usedRange.values[0].filter(h => h !== null && h !== "");
            return headers.map(String).filter(h => h !== "__v");
        } catch {
            return [];
        }
    });
}

/**
 * Reads the sheet schema and renders the schema bar.
 * Also refreshes any condition field selectors.
 */
async function refreshSchemaBar() {
    if (!schemaBar || !schemaBarFields) return;
    const fields = await detectSheetSchema();
    currentSchema = fields;
    
    let countText = "";
    const collection = collectionSelect ? collectionSelect.value : "";
    if (collection && isApiOnline) {
        try {
            const res = await fetchWithTimeout(`${API_BASE}/schema?collection=${encodeURIComponent(collection)}`);
            if (res.ok) {
                const data = await res.json();
                if (data.total_count !== undefined) {
                    countText = `${data.total_count.toLocaleString()} total records in DB`;
                }
            }
        } catch (e) {}
    }

    renderSchemaBar(fields, countText);
    refreshConditionFields();
}

/**
 * Renders the schema field tags inside the schema bar.
 */
function renderSchemaBar(fields, countText = "") {
    if (!fields || fields.length === 0) {
        schemaBar.classList.add("hidden");
        return;
    }
    schemaBarFields.innerHTML = "";
    fields.forEach(f => {
        const tag = document.createElement("span");
        tag.className = "schema-field-tag" + (f === "_id" ? " id-field" : "");
        tag.textContent = f;
        schemaBarFields.appendChild(tag);
    });

    let countEl = document.getElementById('schema-total-records');
    if (countEl) {
        countEl.textContent = countText;
    }

    schemaBar.classList.remove("hidden");
}

/**
 * Fetches the schema from MongoDB and writes headers to row 1 of the sheet.
 * Blocked if no collection is selected.
 */
async function importSchema() {
    const collection = collectionSelect.value;
    if (!collection) {
        showError("Select a collection before importing schema.");
        return;
    }

    await ensureTargetSheet(collection);

    btnImportSchema.disabled = true;
    btnImportSchema.textContent = "Loading...";
    hideFeedback();

    try {
        const res = await fetchWithTimeout(`${API_BASE}/schema?collection=${encodeURIComponent(collection)}`);
        const data = await res.json();

        if (!res.ok) throw new Error(data.detail || "Schema fetch failed");
        if (!data.fields || data.fields.length === 0) {
            showError(data.message || "Collection is empty — no schema could be inferred.");
            return;
        }

        // Write headers to the sheet
        await writeSchemaToSheet(data.fields);
        await refreshSchemaBar();

        showSuccess(
            `Schema imported: ${data.fields.length} fields from ${data.sampled} sample document(s). ` +
            `Add new rows below the headers and click Sync & Fetch Data to insert them.`,
            0, 0, 0
        );
    } catch (err) {
        showError("Schema import failed: " + err.message);
    } finally {
        btnImportSchema.disabled = false;
        btnImportSchema.innerHTML = "⬇ Schema";
    }
}

/**
 * Writes field names as styled headers to row 1 without touching data rows.
 */
async function writeSchemaToSheet(fields) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

        // Only clear row 1 header (don't wipe existing data rows)
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, fields.length);
        headerRange.values = [fields];
        headerRange.format.fill.color = "#13AA52";
        headerRange.format.font.color = "#FFFFFF";
        headerRange.format.font.bold = true;
        headerRange.format.autofitColumns();

        await context.sync();
    });
}

// ============================================================
// GUI QUERY BUILDER
// ============================================================



/** Switch between 'builder' and 'json' tabs */
function switchQueryMode(mode) {
    const builderPanel = document.getElementById('builder-panel');
    const jsonPanel    = document.getElementById('json-panel');
    const tabBuilder   = document.getElementById('tab-builder');
    const tabJson      = document.getElementById('tab-json');

    if (mode === 'builder') {
        builderPanel.classList.remove('hidden');
        jsonPanel.classList.add('hidden');
        tabBuilder.classList.add('active');
        tabJson.classList.remove('active');
        updateQueryFromGui(); // keep JSON in sync
    } else {
        builderPanel.classList.add('hidden');
        jsonPanel.classList.remove('hidden');
        tabBuilder.classList.remove('active');
        tabJson.classList.add('active');
    }
}

// ============================================================================
// GUI QUERY BUILDER (MULTI-BLOCK)
// ============================================================================

let queryBlocks = [
    {
        id: 0,
        isPrimary: true,
        collection: "",
        schema: [],
        logic: 'and',
        conditions: [],
        conditionCounter: 0
    }
];
let nextBlockId = 1;

function getBlock(id) {
    return queryBlocks.find(b => b.id === parseInt(id));
}

function renderBlocks() {
    const container = document.getElementById('blocks-container');
    if (!container) return;
    container.innerHTML = '';
    
    queryBlocks.forEach((block, index) => {
        const blockEl = document.createElement('div');
        blockEl.className = 'border-border-width border-black p-3 bg-surface-container-low neo-shadow-sm space-y-3 relative group';
        blockEl.dataset.blockId = block.id;
        
        let headerHtml = '';
        if (block.isPrimary) {
            headerHtml = `
                <div class="flex justify-between items-center">
                    <span class="font-label-mono text-[11px] uppercase">Match (Primary Collection)</span>
                    <div class="flex border-2 border-black overflow-hidden">
                        <button class="btn-logic btn-logic-and px-3 py-1 text-[10px] font-bold ${block.logic==='and'?'bg-white text-black':'bg-surface-variant text-on-surface-variant'} border-r-2 border-black" data-block="${block.id}" data-logic="and">ALL (AND)</button>
                        <button class="btn-logic btn-logic-or px-3 py-1 text-[10px] font-bold ${block.logic==='or'?'bg-white text-black':'bg-surface-variant text-on-surface-variant'}" data-block="${block.id}" data-logic="or">ANY (OR)</button>
                    </div>
                </div>
            `;
        } else {
            const options = Array.from(collectionSelect.options).map(opt => `<option value="${opt.value}" ${opt.value === block.collection ? 'selected' : ''}>${opt.text}</option>`).join('');
            headerHtml = `
                <div class="flex justify-between items-center mb-2">
                    <select class="block-collection-select bg-surface-container-lowest border-2 border-black p-1 text-[11px] font-bold font-label-mono uppercase outline-none" data-block="${block.id}">
                        ${options}
                    </select>
                    <button class="btn-remove-block text-error hover:scale-110 transition-transform" data-block="${block.id}" title="Remove Collection Query">
                        <span class="material-symbols-outlined text-[18px]">delete</span>
                    </button>
                </div>
                <div class="flex justify-between items-center">
                    <span class="font-label-mono text-[11px] uppercase">Match</span>
                    <div class="flex border-2 border-black overflow-hidden">
                        <button class="btn-logic btn-logic-and px-3 py-1 text-[10px] font-bold ${block.logic==='and'?'bg-white text-black':'bg-surface-variant text-on-surface-variant'} border-r-2 border-black" data-block="${block.id}" data-logic="and">ALL</button>
                        <button class="btn-logic btn-logic-or px-3 py-1 text-[10px] font-bold ${block.logic==='or'?'bg-white text-black':'bg-surface-variant text-on-surface-variant'}" data-block="${block.id}" data-logic="or">ANY</button>
                    </div>
                </div>
            `;
        }
        
        let conditionsHtml = '<div class="conditions-list space-y-2 mt-2">';
        block.conditions.forEach(cond => {
            const schemaFields = block.schema.filter(f => f !== '_id');
            let fieldInput = '';
            if (schemaFields.length > 0) {
                const fOpts = schemaFields.map(f => `<option value="${f}" ${f === cond.field ? 'selected' : ''}>${f}</option>`).join('');
                fieldInput = `<select class="cond-field bg-surface-container-lowest border-2 border-black p-1 font-data-field w-full outline-none" data-block="${block.id}" data-cond="${cond.id}">${fOpts}</select>`;
            } else {
                fieldInput = `<input type="text" class="cond-field bg-surface-container-lowest border-2 border-black p-1 font-data-field w-full outline-none" data-block="${block.id}" data-cond="${cond.id}" value="${cond.field}" placeholder="field">`;
            }
            
            const opOpts = GUI_OPERATORS.map(o => `<option value="${o.value}" ${o.value === cond.op ? 'selected' : ''}>${o.label}</option>`).join('');
            const opInput = `<select class="cond-op bg-surface-container-lowest border-2 border-black p-1 font-label-mono text-[10px] w-full outline-none" data-block="${block.id}" data-cond="${cond.id}">${opOpts}</select>`;
            
            const selectedOpDef = GUI_OPERATORS.find(o => o.value === cond.op);
            const valInput = selectedOpDef && selectedOpDef.hasValue 
                ? `<input type="text" class="cond-val bg-surface-container-lowest border-2 border-black p-1 font-data-field w-full outline-none" data-block="${block.id}" data-cond="${cond.id}" value="${cond.value}" placeholder="value">`
                : `<div class="w-full text-[10px] text-on-surface-variant italic py-1 px-2">no value</div>`;
                
            conditionsHtml += `
                <div class="flex gap-2 items-center bg-background p-2 border-2 border-black">
                    <div class="flex-grow min-w-[30%]">${fieldInput}</div>
                    <div class="flex-grow min-w-[30%]">${opInput}</div>
                    <div class="flex-grow min-w-[30%]">${valInput}</div>
                    <button class="btn-remove-cond text-error hover:scale-110 transition-transform" data-block="${block.id}" data-cond="${cond.id}">
                        <span class="material-symbols-outlined text-[16px]">close</span>
                    </button>
                </div>
            `;
        });
        conditionsHtml += '</div>';
        
        const addBtnHtml = `
            <button class="btn-add-cond w-full bg-white border-2 border-black p-2 neo-shadow-sm flex items-center justify-center gap-2 hover:-translate-y-1 transition-transform mt-3 outline-none" data-block="${block.id}">
                <span class="material-symbols-outlined text-primary-container text-[16px]">add_circle</span>
                <span class="font-label-mono text-[10px] uppercase font-bold">Add Condition</span>
            </button>
        `;
        
        blockEl.innerHTML = headerHtml + conditionsHtml + addBtnHtml;
        container.appendChild(blockEl);
    });

    updateQueryFromGui();
}

function addQueryBlock() {
    // Determine a default collection (either current or first available)
    let defaultCol = collectionSelect.value;
    if (collectionSelect.options.length > 0 && !defaultCol) {
        defaultCol = collectionSelect.options[0].value;
    }

    const newBlock = {
        id: nextBlockId++,
        isPrimary: false,
        collection: defaultCol,
        schema: [],
        logic: 'and',
        conditions: [],
        conditionCounter: 0
    };
    queryBlocks.push(newBlock);
    renderBlocks();
    
    // Auto-fetch schema for the new block
    if (defaultCol) {
        handleBlockCollectionChange(newBlock.id, defaultCol);
    }
}

async function handleBlockCollectionChange(blockId, newCol) {
    const block = getBlock(blockId);
    if (!block) return;
    block.collection = newCol;
    block.schema = []; // reset
    renderBlocks();
    
    if (!newCol) return;
    
    try {
        const res = await fetchWithTimeout(`${API_BASE}/schema?collection=${encodeURIComponent(newCol)}`);
        const data = await res.json();
        if (res.ok && data.fields) {
            block.schema = data.fields;
            block.conditions.forEach(cond => {
                if (data.fields.length > 0 && !data.fields.includes(cond.field)) {
                    cond.field = data.fields.find(f => f !== '_id') || '_id';
                }
            });
            renderBlocks();
        }
    } catch (e) {
        console.error("Auto-fetch schema failed for block", e);
    }
}

function addConditionToBlock(blockId) {
    const block = getBlock(blockId);
    if (!block) return;
    
    let defaultField = '';
    const schemaFields = block.schema.filter(f => f !== '_id');
    if (schemaFields.length > 0) defaultField = schemaFields[0];

    block.conditions.push({
        id: ++block.conditionCounter,
        field: defaultField,
        op: 'eq',
        value: ''
    });
    
    renderBlocks();
}

function bindBlocksContainerEvents() {
    const container = document.getElementById('blocks-container');
    if (!container) return;
    
    container.addEventListener('click', (e) => {
        const btnLogic = e.target.closest('.btn-logic');
        if (btnLogic) {
            const block = getBlock(btnLogic.dataset.block);
            if (block) {
                block.logic = btnLogic.dataset.logic;
                renderBlocks();
            }
            return;
        }
        
        const btnAddCond = e.target.closest('.btn-add-cond');
        if (btnAddCond) {
            addConditionToBlock(btnAddCond.dataset.block);
            return;
        }
        
        const btnRemoveCond = e.target.closest('.btn-remove-cond');
        if (btnRemoveCond) {
            const block = getBlock(btnRemoveCond.dataset.block);
            if (block) {
                block.conditions = block.conditions.filter(c => c.id !== parseInt(btnRemoveCond.dataset.cond));
                renderBlocks();
            }
            return;
        }
        
        const btnRemoveBlock = e.target.closest('.btn-remove-block');
        if (btnRemoveBlock) {
            queryBlocks = queryBlocks.filter(b => b.id !== parseInt(btnRemoveBlock.dataset.block));
            renderBlocks();
            return;
        }
    });

    container.addEventListener('change', (e) => {
        const colSelect = e.target.closest('.block-collection-select');
        if (colSelect) {
            handleBlockCollectionChange(colSelect.dataset.block, colSelect.value);
            return;
        }
        
        const condField = e.target.closest('.cond-field');
        const condOp = e.target.closest('.cond-op');
        const condVal = e.target.closest('.cond-val');
        
        if (condField) {
            const block = getBlock(condField.dataset.block);
            const cond = block.conditions.find(c => c.id === parseInt(condField.dataset.cond));
            cond.field = condField.value;
            updateQueryFromGui();
        } else if (condOp) {
            const block = getBlock(condOp.dataset.block);
            const cond = block.conditions.find(c => c.id === parseInt(condOp.dataset.cond));
            cond.op = condOp.value;
            renderBlocks();
        } else if (condVal) {
            const block = getBlock(condVal.dataset.block);
            const cond = block.conditions.find(c => c.id === parseInt(condVal.dataset.cond));
            cond.value = condVal.value;
            updateQueryFromGui();
        }
    });

    container.addEventListener('input', (e) => {
        const condField = e.target.closest('.cond-field');
        const condVal = e.target.closest('.cond-val');
        
        if (condField && condField.tagName === 'INPUT') {
            const block = getBlock(condField.dataset.block);
            const cond = block.conditions.find(c => c.id === parseInt(condField.dataset.cond));
            cond.field = condField.value;
            updateQueryFromGui();
        } else if (condVal && condVal.tagName === 'INPUT') {
            const block = getBlock(condVal.dataset.block);
            const cond = block.conditions.find(c => c.id === parseInt(condVal.dataset.cond));
            cond.value = condVal.value;
            updateQueryFromGui();
        }
    });
}

/** Rebuild GUI filter to a MongoDB query object or array of objects */
function buildGuiQuery() {
    const parseVal = v => {
        if (v === 'true')  return true;
        if (v === 'false') return false;
        const n = Number(v);
        return (!isNaN(n) && v !== '') ? n : v;
    };

    let queries = [];
    queryBlocks.forEach(block => {
        const conditions = [];
        block.conditions.forEach(c => {
            const field = (c.field || '').trim();
            const op    = c.op;
            const raw   = (c.value || '').trim();
            if (!field) return;

            let cond = {};
            switch (op) {
                case 'eq':      cond = { [field]: parseVal(raw) }; break;
                case 'ne':      cond = { [field]: { $ne:  parseVal(raw) } }; break;
                case 'gt':      cond = { [field]: { $gt:  parseVal(raw) } }; break;
                case 'gte':     cond = { [field]: { $gte: parseVal(raw) } }; break;
                case 'lt':      cond = { [field]: { $lt:  parseVal(raw) } }; break;
                case 'lte':     cond = { [field]: { $lte: parseVal(raw) } }; break;
                case 'regex':   cond = { [field]: { $regex: raw, $options: 'i' } }; break;
                case 'starts':  cond = { [field]: { $regex: `^${raw}`, $options: 'i' } }; break;
                case 'in':      cond = { [field]: { $in:  raw.split(',').map(v => parseVal(v.trim())) } }; break;
                case 'nin':     cond = { [field]: { $nin: raw.split(',').map(v => parseVal(v.trim())) } }; break;
                case 'exists':  cond = { [field]: { $exists: true  } }; break;
                case 'nexists': cond = { [field]: { $exists: false } }; break;
            }
            conditions.push(cond);
        });

        let query = {};
        if (conditions.length === 1) {
            query = conditions[0];
        } else if (conditions.length > 1) {
            query = block.logic === 'and' ? { $and: conditions } : { $or: conditions };
        }
        
        const col = block.isPrimary ? (collectionSelect.value || "") : block.collection;
        queries.push({
            collection: col,
            filters: query
        });
    });

    if (queries.length === 1) return queries[0].filters;
    return queries;
}

/** Sync GUI -> JSON preview and JSON inputs */
function updateQueryFromGui() {
    const output = buildGuiQuery();
    
    if (queryBlocks.length === 1) {
        // Single block mode (backwards compatible)
        const isEmpty = Object.keys(output).length === 0;
        const pretty  = isEmpty ? '' : JSON.stringify(output, null, 2);
        queryInput.value = pretty;
        document.getElementById('query-preview-code').textContent = isEmpty ? '{}  // fetch all' : pretty;
        validateQuerySyntax();
    } else {
        // Multi-block mode
        const pretty = JSON.stringify(output, null, 2);
        document.getElementById('multi-input').value = pretty;
        document.getElementById('query-preview-code').textContent = pretty;
        validateMultiSyntax();
    }
    saveSheetState();
}

/** When primary schema updates, rebuild primary block */
function refreshConditionFields() {
    if (queryBlocks[0]) {
        queryBlocks[0].schema = [...currentSchema];
        queryBlocks[0].collection = collectionSelect.value;
        queryBlocks[0].conditions.forEach(cond => {
            if (currentSchema.length > 0 && !currentSchema.includes(cond.field)) {
                cond.field = currentSchema.find(f => f !== '_id') || '_id';
            }
        });
        renderBlocks();
    }
}

// --- Helpers ---
function updateProgress(msg, percentage = null) {
    const el = document.getElementById('progress-text');
    el.textContent = msg;
    el.classList.remove('hidden');

    const barContainer = document.getElementById('progress-bar-container');
    const barFill = document.getElementById('progress-bar-fill');
    
    if (percentage !== null) {
        barContainer.classList.remove('hidden');
        barFill.style.width = `${Math.min(100, Math.max(0, percentage))}%`;
    } else {
        barContainer.classList.add('hidden');
    }
}

function hideProgress() {
    document.getElementById('progress-text').classList.add('hidden');
    const barContainer = document.getElementById('progress-bar-container');
    if (barContainer) barContainer.classList.add('hidden');
}

function updateStatus(className, text) {
    connectionStatus.className = `status-badge ${className}`;
    connectionStatus.querySelector(".status-text").textContent = text;
}

function setLoading(isLoading) {
    if (btnRun) btnRun.disabled = isLoading;
    if (collectionSelect) collectionSelect.disabled = isLoading;
    if (btnRefreshCollections) btnRefreshCollections.disabled = isLoading;
    if (queryInput) queryInput.disabled = isLoading;

    if (isLoading) {
        if (btnText) btnText.textContent = "Processing...";
        if (btnSpinner) btnSpinner.classList.remove("hidden");
    } else {
        if (btnText) btnText.textContent = "Sync & Fetch Data";
        if (btnSpinner) btnSpinner.classList.add("hidden");
    }
}

function showSuccess(msg, fetched = 0, inserted = 0, updated = 0) {
    statFetched.textContent = fetched;
    statInserted.textContent = inserted;
    statUpdated.textContent = updated;
    
    const el = document.getElementById('feedback-message');
    el.textContent = msg;
    el.className = 'feedback-message success';
    document.getElementById('feedback-panel').classList.remove('hidden');
}

// Login Handler
document.getElementById('btn-login').addEventListener('click', async () => {
    const pastedToken = document.getElementById('mongo-uri-input').value.trim();
    const errEl = document.getElementById('login-error');
    const loader = document.getElementById('login-loader');

    if (!pastedToken) {
        errEl.textContent = 'Please paste your Finnoto Bearer Token.';
        errEl.classList.remove('hidden');
        return;
    }

    // Finnoto instance URL — use selected env preset or custom
    const useUserApiBase = document.getElementById('use-custom-backend-checkbox').checked;
    const userApiBase = useUserApiBase ? document.getElementById('api-base-input').value.trim() : '';

    if (useUserApiBase && userApiBase) {
        currentFinnotoBase = userApiBase.replace(/\/+$/, '');
        _selectedEnvKey = '';
    } else {
        currentFinnotoBase = FINNOTO_BASE_PRESETS[_selectedEnvKey] || 'https://eapi.finnoto.dev';
    }

    if (useUserApiBase && userApiBase) {
        API_BASE = userApiBase.replace(/\/+$/, '');
    } else {
        API_BASE = (window.EXCEL_MONGO_CONFIG && window.EXCEL_MONGO_CONFIG.apiBase) || window.location.origin;
    }

    errEl.classList.add('hidden');
    loader.classList.remove('hidden');
    document.getElementById('btn-login').disabled = true;

    try {
        const res = await fetch(`${API_BASE}/auth/connect`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ token: pastedToken, base_url: currentFinnotoBase })
        });
        const data = await res.json();
        if (!res.ok) throw new Error(data.detail || 'Connection failed.');

        currentToken = data.token;
        Office.context.document.settings.set('FinnotoToken', currentToken);
        Office.context.document.settings.set('FinnotoBase', currentFinnotoBase);
        if (_selectedEnvKey) {
            Office.context.document.settings.set('FinnotoEnv', _selectedEnvKey);
        }

        if (useUserApiBase && userApiBase) {
            Office.context.document.settings.set('ApiBase', API_BASE);
        } else {
            Office.context.document.settings.remove('ApiBase');
        }
        Office.context.document.settings.saveAsync();

        document.getElementById('login-view').classList.add('hidden');
        document.getElementById('app-view').classList.remove('hidden');
        await initApp();
    } catch (e) {
        errEl.textContent = e.message;
        errEl.classList.remove('hidden');
    } finally {
        loader.classList.add('hidden');
        document.getElementById('btn-login').disabled = false;
    }
});

function showError(msg) {
    statFetched.textContent = "0";
    statInserted.textContent = "0";
    statUpdated.textContent = "0";

    feedbackMessage.className = "feedback-message error";
    feedbackMessage.textContent = typeof msg === 'object' ? JSON.stringify(msg) : msg;
    feedbackPanel.classList.remove("hidden");
}

function hideFeedback() {
    feedbackPanel.classList.add("hidden");
}
