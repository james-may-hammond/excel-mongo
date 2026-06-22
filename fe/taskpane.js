// Local dev uses fe/config.js. Production builds generate this value from ADDIN_API_BASE.
let API_BASE = (window.EXCEL_MONGO_CONFIG && window.EXCEL_MONGO_CONFIG.apiBase) || window.location.origin;

let currentMongoUri = "";
let currentMongoDb = "";

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
        let wsUrl = `${wsProtocol}${wsHost}/ws?uri=${encodeURIComponent(currentMongoUri)}&db_name=${encodeURIComponent(currentMongoDb)}`;
        
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
            
            if (data.status === "chunk" && wsStreamHandlers.has(reqId)) {
                wsStreamHandlers.get(reqId)(data.data);
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
        "X-Mongo-URI": currentMongoUri,
        "X-Mongo-DB": currentMongoDb,
        ...existingHeaders
    };
}

// Fetch with a timeout so we fail fast instead of hanging
async function fetchWithTimeout(url, options = {}, timeoutMs = 60000) {
    if (url.endsWith('/health') || url.endsWith('/create_collection')) {
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
const btnText = btnRun.querySelector(".btn-text");
const btnSpinner = btnRun.querySelector(".btn-spinner");
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
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        // Automatically open the taskpane next time this document is opened
        Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
        Office.context.document.settings.saveAsync();

        // Restore credentials
        currentMongoUri = Office.context.document.settings.get("MongoUri") || "";
        currentMongoDb = Office.context.document.settings.get("MongoDb") || "";
        
        const savedApiBase = Office.context.document.settings.get("ApiBase");
        if (savedApiBase) {
            API_BASE = savedApiBase;
            document.getElementById("api-base-input").value = API_BASE;
            document.getElementById("use-custom-backend-checkbox").checked = true;
            document.getElementById("custom-backend-container").classList.remove("hidden");
        } else {
            document.getElementById("api-base-input").value = "";
        }
        
        document.getElementById('use-custom-backend-checkbox').addEventListener('change', (e) => {
            if (e.target.checked) {
                document.getElementById('custom-backend-container').classList.remove('hidden');
            } else {
                document.getElementById('custom-backend-container').classList.add('hidden');
            }
        });
        
        if (currentMongoUri && currentMongoDb) {
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
    // Collection controls
    btnRefreshCollections.addEventListener('click', loadCollections);
    btnImportSchema.addEventListener('click', importSchema);
    collectionSelect.addEventListener('change', async () => {
        const col = collectionSelect.value;
        if (col) {
            await ensureTargetSheet(col);
        }
        onCollectionChange();
        saveSheetState();
    });

    // Query builder controls
    document.getElementById('tab-builder').addEventListener('click', () => switchQueryMode('builder'));
    document.getElementById('tab-json').addEventListener('click',    () => switchQueryMode('json'));
    
    bindBlocksContainerEvents();
    renderBlocks();

    // JSON textarea
    queryInput.addEventListener('input', () => {
        validateQuerySyntax();
        saveSheetState();
    });

    // Main buttons
    btnRun.addEventListener('click', runSyncAndFetch);
    const btnDelete = document.getElementById('btn-delete');
    if (btnDelete) {
        btnDelete.addEventListener('click', runDeleteSelected);
    }

    // Boot sequence
    await checkApiHealth();
    if (isApiOnline) await loadCollections();
    
    // Register workbook activation event and load current state
    await Excel.run(async (context) => {
        context.workbook.worksheets.onActivated.add(onWorksheetActivated);
        await context.sync();
    }).catch(console.error);
    await loadSheetState();
    
    await refreshSchemaBar();
}

async function saveSheetState() {
    const col = collectionSelect.value || "";
    const query = queryInput.value || "";
    
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
        
        if (!colProp.isNullObject && colProp.value) {
            collectionSelect.value = colProp.value;
        }
        if (!queryProp.isNullObject && queryProp.value) {
            queryInput.value = queryProp.value;
            validateQuerySyntax();
        }
    }).catch(console.error);
}

async function onWorksheetActivated(event) {
    await loadSheetState();
    await refreshSchemaBar();
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

// Check if FastAPI is running and connected to MongoDB
async function checkApiHealth() {
    try {
        updateStatus("checking", "Checking API...");
        const res = await fetchWithTimeout(`${API_BASE}/health`);
        const data = await res.json();
        
        if (data.status === "ok") {
            try {
                await connectWebSocket();
                isApiOnline = true;
                updateStatus("online", "Connected to Mongo");
                btnRun.disabled = false;
                if (document.getElementById('btn-delete')) document.getElementById('btn-delete').disabled = false;
            } catch (wsErr) {
                throw new Error("WebSocket connection failed");
            }
        } else {
            throw new Error("Health check returned status " + data.status);
        }
    } catch (err) {
        isApiOnline = false;
        updateStatus("offline", "API Offline");
        btnRun.disabled = true;
        if (document.getElementById('btn-delete')) document.getElementById('btn-delete').disabled = true;
        const msg = err.name === "AbortError"
            ? "Connection timed out. Make sure ./run.sh is running in your terminal."
            : "Backend API is offline. Run ./run.sh in your terminal first.";
        showError(msg);
    }
}

// Load MongoDB Collections into selector dropdown
async function loadCollections() {
    if (!isApiOnline) return;
    
    collectionSelect.disabled = true;
    btnImportSchema.disabled = true;
    collectionSelect.innerHTML = '<option value="">Loading...</option>';
    
    try {
        const res = await fetchWithTimeout(`${API_BASE}/collections`);
        const data = await res.json();
        
        if (res.ok && data.collections) {
            collectionSelect.innerHTML = "";
            if (data.collections.length === 0) {
                collectionSelect.innerHTML = '<option value="">(No collections found)</option>';
            } else {
                data.collections.forEach(colName => {
                    const opt = document.createElement("option");
                    opt.value = colName;
                    opt.textContent = colName;
                    collectionSelect.appendChild(opt);
                });
                collectionSelect.disabled = false;
                btnImportSchema.disabled = false;
            }
        } else {
            throw new Error(data.detail || "Failed to load collections");
        }
    } catch (err) {
        showError("Failed to fetch MongoDB collections: " + err.message);
        collectionSelect.innerHTML = '<option value="">Error loading collections</option>';
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
    const val = queryInput.value.trim();
    if (!val) {
        jsonValidity.className = "validity-indicator empty";
        jsonValidity.textContent = "Empty (Fetch All)";
        btnRun.disabled = !isApiOnline;
        if (document.getElementById('btn-delete')) document.getElementById('btn-delete').disabled = !isApiOnline;
        return true;
    }
    
    try {
        const parsed = JSON.parse(val);
        if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) {
            throw new Error("Must be a JSON object");
        }
        jsonValidity.className = "validity-indicator valid";
        jsonValidity.textContent = "Valid Filter JSON";
        btnRun.disabled = !isApiOnline;
        if (document.getElementById('btn-delete')) document.getElementById('btn-delete').disabled = !isApiOnline;
        return true;
    } catch (err) {
        jsonValidity.className = "validity-indicator invalid";
        jsonValidity.textContent = "Invalid JSON";
        btnRun.disabled = true;
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
        btnStop.querySelector(".btn-text").textContent = "Stop Sync";
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

    updateProgress("Reading Excel data...", 0);

    try {
        // Suspend auto-calculation and events for extreme performance boost
        await Excel.run(async (context) => {
            context.runtime.enableEvents = false;
            context.workbook.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
        }).catch(() => {});

        // Local syncing is bypassed in multi-mode for safety
        if (!isMultiMode && queriesToFetch.length > 0) {
            const rawSheetData = await getSheetData();
            const syncPayload = parseSheetData(rawSheetData);
            const primaryCol = queriesToFetch[0].collection;

            if (syncPayload.inserts.length > 0 && currentSchema.length === 0) {
                showError("Cannot insert: no schema found in the sheet. Click '⬇ Schema' to import first.");
                setLoading(false);
                hideProgress();
                return;
            }
            
            let conflicts = [];

            if (syncPayload.inserts.length > 0 || syncPayload.updates.length > 0) {
                const chunkSize = 1000;
                const totalTasks = syncPayload.inserts.length + syncPayload.updates.length;
                let completedTasks = 0;
                
                for (let i = 0; i < syncPayload.inserts.length; i += chunkSize) {
                    if (isSyncCancelled) break;
                    const chunk = syncPayload.inserts.slice(i, i + chunkSize);
                    updateProgress(`Bulk inserting ${completedTasks + chunk.length} / ${totalTasks}...`, ((completedTasks + chunk.length) / totalTasks) * 100);
                    const res = await fetchWithTimeout(`${API_BASE}/bulk_insert`, {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ collection: primaryCol, data: chunk })
                    });
                    const data = await res.json();
                    if (!res.ok) throw new Error(data.detail || "Bulk insert failed");
                    inserted += data.inserted || 0;
                    completedTasks += chunk.length;
                }

                for (let i = 0; i < syncPayload.updates.length; i += chunkSize) {
                    if (isSyncCancelled) break;
                    const chunk = syncPayload.updates.slice(i, i + chunkSize);
                    updateProgress(`Bulk updating ${completedTasks + chunk.length} / ${totalTasks}...`, ((completedTasks + chunk.length) / totalTasks) * 100);
                    const res = await fetchWithTimeout(`${API_BASE}/bulk_update`, {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ collection: primaryCol, data: chunk })
                    });
                    const data = await res.json();
                    if (!res.ok) throw new Error(data.detail || "Bulk update failed");
                    updated += data.updated || 0;
                    if (data.conflicts) conflicts.push(...data.conflicts);
                    completedTasks += chunk.length;
                }

                if (conflicts.length > 0) {
                    await Excel.run(async (context) => {
                        const sheet = context.workbook.worksheets.getActiveWorksheet();
                        conflicts.forEach(c => {
                            if (c._rowIndex !== undefined) {
                                const range = sheet.getRangeByIndexes(c._rowIndex, 0, 1, currentSchema.length || 10);
                                range.format.fill.color = "#FFCCCC";
                            }
                        });
                        await context.sync();
                    });
                    showError(`${conflicts.length} conflict(s) detected. Conflicting rows are red. Sync again to overwrite with server data.`);
                    setLoading(false);
                    hideProgress();
                    return;
                }
            }
        }

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
                ? `Sync stopped by user. Loaded ${totalFetched} records.`
                : (totalFetched > 0 
                    ? `Sync process complete! Loaded ${totalFetched} records.`
                    : `Sync complete. No records matched the query.`),
            totalFetched, inserted, updated
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

// --- Delete Selected Functionality ---
async function runDeleteSelected() {
    const collection = collectionSelect.value;
    if (!collection) {
        showError("Please select a database collection first.");
        return;
    }

    if (!currentSchema || currentSchema.length === 0) {
        showError("No schema loaded. Please load data or click '⬇ Schema' first to ensure we can map IDs.");
        return;
    }

    const idIndex = currentSchema.indexOf("_id");
    if (idIndex === -1) {
        showError("Could not find the '_id' column in the current schema. Cannot delete without IDs.");
        return;
    }

    const btnDelete = document.getElementById('btn-delete');
    const deleteText = btnDelete.querySelector(".btn-text");
    const deleteSpinner = btnDelete.querySelector(".btn-spinner");

    deleteText.textContent = "Deleting...";
    deleteSpinner.classList.remove("hidden");
    btnDelete.disabled = true;
    hideFeedback();

    try {
        let selectedIds = [];
        let rowsToClear = [];

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const selectedRange = context.workbook.getSelectedRange();
            selectedRange.load(["rowIndex", "rowCount"]);
            await context.sync();

            const startRow = selectedRange.rowIndex;
            const rowCount = selectedRange.rowCount;

            // Load just the _id column for the selected rows
            const idRange = sheet.getRangeByIndexes(startRow, idIndex, rowCount, 1);
            idRange.load("values");
            await context.sync();

            idRange.values.forEach((row, i) => {
                const idVal = row[0];
                if (idVal && String(idVal).trim() !== "" && String(idVal).trim() !== "_id") {
                    selectedIds.push(String(idVal).trim());
                    rowsToClear.push(startRow + i);
                }
            });
        });

        if (selectedIds.length === 0) {
            showError("No valid records selected. Make sure you select cells belonging to rows with an '_id'.");
            return;
        }

        const confirmDelete = await customConfirm(`Are you sure you want to delete ${selectedIds.length} record(s)? This cannot be undone.`);
        if (!confirmDelete) {
            showSuccess("Delete cancelled.", 0, 0, 0);
            return;
        }

        updateProgress("Deleting from server...", 50);

        const res = await fetchWithTimeout(`${API_BASE}/bulk_delete`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ collection, ids: selectedIds })
        });
        const data = await res.json();

        if (!res.ok) throw new Error(data.detail || "Bulk delete failed");

        updateProgress("Clearing rows from sheet...", 90);

        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            for (let rIndex of rowsToClear) {
                sheet.getRangeByIndexes(rIndex, 0, 1, currentSchema.length).clear();
            }
            await context.sync();
        });

        showSuccess(`Successfully deleted ${data.deleted_count || selectedIds.length} record(s).`, 0, 0, 0);

    } catch (error) {
        showError("Delete failed: " + error.message);
    } finally {
        deleteText.textContent = "Delete Selected";
        deleteSpinner.classList.add("hidden");
        btnDelete.disabled = false;
        hideProgress();
    }
}

// Custom Confirm Dialog for Office.js
function customConfirm(message) {
    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.style.position = 'fixed';
        overlay.style.top = '0';
        overlay.style.left = '0';
        overlay.style.width = '100%';
        overlay.style.height = '100%';
        overlay.style.backgroundColor = 'rgba(0,0,0,0.5)';
        overlay.style.display = 'flex';
        overlay.style.alignItems = 'center';
        overlay.style.justifyContent = 'center';
        overlay.style.zIndex = '9999';

        const dialog = document.createElement('div');
        dialog.style.backgroundColor = 'var(--color-surface, #fff)';
        dialog.style.padding = '20px';
        dialog.style.borderRadius = '8px';
        dialog.style.boxShadow = '0 4px 12px rgba(0,0,0,0.15)';
        dialog.style.maxWidth = '85%';
        dialog.style.textAlign = 'center';

        const text = document.createElement('p');
        text.textContent = message;
        text.style.marginBottom = '20px';
        text.style.color = 'var(--color-text, #333)';
        text.style.fontWeight = '500';

        const btnContainer = document.createElement('div');
        btnContainer.style.display = 'flex';
        btnContainer.style.gap = '10px';
        btnContainer.style.justifyContent = 'center';

        const btnCancel = document.createElement('button');
        btnCancel.textContent = 'Cancel';
        btnCancel.style.padding = '8px 16px';
        btnCancel.style.border = '1px solid var(--color-border, #ccc)';
        btnCancel.style.borderRadius = '4px';
        btnCancel.style.background = 'transparent';
        btnCancel.style.cursor = 'pointer';
        btnCancel.style.flex = '1';

        const btnOk = document.createElement('button');
        btnOk.textContent = 'Delete';
        btnOk.style.padding = '8px 16px';
        btnOk.style.border = '1px solid #E02424';
        btnOk.style.borderRadius = '4px';
        btnOk.style.backgroundColor = '#E02424';
        btnOk.style.color = '#fff';
        btnOk.style.cursor = 'pointer';
        btnOk.style.flex = '1';

        btnCancel.onclick = () => {
            document.body.removeChild(overlay);
            resolve(false);
        };

        btnOk.onclick = () => {
            document.body.removeChild(overlay);
            resolve(true);
        };

        btnContainer.appendChild(btnCancel);
        btnContainer.appendChild(btnOk);
        dialog.appendChild(text);
        dialog.appendChild(btnContainer);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);
    });
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
    const fields = await detectSheetSchema();
    currentSchema = fields;
    
    let countText = "";
    const collection = collectionSelect.value;
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
    btnRun.disabled = isLoading;
    collectionSelect.disabled = isLoading;
    btnRefreshCollections.disabled = isLoading;
    queryInput.disabled = isLoading;

    if (isLoading) {
        btnText.textContent = "Processing...";
        btnSpinner.classList.remove("hidden");
    } else {
        btnText.textContent = "Sync & Fetch Data";
        btnSpinner.classList.add("hidden");
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
    const uri = document.getElementById('mongo-uri-input').value.trim();
    const dbName = document.getElementById('mongo-db-input').value.trim();
    const errEl = document.getElementById('login-error');
    const loader = document.getElementById('login-loader');
    
    if (!uri || !dbName) {
        errEl.textContent = "Please fill in both fields.";
        errEl.classList.remove('hidden');
        return;
    }
    
    const useUserApiBase = document.getElementById('use-custom-backend-checkbox').checked;
    let userApiBase = useUserApiBase ? document.getElementById('api-base-input').value.trim() : "";
    
    if (useUserApiBase && userApiBase) {
        API_BASE = userApiBase.replace(/\/+$/, "");
    } else {
        API_BASE = (window.EXCEL_MONGO_CONFIG && window.EXCEL_MONGO_CONFIG.apiBase) || window.location.origin;
    }
    
    errEl.classList.add('hidden');
    loader.classList.remove('hidden');
    document.getElementById('btn-login').disabled = true;
    
    currentMongoUri = uri;
    currentMongoDb = dbName;
    
    try {
        const res = await fetch(`${API_BASE}/health`, {
            headers: getAuthHeaders()
        });
        if (!res.ok) {
            const data = await res.json();
            throw new Error(data.detail || "Connection failed.");
        }
        
        Office.context.document.settings.set("MongoUri", uri);
        Office.context.document.settings.set("MongoDb", dbName);
        if (useUserApiBase && userApiBase) {
            Office.context.document.settings.set("ApiBase", API_BASE);
        } else {
            Office.context.document.settings.remove("ApiBase");
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

// New Collection UI Logic
document.getElementById('btn-new-collection').addEventListener('click', () => {
    document.getElementById('new-collection-container').classList.remove('hidden');
    document.getElementById('new-collection-input').focus();
});

document.getElementById('btn-cancel-collection').addEventListener('click', () => {
    document.getElementById('new-collection-container').classList.add('hidden');
    document.getElementById('new-collection-input').value = "";
});

document.getElementById('btn-create-collection').addEventListener('click', async () => {
    const input = document.getElementById('new-collection-input');
    const schemaInput = document.getElementById('new-collection-schema');
    const collName = input.value.trim();
    const schemaStr = schemaInput.value.trim();
    if (!collName) return;

    let fields = schemaStr.split(',').map(s => s.trim()).filter(s => s && s !== '_id');
    fields = ['_id', ...fields];

    const createBtn = document.getElementById('btn-create-collection');
    createBtn.textContent = "Creating...";
    createBtn.disabled = true;

    try {
        const response = await fetch(`${API_BASE}/create_collection`, {
            method: "POST",
            headers: getAuthHeaders({ "Content-Type": "application/json" }),
            body: JSON.stringify({ collection: collName })
        });

        const data = await response.json();
        if (response.ok) {
            document.getElementById('new-collection-container').classList.add('hidden');
            input.value = "";
            if(schemaInput) schemaInput.value = "";
            showSuccess(`Collection '${collName}' created!`);
            
            // Reload collections and auto-select
            await loadCollections();
            const select = document.getElementById("collection-select");
            select.value = collName;
            
            // Create Excel sheet and write schema
            await Excel.run(async (context) => {
                let sheet = context.workbook.worksheets.getItemOrNullObject(collName);
                await context.sync();
                
                if (sheet.isNullObject) {
                    sheet = context.workbook.worksheets.add(collName);
                }
                sheet.activate();
                sheet.getUsedRange().clear();
                
                const range = sheet.getRangeByIndexes(0, 0, 1, fields.length);
                range.values = [fields];
                
                sheet.customProperties.add("syncCollection", collName);
                sheet.customProperties.add("syncQuery", "{}");
                
                range.format.font.bold = true;
                range.format.fill.color = "#FFD300";
                range.format.font.color = "black";
                range.format.borders.getItem('EdgeBottom').style = 'Continuous';
                range.format.borders.getItem('EdgeBottom').weight = 'Thick';
                
                await context.sync();
            });

            // Manually trigger change to load schema
            const event = new Event('change');
            select.dispatchEvent(event);
        } else {
            showError(data.detail || "Failed to create collection");
        }
    } catch (e) {
        showError(e.message);
    } finally {
        createBtn.textContent = "Create";
        createBtn.disabled = false;
    }
});

function hideFeedback() {
    feedbackPanel.classList.add("hidden");
}
