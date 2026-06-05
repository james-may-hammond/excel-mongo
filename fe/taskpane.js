// Config: everything served from https://localhost:8000 (mkcert cert trusted by WKWebView)
const API_BASE = "https://localhost:8000";

// Fetch with a timeout so we fail fast instead of hanging
async function fetchWithTimeout(url, options = {}, timeoutMs = 6000) {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    try {
        const res = await fetch(url, { ...options, signal: controller.signal });
        return res;
    } finally {
        clearTimeout(timer);
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
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initApp();
    }
});

// App Initialization
async function initApp() {
    // Collection controls
    btnRefreshCollections.addEventListener('click', loadCollections);
    btnImportSchema.addEventListener('click', importSchema);
    collectionSelect.addEventListener('change', onCollectionChange);

    // Query builder controls
    document.getElementById('tab-builder').addEventListener('click', () => switchQueryMode('builder'));
    document.getElementById('tab-json').addEventListener('click',    () => switchQueryMode('json'));
    document.getElementById('btn-add-condition').addEventListener('click', handleAddCondition);
    document.getElementById('btn-logic-and').addEventListener('click', () => setGuiLogic('and'));
    document.getElementById('btn-logic-or').addEventListener('click',  () => setGuiLogic('or'));

    // JSON textarea
    queryInput.addEventListener('input', validateQuerySyntax);

    // Main button
    btnRun.addEventListener('click', runSyncAndFetch);

    // Boot sequence
    await checkApiHealth();
    if (isApiOnline) await loadCollections();
    await refreshSchemaBar();
}

// Check if FastAPI is running and connected to MongoDB
async function checkApiHealth() {
    try {
        updateStatus("checking", "Checking API...");
        const res = await fetchWithTimeout(`${API_BASE}/health`);
        const data = await res.json();
        
        if (data.status === "ok") {
            isApiOnline = true;
            updateStatus("online", "Connected to Mongo");
            btnRun.disabled = false;
        } else {
            throw new Error("Health check returned status " + data.status);
        }
    } catch (err) {
        isApiOnline = false;
        updateStatus("offline", "API Offline");
        btnRun.disabled = true;
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
        return true;
    } catch (err) {
        jsonValidity.className = "validity-indicator invalid";
        jsonValidity.textContent = "Invalid JSON";
        btnRun.disabled = true;
        return false;
    }
}

// Core Workflow: Sync Excel adjustments then Fetch matching data
async function runSyncAndFetch() {
    const collection = collectionSelect.value;
    if (!collection) {
        showError("Please select a database collection first.");
        return;
    }

    // Double check syntax validation
    if (!validateQuerySyntax()) {
        showError("Query filter is not valid JSON.");
        return;
    }

    setLoading(true);
    hideFeedback();

    try {
        // 1. Read sheet data to prepare inserts/updates
        const rawSheetData = await getSheetData();
        const syncPayload = parseSheetData(rawSheetData);

        // Guard: if there are new rows to insert but no schema in row 1, block it
        if (syncPayload.inserts.length > 0 && currentSchema.length === 0) {
            showError(
                "Cannot insert: no schema found in the sheet. " +
                "Click '⬇ Schema' to import the collection's schema first, " +
                "or run a Sync & Fetch to load existing data."
            );
            setLoading(false);
            return;
        }
        
        let inserted = 0;
        let updated = 0;

        // 2. Perform Sync if sheets had records (using backend /insert and /update endpoints)
        if (syncPayload.inserts.length > 0 || syncPayload.updates.length > 0) {
            const insertPromises = syncPayload.inserts.map(async (doc) => {
                const res = await fetchWithTimeout(`${API_BASE}/insert`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        collection: collection,
                        data: doc
                    })
                });
                const data = await res.json();
                if (!res.ok) {
                    throw new Error(data.detail || "Insert failed");
                }
                return true;
            });

            const updatePromises = syncPayload.updates.map(async (doc) => {
                const docId = doc._id;
                const docData = { ...doc };
                delete docData._id;

                const res = await fetchWithTimeout(`${API_BASE}/update`, {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        collection: collection,
                        id: docId,
                        data: docData
                    })
                });
                const data = await res.json();
                if (!res.ok) {
                    throw new Error(data.detail || "Update failed");
                }
                return true;
            });

            await Promise.all([...insertPromises, ...updatePromises]);
            inserted = syncPayload.inserts.length;
            updated = syncPayload.updates.length;
        }


        // 3. Fetch data matching query filters
        let filters = {};
        if (queryInput.value.trim()) {
            filters = JSON.parse(queryInput.value.trim());
        }

        const fetchRes = await fetchWithTimeout(`${API_BASE}/fetch`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                collection: collection,
                filters: filters
            })
        });
        const fetchData = await fetchRes.json();
        
        if (!fetchRes.ok) {
            throw new Error(fetchData.detail || "Fetch query failed on backend");
        }

        const records = fetchData.data || [];
        
        // 4. Overwrite worksheet with fetched records
        await writeDataToSheet(records);

        // 5. Refresh schema bar to show updated headers
        await refreshSchemaBar();

        // 6. Present statistics
        showSuccess(
            records.length > 0 
                ? `Sync process complete! Loaded ${records.length} records into the sheet.`
                : `Sync complete. No records matched the query. Excel sheet cleared.`,
            records.length,
            inserted,
            updated
        );
    } catch (error) {
        showError("Execution failed: " + error.message);
    } finally {
        setLoading(false);
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

    dataRows.forEach(row => {
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
            if (doc["_id"]) {
                payload.updates.push(doc);
            } else {
                payload.inserts.push(doc);
            }
        }
    });

    return payload;
}

// Overwrite the current active worksheet
async function writeDataToSheet(records) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Safely clear sheet before rewriting
        try {
            const usedRange = sheet.getUsedRange();
            usedRange.clear();
            await context.sync();
        } catch (err) {
            // Sheet is already clear
        }

        if (records.length === 0) {
            return;
        }

        // Dynamically compile headers from records (putting _id first)
        const headerSet = new Set();
        headerSet.add("_id");
        records.forEach(rec => {
            Object.keys(rec).forEach(key => headerSet.add(key));
        });
        const headers = Array.from(headerSet);

        // Build values array
        const values = [headers];
        records.forEach(rec => {
            const row = [];
            headers.forEach(h => {
                let cell = rec[h];
                if (cell !== null && typeof cell === "object") {
                    cell = JSON.stringify(cell);
                } else if (cell === undefined) {
                    cell = "";
                }
                row.push(cell);
            });
            values.push(row);
        });

        // Set cells range starting at A1 (0, 0)
        const rowCount = values.length;
        const colCount = headers.length;
        const range = sheet.getRangeByIndexes(0, 0, rowCount, colCount);
        range.values = values;

        // Header styles (MongoDB Dark Green theme)
        const headerRange = sheet.getRangeByIndexes(0, 0, 1, colCount);
        headerRange.format.fill.color = "#5BAD7F";
        headerRange.format.font.color = "#FFFFFF";
        headerRange.format.font.bold = true;

        // Auto format column widths
        range.format.autofitColumns();

        await context.sync();
    });
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
            return headers.map(String);
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
    renderSchemaBar(fields);
    refreshConditionFields();
}

/**
 * Renders the schema field tags inside the schema bar.
 */
function renderSchemaBar(fields) {
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

/**
 * Called when user clicks ＋ Add Condition.
 * Auto-fetches schema first if none is loaded, so the field is always a dropdown.
 */
async function handleAddCondition() {
    const btn = document.getElementById('btn-add-condition');

    // If schema is already known, just add the row immediately
    if (currentSchema.filter(f => f !== '_id').length > 0) {
        addCondition();
        return;
    }

    // No schema yet — try to fetch it silently from the selected collection
    const collection = collectionSelect.value;
    if (!collection) {
        showError('Select a collection first so we can load its fields.');
        return;
    }

    btn.disabled = true;
    btn.textContent = 'Loading fields...';

    try {
        const res  = await fetchWithTimeout(`${API_BASE}/schema?collection=${encodeURIComponent(collection)}`);
        const data = await res.json();
        if (res.ok && data.fields && data.fields.length > 0) {
            currentSchema = data.fields;
            renderSchemaBar(data.fields);
        }
    } catch (_) {
        // Silently fall back — addCondition will use free-text if schema still empty
    } finally {
        btn.disabled = false;
        btn.innerHTML = '&#xFF0B; Add Condition';
    }

    addCondition();
}

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

/** Toggle AND / OR logic */
function setGuiLogic(logic) {
    guiLogic = logic;
    document.getElementById('btn-logic-and').classList.toggle('active', logic === 'and');
    document.getElementById('btn-logic-or').classList.toggle('active',  logic === 'or');
    updateQueryFromGui();
}

/** Add a new condition row to the builder */
function addCondition(field = '', op = 'eq', value = '') {
    const id = ++conditionCounter;
    const row = document.createElement('div');
    row.className = 'condition-row';
    row.dataset.id = id;

    // Field: dropdown if schema known, else free-text input
    let fieldHtml;
    const schemaFields = currentSchema.filter(f => f !== '_id');
    if (schemaFields.length > 0) {
        const opts = schemaFields
            .map(f => `<option value="${f}"${f === field ? ' selected' : ''}>${f}</option>`)
            .join('');
        fieldHtml = `<select class="cond-field cond-select">${opts}</select>`;
    } else {
        fieldHtml = `<input class="cond-field cond-input" type="text" placeholder="field" value="${field}">`;
    }

    // Operator dropdown
    const opOpts = GUI_OPERATORS
        .map(o => `<option value="${o.value}"${o.value === op ? ' selected' : ''}>${o.label}</option>`)
        .join('');

    const selectedOp = GUI_OPERATORS.find(o => o.value === op);
    const hideValue  = selectedOp && !selectedOp.hasValue ? 'hidden' : '';

    row.innerHTML = `
        ${fieldHtml}
        <select class="cond-op cond-select">${opOpts}</select>
        <input  class="cond-value cond-input ${hideValue}" type="text" placeholder="value" value="${value}">
        <button class="btn-remove-cond" title="Remove">&#x2715;</button>
    `;

    // Operator change: toggle value input visibility
    row.querySelector('.cond-op').addEventListener('change', e => {
        const def = GUI_OPERATORS.find(o => o.value === e.target.value);
        row.querySelector('.cond-value').classList.toggle('hidden', !def.hasValue);
        updateQueryFromGui();
    });
    row.querySelector('.btn-remove-cond').addEventListener('click', () => {
        row.remove();
        _updateConditionCount();
        updateQueryFromGui();
    });
    row.querySelector('.cond-field').addEventListener('change', updateQueryFromGui);
    row.querySelector('.cond-field').addEventListener('input',  updateQueryFromGui);
    row.querySelector('.cond-value').addEventListener('input',  updateQueryFromGui);

    document.getElementById('conditions-list').appendChild(row);
    _updateConditionCount();
    updateQueryFromGui();
}

/** Update the condition count label */
function _updateConditionCount() {
    const n = document.getElementById('conditions-list').children.length;
    const el = document.getElementById('condition-count');
    el.textContent = n === 0 ? 'fetch all' : n === 1 ? '1 condition' : `${n} conditions`;
}

/** Rebuild GUI filter to a MongoDB query object */
function buildGuiQuery() {
    const rows = document.querySelectorAll('.condition-row');
    if (rows.length === 0) return {};

    const parseVal = v => {
        if (v === 'true')  return true;
        if (v === 'false') return false;
        const n = Number(v);
        return (!isNaN(n) && v !== '') ? n : v;
    };

    const conditions = [];
    rows.forEach(row => {
        const field = row.querySelector('.cond-field').value.trim();
        const op    = row.querySelector('.cond-op').value;
        const raw   = (row.querySelector('.cond-value').value || '').trim();
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

    if (conditions.length === 0) return {};
    if (conditions.length === 1) return conditions[0];
    return guiLogic === 'and' ? { $and: conditions } : { $or: conditions };
}

/** Sync GUI → queryInput textarea + JSON preview */
function updateQueryFromGui() {
    const query   = buildGuiQuery();
    const isEmpty = Object.keys(query).length === 0;
    const pretty  = isEmpty ? '' : JSON.stringify(query, null, 2);

    queryInput.value = pretty;
    document.getElementById('query-preview-code').textContent = isEmpty ? '{}  // fetch all' : pretty;
    validateQuerySyntax();
}

/** When schema updates, rebuild condition field dropdowns in-place */
function refreshConditionFields() {
    const schemaFields = currentSchema.filter(f => f !== '_id');
    if (schemaFields.length === 0) return;
    document.querySelectorAll('.condition-row').forEach(row => {
        const fieldEl = row.querySelector('.cond-field');
        if (fieldEl.tagName === 'INPUT') {
            // Replace free-text with dropdown now that we have schema
            const cur  = fieldEl.value;
            const opts = schemaFields
                .map(f => `<option value="${f}"${f === cur ? ' selected' : ''}>${f}</option>`)
                .join('');
            const sel  = document.createElement('select');
            sel.className = 'cond-field cond-select';
            sel.innerHTML = opts;
            sel.addEventListener('change', updateQueryFromGui);
            fieldEl.replaceWith(sel);
        } else {
            // Already a select — update options keeping current selection
            const cur  = fieldEl.value;
            const opts = schemaFields
                .map(f => `<option value="${f}"${f === cur ? ' selected' : ''}>${f}</option>`)
                .join('');
            fieldEl.innerHTML = opts;
        }
    });
    updateQueryFromGui();
}

// --- Helpers ---
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

    feedbackMessage.className = "feedback-message success";
    feedbackMessage.textContent = msg;
    feedbackPanel.classList.remove("hidden");
}

function showError(msg) {
    statFetched.textContent = "0";
    statInserted.textContent = "0";
    statUpdated.textContent = "0";

    feedbackMessage.className = "feedback-message error";
    feedbackMessage.textContent = msg;
    feedbackPanel.classList.remove("hidden");
}

function hideFeedback() {
    feedbackPanel.classList.add("hidden");
}
