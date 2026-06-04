// Config
const API_BASE = window.location.origin;

// DOM Elements
const connectionStatus = document.getElementById("connection-status");
const collectionSelect = document.getElementById("collection-select");
const btnRefreshCollections = document.getElementById("btn-refresh-collections");
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

// Initialize Add-in
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        initApp();
    }
});

// App Initialization
async function initApp() {
    // Setup event listeners
    btnRefreshCollections.addEventListener("click", loadCollections);
    queryInput.addEventListener("input", validateQuerySyntax);
    btnRun.addEventListener("click", runSyncAndFetch);

    // Initial check and load
    await checkApiHealth();
    if (isApiOnline) {
        await loadCollections();
    }
}

// Check if FastAPI is running and connected to MongoDB
async function checkApiHealth() {
    try {
        updateStatus("checking", "Checking API...");
        const res = await fetch(`${API_BASE}/health`);
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
        showError("Backend API is offline. Ensure FastAPI is running on port 8000 and MongoDB is active.");
    }
}

// Load MongoDB Collections into selector dropdown
async function loadCollections() {
    if (!isApiOnline) return;
    
    collectionSelect.disabled = true;
    collectionSelect.innerHTML = '<option value="">Loading...</option>';
    
    try {
        const res = await fetch(`${API_BASE}/collections`);
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
            }
        } else {
            throw new Error(data.detail || "Failed to load collections");
        }
    } catch (err) {
        showError("Failed to fetch MongoDB collections: " + err.message);
        collectionSelect.innerHTML = '<option value="">Error loading collections</option>';
    }
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
        
        let inserted = 0;
        let updated = 0;

        // 2. Perform Sync if sheets had records (using backend /insert and /update endpoints)
        if (syncPayload.inserts.length > 0 || syncPayload.updates.length > 0) {
            const insertPromises = syncPayload.inserts.map(async (doc) => {
                const res = await fetch(`${API_BASE}/insert`, {
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

                const res = await fetch(`${API_BASE}/update`, {
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

        const fetchRes = await fetch(`${API_BASE}/fetch`, {
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

        // 5. Present statistics
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
        headerRange.format.fill.color = "#13AA52";
        headerRange.format.font.color = "#FFFFFF";
        headerRange.format.font.bold = true;

        // Auto format column widths
        range.format.autofitColumns();

        await context.sync();
    });
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
