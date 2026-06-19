import os
import re

base_dir = '/Users/rmgx/Actice Projects/Excel-Mongo'

with open(os.path.join(base_dir, 'fe/taskpane.html'), 'r') as f:
    html = f.read()

with open(os.path.join(base_dir, 'fe/taskpane.js'), 'r') as f:
    js = f.read()

with open(os.path.join(base_dir, 'fe/taskpane.css'), 'r') as f:
    css = f.read()

# JS Replacements

# Add runGoogleScript helper
helper = """
function runGoogleScript(funcName, ...args) {
    return new Promise((resolve, reject) => {
        google.script.run
            .withSuccessHandler(resolve)
            .withFailureHandler(reject)
            [funcName](...args);
    });
}
"""

js = helper + "\n" + js

# Replace API_BASE
js = re.sub(r'const API_BASE = .*?;', 'const API_BASE = "https://excel-mongo-api.up.railway.app";', js)

# 2. Init Block
old_init = """// Initialize Add-in
window.addEventListener("load", async () => {
    if (true) {"""

# Replace Office.onReady
js = re.sub(r'Office\.onReady\(async \(info\) => \{\n\s*if \(info\.host === Office\.HostType\.Excel\) \{.*?(?=\n\}\);\n)', 
"""window.addEventListener("load", async () => {
    try {
        const settings = await runGoogleScript('getSettings');
        currentMongoUri = settings.MongoUri || "";
        currentMongoDb = settings.MongoDb || "";
        
        if (currentMongoUri && currentMongoDb) {
            document.getElementById('login-view').classList.add('hidden');
            document.getElementById('app-view').classList.remove('hidden');
            await initApp();
        } else {
            document.getElementById('login-view').classList.remove('hidden');
            document.getElementById('app-view').classList.add('hidden');
        }
    } catch (e) {
        console.error("Init error:", e);
    }""", js, flags=re.DOTALL)

# 3. saveSheetState
js = re.sub(r'async function saveSheetState\(\) \{.*?\n\}', """async function saveSheetState() {
    const col = collectionSelect.value || "";
    const query = queryInput.value || "";
    await runGoogleScript('saveSettings', { syncCollection: col, syncQuery: query });
}""", js, flags=re.DOTALL)

# 4. loadSheetState
js = re.sub(r'async function loadSheetState\(\) \{.*?\n\}', """async function loadSheetState() {
    try {
        const settings = await runGoogleScript('getSettings');
        if (settings.syncCollection) collectionSelect.value = settings.syncCollection;
        if (settings.syncQuery) {
            queryInput.value = settings.syncQuery;
            validateQuerySyntax();
        }
    } catch (e) { console.error(e); }
}""", js, flags=re.DOTALL)

# 5. initApp onActivated
js = re.sub(r'await Excel\.run\(async \(context\) => \{\n\s*context\.workbook\.worksheets\.onActivated\.add\(onWorksheetActivated\);\n\s*await context\.sync\(\);\n\s*\}\)\.catch\(console\.error\);', '', js)

# 6. getSheetData
js = re.sub(r'async function getSheetData\(\) \{.*?\n\}', """async function getSheetData() {
    return await runGoogleScript('getSheetData');
}""", js, flags=re.DOTALL)

# 7. appendDataToSheet
js = re.sub(r'async function appendDataToSheet\(baseSheetName, records, headers, globalStartRow\) \{.*?\n\}', """async function appendDataToSheet(baseSheetName, records, headers, globalStartRow) {
    if (records.length === 0) return globalStartRow;
    return await runGoogleScript('appendDataToSheet', baseSheetName, JSON.stringify(records), headers, globalStartRow);
}""", js, flags=re.DOTALL)

# 8. detectSheetSchema
js = re.sub(r'async function detectSheetSchema\(\) \{.*?\n\}', """async function detectSheetSchema() {
    try {
        const data = await runGoogleScript('getSheetData');
        if (!data || !data.values || data.values.length === 0) return [];
        const headers = data.values[0].filter(h => h !== null && h !== "");
        return headers.map(String).filter(h => h !== "__v");
    } catch (e) {
        return [];
    }
}""", js, flags=re.DOTALL)

# 9. writeSchemaToSheet
js = re.sub(r'async function writeSchemaToSheet\(fields\) \{.*?\n\}', """async function writeSchemaToSheet(fields) {
    await runGoogleScript('clearSheet', collectionSelect.value);
    await runGoogleScript('appendDataToSheet', collectionSelect.value, JSON.stringify([]), fields, 1);
}""", js, flags=re.DOTALL)

# 10. Login save properties
js = re.sub(r'Office\.context\.document\.settings\.set\("MongoUri", uri\);\n\s*Office\.context\.document\.settings\.set\("MongoDb", dbName\);\n\s*Office\.context\.document\.settings\.saveAsync\(\);', 
"""await runGoogleScript('saveSettings', { MongoUri: uri, MongoDb: dbName });""", js)

# 11. Create collection write schema
create_coll = """// Create Excel sheet and write schema
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
            });"""
js = js.replace(create_coll, """// Create Sheet and write schema
            await runGoogleScript('clearSheet', collName);
            await runGoogleScript('appendDataToSheet', collName, JSON.stringify([]), fields, 1);
            await runGoogleScript('saveSettings', { syncCollection: collName, syncQuery: "{}" });""")

# 12. Suspend events
suspend = """// Suspend auto-calculation and events for extreme performance boost
        await Excel.run(async (context) => {
            context.runtime.enableEvents = false;
            context.workbook.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
        }).catch(() => {});"""
js = js.replace(suspend, "")

restore = """// Restore auto-calculation and events
        await Excel.run(async (context) => {
            context.runtime.enableEvents = true;
            context.workbook.application.calculationMode = Excel.CalculationMode.automatic;
            await context.sync();
        }).catch(() => {});"""
js = js.replace(restore, "")

# 13. Highlight conflicts
conflict = """await Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    conflicts.forEach(c => {
                        if (c._rowIndex !== undefined) {
                            const range = sheet.getRangeByIndexes(c._rowIndex, 0, 1, currentSchema.length || 10);
                            range.format.fill.color = "#FFCCCC";
                        }
                    });
                    await context.sync();
                });"""
js = js.replace(conflict, "await runGoogleScript('highlightConflicts', conflicts, currentSchema.length || 10);")

# 14. Clear sheet totalFetched 0
clear_fetched_0 = """await Excel.run(async (context) => {
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                try { sheet.getUsedRange().clear(); await context.sync(); } catch(e){}
            });"""
js = js.replace(clear_fetched_0, "await runGoogleScript('clearSheet', collectionSelect.value);")

# 15. runDeleteSelected
del_1 = """let selectedIds = [];
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
        });"""
js = js.replace(del_1, "const { selectedIds, rowsToClear } = await runGoogleScript('getSelectedIds', idIndex);")

del_2 = """await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            for (let rIndex of rowsToClear) {
                sheet.getRangeByIndexes(rIndex, 0, 1, currentSchema.length).clear();
            }
            await context.sync();
        });"""
js = js.replace(del_2, "await runGoogleScript('clearRows', rowsToClear, currentSchema.length);")

# HTML processing
html = re.sub(r'<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js".*?</script>', '', html)
html = re.sub(r'<script src="config.js"></script>', '', html)
html = re.sub(r'<link rel="stylesheet" href="taskpane.css">', f'<style>{css}</style>', html)
html = re.sub(r'<script src="taskpane.js\?v=5"></script>', '', html)
html = html.replace('</body>', f'<script>\n{js}\n</script>\n</body>')

with open(os.path.join(base_dir, 'fe_sheets/Sidebar.html'), 'w') as f:
    f.write(html)

print("Generated clean Sidebar.html")
