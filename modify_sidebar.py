import os
import re

out_path = '/Users/rmgx/Actice Projects/Excel-Mongo/fe_sheets/Sidebar.html'

with open(out_path, 'r') as f:
    content = f.read()

# Add runGoogleScript helper at the top of the JS block
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
content = content.replace('<script>\n', f'<script>\n{helper}\n')

# 1. Init block
init_replace = """
window.addEventListener("load", async () => {
    if (true) {
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
            console.error(e);
        }
    }
});
"""
# Replace the old Office.onReady
content = re.sub(r'window\.addEventListener\("load", async \(\) => \{\n\s*if \(true\) \{.*?(?=\n\}\);\n)', init_replace.strip(), content, flags=re.DOTALL)

# 2. saveSheetState
save_sheet = """
async function saveSheetState() {
    const col = collectionSelect.value || "";
    const query = queryInput.value || "";
    await runGoogleScript('saveSettings', { syncCollection: col, syncQuery: query });
}
"""
content = re.sub(r'async function saveSheetState\(\) \{.*?\n\}', save_sheet.strip(), content, flags=re.DOTALL)

# 3. loadSheetState
load_sheet = """
async function loadSheetState() {
    try {
        const settings = await runGoogleScript('getSettings');
        if (settings.syncCollection) collectionSelect.value = settings.syncCollection;
        if (settings.syncQuery) {
            queryInput.value = settings.syncQuery;
            validateQuerySyntax();
        }
    } catch (e) { console.error(e); }
}
"""
content = re.sub(r'async function loadSheetState\(\) \{.*?\n\}', load_sheet.strip(), content, flags=re.DOTALL)

# 4. onWorksheetActivated
# Remove Excel.run for onActivated, GAS doesn't have an exact equivalent client-side event for sheet activation easily,
# but we can poll or just ignore it. We'll remove the Excel.run part in initApp:
content = re.sub(r'await Excel\.run\(async \(context\) => \{.*?onActivated.*?\n\s*\}\)\.catch\(console\.error\);', '', content, flags=re.DOTALL)

# 5. getSheetData
get_sheet_data = """
async function getSheetData() {
    return await runGoogleScript('getSheetData');
}
"""
content = re.sub(r'async function getSheetData\(\) \{.*?\n\}', get_sheet_data.strip(), content, flags=re.DOTALL)

# 6. appendDataToSheet
append_data = """
async function appendDataToSheet(baseSheetName, records, headers, globalStartRow) {
    if (records.length === 0) return globalStartRow;
    return await runGoogleScript('appendDataToSheet', baseSheetName, JSON.stringify(records), headers, globalStartRow);
}
"""
content = re.sub(r'async function appendDataToSheet\(baseSheetName, records, headers, globalStartRow\) \{.*?\n\}', append_data.strip(), content, flags=re.DOTALL)

# 7. detectSheetSchema
detect_schema = """
async function detectSheetSchema() {
    try {
        const data = await runGoogleScript('getSheetData');
        if (!data || !data.values || data.values.length === 0) return [];
        const headers = data.values[0].filter(h => h !== null && h !== "");
        return headers.map(String).filter(h => h !== "__v");
    } catch {
        return [];
    }
}
"""
content = re.sub(r'async function detectSheetSchema\(\) \{.*?\n\}', detect_schema.strip(), content, flags=re.DOTALL)

# 8. writeSchemaToSheet
write_schema = """
async function writeSchemaToSheet(fields) {
    await runGoogleScript('clearSheet', collectionSelect.value);
    await runGoogleScript('appendDataToSheet', collectionSelect.value, "[]", fields, 1);
}
"""
content = re.sub(r'async function writeSchemaToSheet\(fields\) \{.*?\n\}', write_schema.strip(), content, flags=re.DOTALL)

# 9. btn-login Office save properties
login_save = """
        await runGoogleScript('saveSettings', { MongoUri: uri, MongoDb: dbName });
"""
content = re.sub(r'Office\.context\.document\.settings\.set.*?;.*?Office\.context\.document\.settings\.saveAsync\(\);', login_save.strip(), content, flags=re.DOTALL)

# 10. create collection clear and write schema
create_coll = """
            await runGoogleScript('clearSheet', collName);
            await runGoogleScript('appendDataToSheet', collName, "[]", fields, 1);
            await runGoogleScript('saveSettings', { syncCollection: collName, syncQuery: "{}" });
"""
content = re.sub(r'await Excel\.run\(async \(context\) => \{.*?sheet\.customProperties.*?\}\);', create_coll.strip(), content, flags=re.DOTALL)

# 11. highlight conflicts
highlight_conf = """
                await runGoogleScript('highlightConflicts', conflicts, currentSchema.length || 10);
"""
content = re.sub(r'await Excel\.run\(async \(context\) => \{.*?range\.format\.fill\.color.*?\}\);', highlight_conf.strip(), content, flags=re.DOTALL)

# 12. Delete Selected
delete_sel = """
        const { selectedIds, rowsToClear } = await runGoogleScript('getSelectedIds', idIndex);
"""
content = re.sub(r'let selectedIds = \[\];\s*let rowsToClear = \[\];\s*await Excel\.run\(async \(context\) => \{.*?\}\);', delete_sel.strip(), content, flags=re.DOTALL)

delete_clear = """
        await runGoogleScript('clearRows', rowsToClear, currentSchema.length);
"""
content = re.sub(r'await Excel\.run\(async \(context\) => \{.*?clear\(\).*?\}\);', delete_clear.strip(), content, flags=re.DOTALL)

# 13. Disable/enable events and calculation
content = re.sub(r'await Excel\.run\(async \(context\) => \{\s*context\.runtime\.enableEvents.*?\n\s*\}\)\.catch\(\(\) => \{\}\);', '', content, flags=re.DOTALL)

# 14. clear sheet for empty fetch
clear_empty = """
        if (totalFetched === 0) {
            await runGoogleScript('clearSheet', collectionSelect.value);
        }
"""
content = re.sub(r'if \(totalFetched === 0\) \{\s*await Excel\.run\(.*?\);.*?\}', clear_empty.strip(), content, flags=re.DOTALL)

with open(out_path, 'w') as f:
    f.write(content)

print("Modified Sidebar.html successfully")
