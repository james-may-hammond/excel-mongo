function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('MongoDB Sync')
      .addItem('Open Sidebar', 'showSidebar')
      .addToUi();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('MongoDB Sync')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getSettings() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return {
    MongoUri: documentProperties.getProperty('MongoUri') || '',
    MongoDb: documentProperties.getProperty('MongoDb') || '',
    syncCollection: documentProperties.getProperty('syncCollection') || '',
    syncQuery: documentProperties.getProperty('syncQuery') || ''
  };
}

function saveSettings(settings) {
  var documentProperties = PropertiesService.getDocumentProperties();
  for (var key in settings) {
    if (settings[key] !== undefined && settings[key] !== null) {
      documentProperties.setProperty(key, settings[key]);
    }
  }
}

function getSheetData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();
  return {
    values: values,
    rowCount: values.length,
    columnCount: values.length > 0 ? values[0].length : 0
  };
}

function appendDataToSheet(baseSheetName, dataString, headers, globalStartRow) {
  var data = JSON.parse(dataString);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var ROW_LIMIT = 500000;
  var sheetIndex = Math.floor((globalStartRow - 1) / ROW_LIMIT) + 1;
  var localStartRow = ((globalStartRow - 1) % ROW_LIMIT) + 1;
  
  var targetSheetName = sheetIndex === 1 ? baseSheetName : baseSheetName + "_" + sheetIndex;
  var sheet = ss.getSheetByName(targetSheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(targetSheetName);
  }
  
  if (localStartRow === 1) {
    sheet.clear();
    if (headers && headers.length > 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    }
    localStartRow = 2;
    globalStartRow++;
  }
  
  if (data.length === 0) return globalStartRow;
  
  var rows = [];
  for (var i = 0; i < data.length; i++) {
    var record = data[i];
    var row = [];
    for (var j = 0; j < headers.length; j++) {
      var val = record[headers[j]];
      row.push(val === undefined || val === null ? "" : val);
    }
    rows.push(row);
  }
  
  sheet.getRange(localStartRow, 1, rows.length, headers.length).setValues(rows);
  return globalStartRow + rows.length;
}

function clearSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear();
  } else {
    ss.getActiveSheet().clear();
  }
}

function highlightConflicts(conflicts, schemaLength) {
  var sheet = SpreadsheetApp.getActiveSheet();
  conflicts.forEach(function(c) {
    if (c._rowIndex !== undefined) {
      var range = sheet.getRange(c._rowIndex, 1, 1, schemaLength || 10);
      range.setBackground("#FFCCCC");
    }
  });
}

function getSelectedIds(idIndex) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeRange = sheet.getActiveRange();
  if (!activeRange) return { selectedIds: [], rowsToClear: [] };
  
  var startRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  
  var idRange = sheet.getRange(startRow, idIndex + 1, numRows, 1);
  var values = idRange.getValues();
  
  var selectedIds = [];
  var rowsToClear = [];
  
  for (var i = 0; i < values.length; i++) {
    var idVal = values[i][0];
    if (idVal && String(idVal).trim() !== "" && String(idVal).trim() !== "_id") {
      selectedIds.push(String(idVal).trim());
      rowsToClear.push(startRow + i);
    }
  }
  
  return { selectedIds: selectedIds, rowsToClear: rowsToClear };
}

function clearRows(rowsToClear, schemaLength) {
  var sheet = SpreadsheetApp.getActiveSheet();
  for (var i = 0; i < rowsToClear.length; i++) {
    sheet.getRange(rowsToClear[i], 1, 1, schemaLength).clearContent();
  }
}
