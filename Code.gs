// Van Stock — Google Apps Script backend
// Can be deployed from script.google.com directly — no Sheet needed first.

const SHEET_NAME  = 'Inventory';
const SS_KEY      = 'vanStockSpreadsheetId';

function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId    = props.getProperty(SS_KEY);

  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch(e) {
      // ID stored but file was deleted — create a new one
    }
  }

  // Create a brand new spreadsheet in the user's Drive
  const ss = SpreadsheetApp.create('Van Stock Inventory');
  props.setProperty(SS_KEY, ss.getId());
  return ss;
}

function getSheet() {
  const ss    = getOrCreateSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Remove default blank sheet if it exists
    const defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && ss.getSheets().length > 1) {
      ss.deleteSheet(defaultSheet);
    }
    sheet.appendRow(['id','name','category','unit','qty','lowStock','note']);
    // Style the header row
    sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#0d1117').setFontColor('#39ff14');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : 'get';
  if (action === 'get') return getItems();
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === 'set') return setItems(payload.items);
    return jsonResponse({ error: 'Unknown action' });
  } catch(err) {
    return jsonResponse({ error: err.toString() });
  }
}

function getItems() {
  const sheet = getSheet();
  const rows  = sheet.getDataRange().getValues();
  if (rows.length <= 1) return jsonResponse({ items: [] });
  const items = rows.slice(1).map(r => ({
    id:       String(r[0]),
    name:     String(r[1]),
    category: String(r[2]),
    unit:     String(r[3]),
    qty:      Number(r[4]),
    lowStock: Number(r[5]),
    note:     String(r[6] || ''),
  }));
  return jsonResponse({ items });
}

function setItems(items) {
  const sheet = getSheet();
  sheet.clearContents();
  sheet.appendRow(['id','name','category','unit','qty','lowStock','note']);
  if (items && items.length) {
    const rows = items.map(i => [
      i.id, i.name, i.category,
      i.unit, i.qty, i.lowStock, i.note || ''
    ]);
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return jsonResponse({ ok: true });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
