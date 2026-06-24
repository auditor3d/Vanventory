// Van Stock — Google Apps Script backend
// Paste this entire file into script.google.com

const SHEET_NAME = 'Inventory';

function doGet(e) {
  const action = e.parameter.action;
  if (action === 'get') {
    return getItems();
  }
  return jsonResponse({ error: 'Unknown action' });
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (payload.action === 'set') {
      return setItems(payload.items);
    }
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
    id:       r[0],
    name:     r[1],
    category: r[2],
    unit:     r[3],
    qty:      Number(r[4]),
    lowStock: Number(r[5]),
    note:     r[6],
  }));
  return jsonResponse({ items });
}

function setItems(items) {
  const sheet = getSheet();
  sheet.clearContents();
  sheet.appendRow(['id','name','category','unit','qty','lowStock','note']);
  if (items && items.length) {
    const rows = items.map(i => [
      i.id, i.name, i.category, i.unit,
      i.qty, i.lowStock, i.note || ''
    ]);
    sheet.getRange(2, 1, rows.length, 7).setValues(rows);
  }
  return jsonResponse({ ok: true });
}

function getSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id','name','category','unit','qty','lowStock','note']);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
