/**
 * MoneyFlow - Google Apps Script Backend
 * -------------------------------------------------------------
 * Deploy: Extensions > Apps Script (bound to the Google Sheet that
 * contains "Monthly_Transactions" and "Fixed_Templates" tabs).
 * Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone (or "Anyone with the link")
 * Copy the deployed /exec URL into GAS_URL in the frontend config.
 * -------------------------------------------------------------
 */

const SHEET_TXN = 'Monthly_Transactions';
const SHEET_TPL = 'Fixed_Templates';

const TXN_COLS = ['id', 'month_key', 'type', 'category', 'title',
  'planned_amount', 'actual_amount', 'is_fixed', 'is_settled',
  'settled_at', 'updated_at'];

const TPL_COLS = ['template_id', 'type', 'category', 'title',
  'default_amount', 'due_day', 'is_active'];

// ---------- entry points ----------

function doGet(e) {
  try {
    const action = (e.parameter.action || 'getData');
    if (action === 'getData') {
      return jsonOut(getData(e.parameter.month));
    }
    if (action === 'ping') {
      return jsonOut({ ok: true, ts: new Date().toISOString() });
    }
    return jsonOut({ error: 'unknown_action' });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action;
    if (action === 'batchSync') return jsonOut(batchSync(body.ops || []));
    if (action === 'generateMonth') return jsonOut(generateMonth(body.month));
    return jsonOut({ error: 'unknown_action' });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

// ---------- data access ----------

function getData(month) {
  const txnSheet = getSheet(SHEET_TXN);
  const tplSheet = getSheet(SHEET_TPL);
  let txns = sheetToObjects(txnSheet, TXN_COLS);
  if (month) txns = txns.filter(function (t) { return t.month_key === month; });
  const tpls = sheetToObjects(tplSheet, TPL_COLS);
  return { transactions: txns, templates: tpls, server_time: new Date().toISOString() };
}

// ops: [{sheet:'txn'|'tpl', op:'upsert'|'delete', data:{...}}]
function batchSync(ops) {
  const txnSheet = getSheet(SHEET_TXN);
  const tplSheet = getSheet(SHEET_TPL);
  const results = [];
  const now = new Date().toISOString();

  ops.forEach(function (item) {
    try {
      if (item.sheet === 'txn') {
        if (item.op === 'delete') {
          deleteRowByKey(txnSheet, 'id', item.data.id, TXN_COLS);
        } else {
          item.data.updated_at = now;
          upsertRow(txnSheet, 'id', item.data, TXN_COLS);
        }
      } else if (item.sheet === 'tpl') {
        if (item.op === 'delete') {
          deleteRowByKey(tplSheet, 'template_id', item.data.template_id, TPL_COLS);
        } else {
          upsertRow(tplSheet, 'template_id', item.data, TPL_COLS);
        }
      }
      results.push({ id: item.data.id || item.data.template_id, ok: true });
    } catch (err) {
      results.push({ id: item.data && (item.data.id || item.data.template_id), ok: false, error: String(err) });
    }
  });

  return { success: true, results: results, server_time: now };
}

// Create this month's transactions from active templates. Idempotent:
// generated ids are deterministic (TXN-<month>-<template_id>).
function generateMonth(month) {
  const tplSheet = getSheet(SHEET_TPL);
  const txnSheet = getSheet(SHEET_TXN);
  const tpls = sheetToObjects(tplSheet, TPL_COLS).filter(function (t) {
    return String(t.is_active).toUpperCase() === 'TRUE' || t.is_active === true;
  });
  const existing = sheetToObjects(txnSheet, TXN_COLS)
    .filter(function (t) { return t.month_key === month; })
    .map(function (t) { return t.id; });

  const now = new Date().toISOString();
  const created = [];

  tpls.forEach(function (tpl) {
    const id = 'TXN-' + month.replace('-', '') + '-' + tpl.template_id;
    if (existing.indexOf(id) !== -1) return; // already generated
    const row = {
      id: id,
      month_key: month,
      type: tpl.type,
      category: tpl.category,
      title: tpl.title,
      planned_amount: tpl.default_amount,
      actual_amount: tpl.default_amount,
      is_fixed: true,
      is_settled: false,
      settled_at: '',
      updated_at: now
    };
    upsertRow(txnSheet, 'id', row, TXN_COLS);
    created.push(id);
  });

  return { success: true, created: created, server_time: now };
}

// ---------- sheet helpers ----------

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function sheetToObjects(sheet, cols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, cols.length).getValues();
  return values
    .filter(function (row) { return row[0] !== '' && row[0] !== null; })
    .map(function (row) {
      const obj = {};
      cols.forEach(function (c, i) { obj[c] = row[i]; });
      return obj;
    });
}

function findRowIndexByKey(sheet, keyCol, keyVal, cols) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const colIndex = cols.indexOf(keyCol);
  const values = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(keyVal)) return i + 2; // 1-based, +header
  }
  return -1;
}

function upsertRow(sheet, keyCol, data, cols) {
  const rowIndex = findRowIndexByKey(sheet, keyCol, data[keyCol], cols);
  const rowValues = cols.map(function (c) {
    return (data[c] === undefined || data[c] === null) ? '' : data[c];
  });
  if (rowIndex === -1) {
    sheet.appendRow(rowValues);
  } else {
    // merge: keep existing values for fields not supplied in `data`
    const existing = sheet.getRange(rowIndex, 1, 1, cols.length).getValues()[0];
    const merged = cols.map(function (c, i) {
      return (data[c] === undefined) ? existing[i] : (data[c] === null ? '' : data[c]);
    });
    sheet.getRange(rowIndex, 1, 1, cols.length).setValues([merged]);
  }
}

function deleteRowByKey(sheet, keyCol, keyVal, cols) {
  const rowIndex = findRowIndexByKey(sheet, keyCol, keyVal, cols);
  if (rowIndex !== -1) sheet.deleteRow(rowIndex);
}

// ---------- output ----------

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
