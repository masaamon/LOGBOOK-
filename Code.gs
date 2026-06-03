const SS_ID = 'YOUR_SPREADSHEET_ID_HERE';
// ═══════════════════════════════════════════
// WEB APP
// ═══════════════════════════════════════════
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Lab Logbook')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
// ═══════════════════════════════════════════
// VALIDATION HELPERS
// ═══════════════════════════════════════════
function validatePayload_(payload, requiredFields) {
  if (!payload) throw new Error('ไม่มีข้อมูลส่งมา');
  for (var i = 0; i < requiredFields.length; i++) {
    var f = requiredFields[i];
    if (payload[f] === undefined || payload[f] === null || String(payload[f]).trim() === '') {
      throw new Error('กรุณากรอก ' + f);
    }
  }
}
function validatePositiveNumber_(val, fieldName) {
  var n = Number(val);
  if (!isFinite(n) || n <= 0) {
    throw new Error(fieldName + ' ต้องเป็นตัวเลขมากกว่า 0');
  }
  return n;
}
function validateNonNegativeNumber_(val, fieldName) {
  var n = Number(val);
  if (!isFinite(n) || n < 0) {
    throw new Error(fieldName + ' ต้องเป็นตัวเลข 0 ขึ้นไป');
  }
  return n;
}
function generateId_(prefix) {
  return prefix + Utilities.getUuid();
}
// ═══════════════════════════════════════════
// CORE HELPERS (Optimized)
// ═══════════════════════════════════════════
function getSS_() {
  return SpreadsheetApp.openById(SS_ID);
}
function sheetToJson_(sh) {
  if (!sh) return [];
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var data = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
  var headers = data[0];
  var hLen = headers.length;
  var rows = new Array(data.length - 1);
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    var row = data[i];
    for (var j = 0; j < hLen; j++) {
      var val = row[j];
      obj[headers[j]] = (typeof val === 'object' && val !== null && typeof val.getTime === 'function')
        ? val.toISOString() : val;
    }
    rows[i - 1] = obj;
  }
  return rows;
}
function clearCache_() {
  try {
    CacheService.getScriptCache().removeAll(['db_0','db_1','db_2','db_3','db_4','db_5','db_6','db_7','db_8','db_9']);
  } catch(e) {}
}
function appendRow_(ss, name, obj) {
  var sh = ss.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบ sheet: ' + name);
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(function(h) { return obj[h] !== undefined ? obj[h] : ''; }));
  clearCache_();
}
// Batch update — เขียนหลาย cell ในแถวเดียวพร้อมกัน
function batchUpdateRow_(sh, rowNum, headers, updates) {
  var rowData = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  for (var col in updates) {
    if (updates.hasOwnProperty(col)) {
      var ci = headers.indexOf(col);
      if (ci >= 0) rowData[ci] = updates[col];
    }
  }
  sh.getRange(rowNum, 1, 1, headers.length).setValues([rowData]);
}
// ═══════════════════════════════════════════
// API: GET DASHBOARD (Optimized cache)
// ═══════════════════════════════════════════
function apiGetDashboard() {
  var cache = CacheService.getScriptCache();
  var keys = ['db_0','db_1','db_2','db_3','db_4','db_5','db_6','db_7','db_8','db_9'];
  var cached = cache.getAll(keys);
  if (cached['db_0']) {
    try {
      var chunks = [];
      for (var i = 0; i < 10; i++) {
        if (cached['db_' + i]) chunks.push(cached['db_' + i]);
        else break;
      }
      if (chunks.length > 0) return JSON.parse(chunks.join(''));
    } catch(e) {}
  }
  var ss = getSS_();
  var sheetMap = {};
  ss.getSheets().forEach(function(s) { sheetMap[s.getName()] = s; });
  var result = {
    items:        sheetToJson_(sheetMap['Items'] || null),
    chemicals:    sheetToJson_(sheetMap['Chemicals'] || null),
    tasks:        sheetToJson_(sheetMap['Tasks'] || null),
    dailyLogs:    sheetToJson_(sheetMap['DailyLogs'] || null),
    transactions: sheetToJson_(sheetMap['Transactions'] || null),
    experiments:  sheetToJson_(sheetMap['Experiments'] || null)
  };
  // เขียน cache — แก้ fix syntax error
  try {
    var json  = JSON.stringify(result);
    var size  = 90000;
    var total = Math.ceil(json.length / size);
    var entries = {};
    for (var k = 0; k < total; k++) {
      entries['db_' + k] = json.substring(k * size, (k + 1) * size);
    }
    for (var m = total; m < 10; m++) entries['db_' + m] = '';
    cache.putAll(entries, 120);
  } catch(e) {}
  return result;
}
// ═══════════════════════════════════════════
// API: CHECKOUT (Fixed + Optimized)
// ═══════════════════════════════════════════
function apiCheckout(payload) {
  validatePayload_(payload, ['itemId', 'qty', 'person']);
  var qty = validatePositiveNumber_(payload.qty, 'จำนวน');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var sh   = ss.getSheetByName('Items');
    if (!sh) throw new Error('ไม่พบ sheet: Items');
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var idI  = h.indexOf('itemId'), nmI = h.indexOf('itemName');
    var qtI  = h.indexOf('qtyAvailable');
    var rn = -1, row;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.itemId)) { rn = i + 1; row = data[i]; break; }
    }
    if (rn < 0) throw new Error('ไม่พบรายการ');
    var currentQty = Number(row[qtI]);
    if (!isFinite(currentQty)) throw new Error('stock ปัจจุบันมีปัญหา กรุณาตรวจสอบ');
    if (currentQty < qty) throw new Error('จำนวนไม่พอ (มี ' + currentQty + ' ต้องการ ' + qty + ')');
    var nq = currentQty - qty;
    batchUpdateRow_(sh, rn, h, { qtyAvailable: nq, updatedAt: new Date() });
    appendRow_(ss, 'Transactions', {
      txId: generateId_('TX-'), dateTime: new Date(),
      itemId: payload.itemId, itemName: row[nmI],
      type: 'OUT', qty: qty,
      person: String(payload.person).trim(), note: payload.note || '',
      expId: ''
    });
    return { ok: true, newQty: nq };
  } finally { lock.releaseLock(); }
}
// ═══════════════════════════════════════════
// API: RETURN (Fixed + Optimized)
// ═══════════════════════════════════════════
function apiReturn(payload) {
  validatePayload_(payload, ['itemId', 'qty', 'person']);
  var qty = validatePositiveNumber_(payload.qty, 'จำนวน');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var sh   = ss.getSheetByName('Items');
    if (!sh) throw new Error('ไม่พบ sheet: Items');
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var idI  = h.indexOf('itemId'), nmI = h.indexOf('itemName');
    var qtI  = h.indexOf('qtyAvailable');
    var rn = -1, row;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.itemId)) { rn = i + 1; row = data[i]; break; }
    }
    if (rn < 0) throw new Error('ไม่พบรายการ');
    var currentQty = Number(row[qtI]);
    if (!isFinite(currentQty)) throw new Error('stock ปัจจุบันมีปัญหา กรุณาตรวจสอบ');
    var nq = currentQty + qty;
    batchUpdateRow_(sh, rn, h, { qtyAvailable: nq, updatedAt: new Date() });
    appendRow_(ss, 'Transactions', {
      txId: generateId_('TX-'), dateTime: new Date(),
      itemId: payload.itemId, itemName: row[nmI],
      type: 'IN', qty: qty,
      person: String(payload.person).trim(), note: payload.note || '',
      expId: ''
    });
    return { ok: true, newQty: nq };
  } finally { lock.releaseLock(); }
}
// ═══════════════════════════════════════════
// API: USE CHEMICAL (Fixed + Optimized)
// ═══════════════════════════════════════════
function apiUseChemical(payload) {
  validatePayload_(payload, ['chemId', 'qty', 'person']);
  var qty = validatePositiveNumber_(payload.qty, 'จำนวน');
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var sh   = ss.getSheetByName('Chemicals');
    if (!sh) throw new Error('ไม่พบ sheet: Chemicals');
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var idI  = h.indexOf('chemId'), nmI = h.indexOf('chemName');
    var qtI  = h.indexOf('qtyRemaining');
    var rn = -1, row;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.chemId)) { rn = i + 1; row = data[i]; break; }
    }
    if (rn < 0) throw new Error('ไม่พบสารเคมี');
    var currentQty = Number(row[qtI]);
    if (!isFinite(currentQty)) throw new Error('stock สารเคมีมีปัญหา กรุณาตรวจสอบ');
    var nq = currentQty - qty;
    if (nq < 0) throw new Error('สารเคมีเหลือไม่พอ (มี ' + currentQty + ' ต้องการ ' + qty + ')');
    batchUpdateRow_(sh, rn, h, { qtyRemaining: nq, updatedAt: new Date() });
    appendRow_(ss, 'Transactions', {
      txId: generateId_('TX-'), dateTime: new Date(),
      itemId: payload.chemId, itemName: row[nmI],
      type: 'CHEM-USE', qty: qty,
      person: String(payload.person).trim(), note: payload.note || '',
      expId: payload.expId || ''
    });
    return { ok: true, newQty: nq };
  } finally { lock.releaseLock(); }
}
// ═══════════════════════════════════════════
// API: DELETE TRANSACTION (Fixed)
// ═══════════════════════════════════════════
function apiDeleteTransaction(payload) {
  validatePayload_(payload, ['txId']);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var shTx = ss.getSheetByName('Transactions');
    if (!shTx) throw new Error('ไม่พบ sheet: Transactions');
    var data = shTx.getDataRange().getValues();
    var h    = data[0];
    var idI  = h.indexOf('txId');
    var typeI = h.indexOf('type');
    var itemI = h.indexOf('itemId');
    var qtyI  = h.indexOf('qty');
    var rn = -1, txRow;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.txId)) { rn = i + 1; txRow = data[i]; break; }
    }
    if (rn < 0) throw new Error('ไม่พบ Transaction');
    var txType = String(txRow[typeI]);
    var txQty  = Number(txRow[qtyI]);
    var txItem = String(txRow[itemI]);
    if (!isFinite(txQty) || txQty <= 0) throw new Error('Transaction qty ไม่ถูกต้อง');
    // คืนสต็อกให้ถูก sheet ตาม type
    if (txType === 'CHEM-USE') {
      var shChem = ss.getSheetByName('Chemicals');
      if (!shChem) throw new Error('ไม่พบ sheet: Chemicals');
      var cData  = shChem.getDataRange().getValues();
      var ch     = cData[0];
      var cIdI   = ch.indexOf('chemId');
      var cQtI   = ch.indexOf('qtyRemaining');
      var found = false;
      for (var j = 1; j < cData.length; j++) {
        if (String(cData[j][cIdI]) === txItem) {
          var newChemQty = Number(cData[j][cQtI]) + txQty;
          batchUpdateRow_(shChem, j + 1, ch, { qtyRemaining: newChemQty, updatedAt: new Date() });
          found = true;
          break;
        }
      }
      if (!found) throw new Error('ไม่พบสารเคมี ' + txItem + ' ที่อ้างอิง');
    } else if (txType === 'OUT') {
      var shIt = ss.getSheetByName('Items');
      if (!shIt) throw new Error('ไม่พบ sheet: Items');
      var iData = shIt.getDataRange().getValues();
      var ih    = iData[0];
      var iIdI  = ih.indexOf('itemId');
      var iQtI  = ih.indexOf('qtyAvailable');
      var foundItem = false;
      for (var k = 1; k < iData.length; k++) {
        if (String(iData[k][iIdI]) === txItem) {
          var newItemQty = Number(iData[k][iQtI]) + txQty;
          batchUpdateRow_(shIt, k + 1, ih, { qtyAvailable: newItemQty, updatedAt: new Date() });
          foundItem = true;
          break;
        }
      }
      if (!foundItem) throw new Error('ไม่พบอุปกรณ์ ' + txItem + ' ที่อ้างอิง');
    } else if (txType === 'IN') {
      var shIt2 = ss.getSheetByName('Items');
      if (!shIt2) throw new Error('ไม่พบ sheet: Items');
      var iData2 = shIt2.getDataRange().getValues();
      var ih2    = iData2[0];
      var iIdI2  = ih2.indexOf('itemId');
      var iQtI2  = ih2.indexOf('qtyAvailable');
      var foundItem2 = false;
      for (var l = 1; l < iData2.length; l++) {
        if (String(iData2[l][iIdI2]) === txItem) {
          var currentStock = Number(iData2[l][iQtI2]);
          var newQ = currentStock - txQty;
          if (newQ < 0) {
            throw new Error('ไม่สามารถลบ Transaction นี้ได้ เพราะ stock ปัจจุบัน (' + currentStock + ') ไม่พอหักกลับ (' + txQty + ')');
          }
          batchUpdateRow_(shIt2, l + 1, ih2, { qtyAvailable: newQ, updatedAt: new Date() });
          foundItem2 = true;
          break;
        }
      }
      if (!foundItem2) throw new Error('ไม่พบอุปกรณ์ ' + txItem + ' ที่อ้างอิง');
    }
    shTx.deleteRow(rn);
    clearCache_();
    return { ok: true };
  } finally { lock.releaseLock(); }
}
// ═══════════════════════════════════════════
// API: DAILY LOG, TASK & EXPERIMENT (Fixed)
// ═══════════════════════════════════════════
function apiAddDailyLog(payload) {
  validatePayload_(payload, ['person', 'detail']);
  var ss = getSS_();
  appendRow_(ss, 'DailyLogs', {
    logId:     generateId_('LOG-'),
    logDate:   payload.logDate || new Date(),
    person:    String(payload.person).trim(),
    detail:    payload.detail,
    createdAt: new Date()
  });
  return { ok: true };
}
function apiAddTask(payload) {
  validatePayload_(payload, ['title']);
  var ss = getSS_();
  appendRow_(ss, 'Tasks', {
    taskId:       generateId_('TASK-'),
    title:        payload.title,
    detail:       payload.detail || '',
    owner:        payload.owner  || '',
    announceDate: new Date(),
    dueDate:      payload.dueDate || '',
    status:       'TODO'
  });
  return { ok: true };
}
function apiUpdateTaskStatus(payload) {
  validatePayload_(payload, ['taskId', 'status']);
  var validStatuses = ['TODO', 'DOING', 'DONE'];
  if (validStatuses.indexOf(payload.status) < 0) {
    throw new Error('status ต้องเป็น: ' + validStatuses.join(', '));
  }
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss = getSS_();
    var sh = ss.getSheetByName('Tasks');
    if (!sh) throw new Error('ไม่พบ sheet: Tasks');
    var data = sh.getDataRange().getValues();
    var h = data[0];
    var idI = h.indexOf('taskId');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.taskId)) {
        batchUpdateRow_(sh, i + 1, h, { status: payload.status });
        clearCache_();
        return { ok: true };
      }
    }
    throw new Error('ไม่พบ Task');
  } finally { lock.releaseLock(); }
}
function apiAddExperiment(payload) {
  validatePayload_(payload, ['expName']);
  var ss = getSS_();
  appendRow_(ss, 'Experiments', {
    expId:      generateId_('EXP-'),
    expName:    payload.expName,
    researcher: payload.researcher || '',
    status:     payload.status || 'Active',
    createdAt:  new Date()
  });
  return { ok: true };
}
// ═══════════════════════════════════════════
// API: ITEMS CRUD (Fixed + Optimized)
// ═══════════════════════════════════════════
function apiAddItem(payload) {
  validatePayload_(payload, ['itemName']);
  var initialQty = 0;
  if (payload.qtyAvailable !== undefined && payload.qtyAvailable !== '') {
    initialQty = validateNonNegativeNumber_(payload.qtyAvailable, 'จำนวนเริ่มต้น');
  }
  var ss = getSS_();
  appendRow_(ss, 'Items', {
    itemId:       generateId_('ITM-'),
    itemName:     payload.itemName,
    category:     payload.category     || 'ทั่วไป',
    size:         payload.size         || '',
    unit:         payload.unit         || 'อัน',
    qtyAvailable: initialQty,
    location:     payload.location     || '',
    updatedAt:    new Date()
  });
  return { ok: true };
}
function apiEditItem(payload) {
  validatePayload_(payload, ['itemId']);
  var editableFields = ['itemName', 'category', 'size', 'unit', 'location', 'qtyAvailable'];
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var sh   = ss.getSheetByName('Items');
    if (!sh) throw new Error('ไม่พบ sheet: Items');
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var idI  = h.indexOf('itemId');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.itemId)) {
        var rn = i + 1;
        var updates = { updatedAt: new Date() };
        editableFields.forEach(function(col) {
          if (payload[col] !== undefined) {
            if (col === 'qtyAvailable') {
              var v = validateNonNegativeNumber_(payload[col], 'จำนวน');
              updates[col] = v;
            } else {
              updates[col] = payload[col];
            }
          }
        });
        batchUpdateRow_(sh, rn, h, updates);
        clearCache_();
        return { ok: true };
      }
    }
    throw new Error('ไม่พบรายการ');
  } finally { lock.releaseLock(); }
}
function apiDeleteItem(payload) {
  validatePayload_(payload, ['itemId']);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var shTx = ss.getSheetByName('Transactions');
    if (shTx && shTx.getLastRow() > 1) {
      var txData = shTx.getDataRange().getValues();
      var txH = txData[0];
      var txItemI = txH.indexOf('itemId');
      for (var t = 1; t < txData.length; t++) {
        if (String(txData[t][txItemI]) === String(payload.itemId)) {
          throw new Error('ไม่สามารถลบได้ — ยังมี Transaction อ้างอิงอยู่ (' + (txData.length - 1) + ' รายการ) กรุณาลบ Transaction ก่อน');
        }
      }
    }
    var sh   = ss.getSheetByName('Items');
    if (!sh) throw new Error('ไม่พบ sheet: Items');
    var data = sh.getDataRange().getValues();
    var idI  = data[0].indexOf('itemId');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.itemId)) {
        sh.deleteRow(i + 1);
        clearCache_();
        return { ok: true };
      }
    }
    throw new Error('ไม่พบรายการ');
  } finally { lock.releaseLock(); }
}
// ═══════════════════════════════════════════
// API: CHEMICALS CRUD (Fixed + Optimized)
// ═══════════════════════════════════════════
function apiAddChemical(payload) {
  validatePayload_(payload, ['chemName']);
  var initialQty = 0;
  if (payload.qtyRemaining !== undefined && payload.qtyRemaining !== '') {
    initialQty = validateNonNegativeNumber_(payload.qtyRemaining, 'จำนวนสารเคมี');
  }
  var ss = getSS_();
  appendRow_(ss, 'Chemicals', {
    chemId:       generateId_('CHEM-'),
    chemName:     payload.chemName,
    casNo:        payload.casNo        || '',
    qtyRemaining: initialQty,
    unit:         payload.unit         || 'mL',
    maxQty:       Number(payload.maxQty) || initialQty,
    mfgDate:      payload.mfgDate      || '',
    expiryDate:   payload.expiryDate   || '',
    location:     payload.location     || '',
    msdsLink:     payload.msdsLink     || '',
    updatedAt:    new Date()
  });
  return { ok: true };
}
function apiEditChemical(payload) {
  validatePayload_(payload, ['chemId']);
  var editableFields = ['chemName', 'casNo', 'unit', 'maxQty', 'mfgDate', 'expiryDate', 'location', 'msdsLink', 'qtyRemaining'];
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var sh   = ss.getSheetByName('Chemicals');
    if (!sh) throw new Error('ไม่พบ sheet: Chemicals');
    var data = sh.getDataRange().getValues();
    var h    = data[0];
    var idI  = h.indexOf('chemId');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.chemId)) {
        var rn = i + 1;
        var updates = { updatedAt: new Date() };
        editableFields.forEach(function(col) {
          if (payload[col] !== undefined) {
            if (col === 'qtyRemaining') {
              updates[col] = validateNonNegativeNumber_(payload[col], 'ปริมาณคงเหลือ');
            } else if (col === 'maxQty') {
              updates[col] = validateNonNegativeNumber_(payload[col], 'ปริมาณสูงสุด');
            } else {
              updates[col] = payload[col];
            }
          }
        });
        batchUpdateRow_(sh, rn, h, updates);
        clearCache_();
        return { ok: true };
      }
    }
    throw new Error('ไม่พบสารเคมี');
  } finally { lock.releaseLock(); }
}
function apiDeleteChemical(payload) {
  validatePayload_(payload, ['chemId']);
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('กรุณาลองใหม่');
  try {
    var ss   = getSS_();
    var shTx = ss.getSheetByName('Transactions');
    if (shTx && shTx.getLastRow() > 1) {
      var txData = shTx.getDataRange().getValues();
      var txH = txData[0];
      var txItemI = txH.indexOf('itemId');
      for (var t = 1; t < txData.length; t++) {
        if (String(txData[t][txItemI]) === String(payload.chemId)) {
          throw new Error('ไม่สามารถลบได้ — ยังมี Transaction อ้างอิงอยู่ กรุณาลบ Transaction ก่อน');
        }
      }
    }
    var sh   = ss.getSheetByName('Chemicals');
    if (!sh) throw new Error('ไม่พบ sheet: Chemicals');
    var data = sh.getDataRange().getValues();
    var idI  = data[0].indexOf('chemId');
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idI]) === String(payload.chemId)) {
        sh.deleteRow(i + 1);
        clearCache_();
        return { ok: true };
      }
    }
    throw new Error('ไม่พบสารเคมี');
  } finally { lock.releaseLock(); }
}
// ═══════════════════════════════════════════
// TEST
// ═══════════════════════════════════════════
function testConnection() {
  var t = Date.now();
  try {
    var ss = getSS_();
    Logger.log('✅ เปิด Sheet ได้: ' + ss.getName());
    Logger.log('📄 Sheets: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
    var r = apiGetDashboard();
    Logger.log('📦 Items: '     + r.items.length);
    Logger.log('🧪 Chemicals: ' + r.chemicals.length);
    Logger.log('📋 Tasks: '     + r.tasks.length);
    Logger.log('📓 Logs: '      + r.dailyLogs.length);
    Logger.log('📜 Tx: '        + r.transactions.length);
    Logger.log('🔬 Experiments: ' + r.experiments.length);
    Logger.log('⏱️ เวลา: '     + ((Date.now() - t) / 1000).toFixed(2) + ' วินาที');
  } catch(e) {
    Logger.log('❌ ERROR: ' + e.message);
    Logger.log(e.stack);
  }
}
// ═══════════════════════════════════════════
// REGRESSION TESTS
// ═══════════════════════════════════════════
function runAllTests() {
  var pass = 0, fail = 0;
  function assert(name, fn) {
    try { fn(); Logger.log('✅ ' + name); pass++; }
    catch(e) { Logger.log('❌ ' + name + ' → ' + e.message); fail++; }
  }
  function expectError(name, fn) {
    try { fn(); Logger.log('❌ ' + name + ' (ควร error แต่ไม่ error)'); fail++; }
    catch(e) { Logger.log('✅ ' + name + ' → caught: ' + e.message); pass++; }
  }
  expectError('Checkout: qty เป็นลบ', function() {
    apiCheckout({ itemId: 'FAKE', qty: -5, person: 'Test' });
  });
  expectError('Checkout: qty = NaN', function() {
    apiCheckout({ itemId: 'FAKE', qty: 'abc', person: 'Test' });
  });
  expectError('Checkout: person ว่าง', function() {
    apiCheckout({ itemId: 'FAKE', qty: 1, person: '' });
  });
  expectError('Return: qty = 0', function() {
    apiReturn({ itemId: 'FAKE', qty: 0, person: 'Test' });
  });
  expectError('UseChemical: ไม่มี chemId', function() {
    apiUseChemical({ qty: 1, person: 'Test' });
  });
  expectError('AddItem: ไม่มี itemName', function() {
    apiAddItem({ category: 'test' });
  });
  expectError('UpdateTask: status ไม่ถูกต้อง', function() {
    apiUpdateTaskStatus({ taskId: 'FAKE', status: 'INVALID' });
  });
  expectError('DeleteItem: ไม่มี itemId', function() {
    apiDeleteItem({});
  });
  assert('UUID ไม่ซ้ำ', function() {
    var ids = {};
    for (var i = 0; i < 100; i++) {
      var id = generateId_('TEST-');
      if (ids[id]) throw new Error('ID ซ้ำ: ' + id);
      ids[id] = true;
    }
  });
  Logger.log('');
  Logger.log('═══════════════════════════');
  Logger.log('ผลทดสอบ: ' + pass + ' ผ่าน, ' + fail + ' ไม่ผ่าน');
  Logger.log('═══════════════════════════');
}
