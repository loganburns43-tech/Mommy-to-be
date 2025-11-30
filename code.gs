var CLIENT_SHEET_NAME = 'MTB Clients';
var LEDGER_SHEET_NAME = 'MTB Ledger';
var CONFIG_SHEET_NAME = 'MTB Config';

var CLIENT_HEADERS = [
  'ClientID',
  'Last Name',
  'First Name',
  'Phone',
  'Status (Active / Closed / Paused / Reopened)',
  'Reason/Notes',
  'Guardians (free text: mom/dad/guardian situation)',
  'Due Month (YYYY-MM or blank)',
  'Signature Base64 (contract)',
  'Last Activity (timestamp)',
  'Created At (timestamp)'
];

var LEDGER_HEADERS = [
  'Timestamp',
  'ClientID',
  'Client Name',
  'Action Type (Provision / Note / Close Case / Reopen Case / Status Update)',
  'Pack Type (Enrollment / Monthly / Emergency / Special / None)',
  'Items Given',
  'Qty/Details (optional text)',
  'Clerk Initials',
  'Signature Base64 (if captured for this action)',
  'Notes'
];

var CONFIG_SEED = [
  ['PROGRAM_NAME', 'Mommy-To-Be Provisions'],
  ['MONTHLY_LIMIT_DAYS', '30'],
  ['EMERGENCY_LIMIT_DAYS', '30']
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Mommy-To-Be Provisions Program');
}

function ensureSetup() {
  var ss = SpreadsheetApp.getActive();
  var results = [];
  var clientSheetResult = ensureSheet(ss, CLIENT_SHEET_NAME, CLIENT_HEADERS);
  if (clientSheetResult) {
    results.push(clientSheetResult);
  }
  var ledgerSheetResult = ensureSheet(ss, LEDGER_SHEET_NAME, LEDGER_HEADERS);
  if (ledgerSheetResult) {
    results.push(ledgerSheetResult);
  }
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  var createdConfig = false;
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    createdConfig = true;
    results.push('Created sheet: ' + CONFIG_SHEET_NAME);
  }
  if (!headerMatches(configSheet, ['Key', 'Value'])) {
    configSheet.clear();
    configSheet.getRange(1, 1, 1, 2).setValues([['Key', 'Value']]);
    if (!createdConfig) {
      results.push('Reset headers on sheet: ' + CONFIG_SHEET_NAME);
    }
  }
  seedConfig(configSheet);
  return results.length ? results.join('\n') : 'Sheets verified';
}

function ensureSheet(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  var created = false;
  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
  }
  if (!headerMatches(sheet, headers)) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (created) {
      return 'Created sheet: ' + name;
    }
    return 'Reset headers on sheet: ' + name;
  }
  if (created) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return 'Created sheet: ' + name;
  }
  return '';
}

function headerMatches(sheet, headers) {
  var range = sheet.getRange(1, 1, 1, headers.length);
  var values = range.getValues()[0];
  for (var i = 0; i < headers.length; i++) {
    if (String(values[i] || '').trim() !== headers[i]) {
      return false;
    }
  }
  return true;
}

function seedConfig(sheet) {
  var existing = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < existing.length; i++) {
    var key = String(existing[i][0] || '').trim();
    var val = String(existing[i][1] || '').trim();
    if (key) {
      map[key] = val;
    }
  }
  var rowsToAdd = [];
  for (var j = 0; j < CONFIG_SEED.length; j++) {
    var pair = CONFIG_SEED[j];
    if (!map[pair[0]]) {
      rowsToAdd.push(pair);
    }
  }
  if (rowsToAdd.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 2).setValues(rowsToAdd);
  }
}

function listClients() {
  var sheet = getClientSheet();
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var status = String(row[4] || '').trim();
    if (!status) {
      continue;
    }
    list.push({
      clientId: String(row[0] || '').trim(),
      lastName: String(row[1] || '').trim(),
      firstName: String(row[2] || '').trim(),
      phone: String(row[3] || '').trim(),
      status: status,
      lastActivity: String(row[9] || '').trim()
    });
  }
  return list;
}

function createClient(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var sheet = getClientSheet();
    var firstName = cleanText(payload.firstName);
    var lastName = cleanText(payload.lastName);
    var phone = cleanText(payload.phone);
    var clerk = cleanText(payload.clerk);
    if (!firstName || !lastName || !phone) {
      throw new Error('First name, last name, and phone are required.');
    }
    if (!clerk) {
      throw new Error('Clerk initials are required.');
    }
    var guardians = cleanText(payload.guardians);
    var dueMonth = cleanText(payload.dueMonth);
    var notes = cleanText(payload.notes);
    var signature = cleanText(payload.signature);
    if (!signature) {
      throw new Error('Contract signature is required.');
    }
    var now = new Date();
    var clientId = 'MTB-' + now.getTime() + '-' + Math.floor(Math.random() * 1000);
    var row = [
      clientId,
      lastName,
      firstName,
      phone,
      'Active',
      notes,
      guardians,
      dueMonth,
      signature,
      now,
      now
    ];
    sheet.appendRow(row);
    appendLedger({
      timestamp: now,
      clientId: clientId,
      clientName: firstName + ' ' + lastName,
      actionType: 'Status Update',
      packType: 'None',
      itemsGiven: 'Client created with signed contract',
      qtyDetails: 'Contract signed at intake',
      clerk: clerk,
      signature: signature,
      notes: notes
    });
    return { success: true, clientId: clientId };
  } finally {
    lock.releaseLock();
  }
}

function getClient(clientId) {
  var sheet = getClientSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]) === String(clientId)) {
      return buildClientObject(row);
    }
  }
  return null;
}

function buildClientObject(row) {
  return {
    clientId: String(row[0] || ''),
    lastName: String(row[1] || ''),
    firstName: String(row[2] || ''),
    phone: String(row[3] || ''),
    status: String(row[4] || ''),
    reason: String(row[5] || ''),
    guardians: String(row[6] || ''),
    dueMonth: String(row[7] || ''),
    signature: String(row[8] || ''),
    lastActivity: String(row[9] || ''),
    createdAt: String(row[10] || '')
  };
}

function getClientHistory(clientId, limit) {
  if (!limit || limit < 1) {
    limit = 10;
  }
  var all = getClientHistoryAll(clientId);
  return all.slice(0, limit);
}

function getClientHistoryAll(clientId) {
  var sheet = getLedgerSheet();
  var data = sheet.getDataRange().getValues();
  var rows = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[1]) === String(clientId)) {
      rows.push({
        timestamp: row[0],
        clientId: row[1],
        clientName: row[2],
        actionType: row[3],
        packType: row[4],
        itemsGiven: row[5],
        qtyDetails: row[6],
        clerk: row[7],
        signature: row[8],
        notes: row[9]
      });
    }
  }
  rows.sort(function(a, b) {
    return new Date(b.timestamp).getTime() - new Date(a.timestamp).getTime();
  });
  return rows;
}

function saveProvision(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var clientId = cleanText(payload.clientId);
    var itemsGiven = cleanText(payload.itemsGiven);
    var clerk = cleanText(payload.clerk);
    if (!clientId) {
      throw new Error('Client is required.');
    }
    if (!itemsGiven) {
      throw new Error('Items Given is required.');
    }
    if (!clerk) {
      throw new Error('Clerk initials are required.');
    }
    var sheet = getClientSheet();
    var info = findClientRow(sheet, clientId);
    if (!info) {
      throw new Error('Client not found.');
    }
    var now = new Date();
    sheet.getRange(info.rowNumber, 10).setValue(now);
    appendLedger({
      timestamp: now,
      clientId: clientId,
      clientName: info.name,
      actionType: 'Provision',
      packType: cleanText(payload.packType) || 'None',
      itemsGiven: itemsGiven,
      qtyDetails: cleanText(payload.qtyDetails),
      clerk: clerk,
      signature: '',
      notes: cleanText(payload.notes)
    });
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function closeCase(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var clientId = cleanText(payload.clientId);
    if (!clientId) {
      throw new Error('Client is required.');
    }
    var clerk = cleanText(payload.clerk);
    if (!clerk) {
      throw new Error('Clerk initials are required.');
    }
    var reason = cleanText(payload.reason);
    if (!reason) {
      throw new Error('Reason is required.');
    }
    var sheet = getClientSheet();
    var info = findClientRow(sheet, clientId);
    if (!info) {
      throw new Error('Client not found.');
    }
    var now = new Date();
    sheet.getRange(info.rowNumber, 5).setValue('Closed');
    sheet.getRange(info.rowNumber, 6).setValue(reason);
    sheet.getRange(info.rowNumber, 10).setValue(now);
    appendLedger({
      timestamp: now,
      clientId: clientId,
      clientName: info.name,
      actionType: 'Close Case',
      packType: 'None',
      itemsGiven: reason,
      qtyDetails: '',
      clerk: clerk,
      signature: '',
      notes: cleanText(payload.notes)
    });
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function reopenCase(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var clientId = cleanText(payload.clientId);
    if (!clientId) {
      throw new Error('Client is required.');
    }
    var clerk = cleanText(payload.clerk);
    if (!clerk) {
      throw new Error('Clerk initials are required.');
    }
    var reason = cleanText(payload.reason);
    if (!reason) {
      throw new Error('Reason is required.');
    }
    var sheet = getClientSheet();
    var info = findClientRow(sheet, clientId);
    if (!info) {
      throw new Error('Client not found.');
    }
    var now = new Date();
    sheet.getRange(info.rowNumber, 5).setValue('Reopened');
    sheet.getRange(info.rowNumber, 6).setValue(reason);
    sheet.getRange(info.rowNumber, 10).setValue(now);
    appendLedger({
      timestamp: now,
      clientId: clientId,
      clientName: info.name,
      actionType: 'Reopen Case',
      packType: 'None',
      itemsGiven: reason,
      qtyDetails: '',
      clerk: clerk,
      signature: '',
      notes: cleanText(payload.notes)
    });
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function saveContractSignature(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var clientId = cleanText(payload.clientId);
    var signature = cleanText(payload.signature);
    var clerk = cleanText(payload.clerk);
    if (!clientId) {
      throw new Error('Client is required.');
    }
    if (!signature) {
      throw new Error('Signature is required.');
    }
    if (!clerk) {
      throw new Error('Clerk initials are required.');
    }
    var sheet = getClientSheet();
    var info = findClientRow(sheet, clientId);
    if (!info) {
      throw new Error('Client not found.');
    }
    var now = new Date();
    sheet.getRange(info.rowNumber, 9).setValue(signature);
    sheet.getRange(info.rowNumber, 10).setValue(now);
    appendLedger({
      timestamp: now,
      clientId: clientId,
      clientName: info.name,
      actionType: 'Contract Signature',
      packType: 'None',
      itemsGiven: 'Contract signed',
      qtyDetails: '',
      clerk: clerk,
      signature: signature,
      notes: ''
    });
    return { success: true };
  } finally {
    lock.releaseLock();
  }
}

function appendLedger(entry) {
  var sheet = getLedgerSheet();
  var row = [
    entry.timestamp,
    entry.clientId,
    entry.clientName,
    entry.actionType,
    entry.packType,
    entry.itemsGiven,
    entry.qtyDetails,
    entry.clerk,
    entry.signature || '',
    entry.notes
  ];
  sheet.appendRow(row);
}

function getClientSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CLIENT_SHEET_NAME);
  if (!sheet) {
    ensureSetup();
    sheet = ss.getSheetByName(CLIENT_SHEET_NAME);
  }
  return sheet;
}

function getLedgerSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(LEDGER_SHEET_NAME);
  if (!sheet) {
    ensureSetup();
    sheet = ss.getSheetByName(LEDGER_SHEET_NAME);
  }
  return sheet;
}

function findClientRow(sheet, clientId) {
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[0]) === String(clientId)) {
      return {
        rowNumber: i + 1,
        name: String(row[2] || '') + ' ' + String(row[1] || '')
      };
    }
  }
  return null;
}

function cleanText(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}
