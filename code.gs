var CLIENT_SHEET_NAME = 'MTB Clients';
var LEDGER_SHEET_NAME = 'MTB Ledger';
var CONFIG_SHEET_NAME = 'MTB Config';
var CASE_LOG_SHEET = 'Case Log';
var PROVISIONS_LOG_SHEET = 'Provisions Log';
var BIRTH_CLOSE_SHEET = 'Birth/Close Log';
var DEFAULT_CONTRACT_CREDIT = 30;

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

var CASE_LOG_HEADERS = [
  'Timestamp',
  'ClientID',
  'Client Name',
  'Action',
  'Reason/Notes',
  'Clerk',
  'Signature Base64'
];

var PROVISION_LOG_HEADERS = [
  'Timestamp',
  'ClientID',
  'Client Name',
  'Pack Type',
  'Items Given',
  'Qty/Details',
  'Final Price',
  'Clerk',
  'Signature Base64',
  'Notes'
];

var BIRTH_CLOSE_HEADERS = [
  'Timestamp',
  'ClientID',
  'Client Name',
  'Action (Birth/Close)',
  'Reason',
  'Clerk',
  'Notes'
];

var CONFIG_SEED = [
  ['PROGRAM_NAME', 'Mommy-To-Be Provisions'],
  ['MONTHLY_LIMIT_DAYS', '30'],
  ['EMERGENCY_LIMIT_DAYS', '30']
];

var MTB_MAIN_SHEET = 'Mommy to Be';
var MTB_CONTRACTS_SHEET = 'Mommy To Be – Contracts';
var MTB_ITEMS_SHEET = 'Mommy To Be – Items';

var MTB_MAIN_HEADERS = [
  'Timestamp',
  'Date',
  'Mom Last Name',
  'Mom First Name',
  'Guardians',
  'Due Date',
  'Baby Status (Active / Born / Miscarriage)',
  'Items Given',
  'Spent',
  'Clerk Initials',
  'Signature (base64 PNG)'
];

var MTB_CONTRACT_HEADERS = [
  'Mom Full Name',
  'Enrollment Date',
  'Status (Active / Paused / Closed / Reopened)',
  'Notes',
  'Last Updated'
];

var MTB_ITEM_HEADERS = [
  'Item Category',
  'Item Name'
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Mommy-To-Be Provisions Program');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mommy To Be Tools')
    .addItem('Generate Needed Sheets', 'mtb_generateSheets')
    .addItem('Recalculate Sheet Data', 'mtb_recalculateData')
    .addItem('Open Dashboard', 'mtb_openDashboard')
    .addItem('Rebuild Data Index', 'mtb_rebuildIndex')
    .addToUi();
}

function ensureSetup(options) {
  var createIfMissing = !(options && options.validateOnly === true);
  var ss = SpreadsheetApp.getActive();
  var results = [];
  var clientSheetResult = ensureSheet(ss, CLIENT_SHEET_NAME, CLIENT_HEADERS, createIfMissing);
  if (clientSheetResult) {
    results.push(clientSheetResult);
  }
  var ledgerSheetResult = ensureSheet(ss, LEDGER_SHEET_NAME, LEDGER_HEADERS, createIfMissing);
  if (ledgerSheetResult) {
    results.push(ledgerSheetResult);
  }
  var caseLogResult = ensureSheet(ss, CASE_LOG_SHEET, CASE_LOG_HEADERS, createIfMissing);
  if (caseLogResult) {
    results.push(caseLogResult);
  }
  var provisionResult = ensureSheet(ss, PROVISIONS_LOG_SHEET, PROVISION_LOG_HEADERS, createIfMissing);
  if (provisionResult) {
    results.push(provisionResult);
  }
  var birthCloseResult = ensureSheet(ss, BIRTH_CLOSE_SHEET, BIRTH_CLOSE_HEADERS, createIfMissing);
  if (birthCloseResult) {
    results.push(birthCloseResult);
  }
  var configResult = ensureConfigSheet(ss, createIfMissing);
  if (configResult) {
    results.push(configResult);
  }
  return results.length ? results.join('\n') : 'Sheets verified';
}

function ensureSheet(ss, name, headers, createIfMissing) {
  if (createIfMissing === undefined) {
    createIfMissing = true;
  }
  var sheet = ss.getSheetByName(name);
  var created = false;
  if (!sheet) {
    if (!createIfMissing) {
      return 'Missing sheet: ' + name;
    }
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

function ensureConfigSheet(ss, createIfMissing) {
  if (createIfMissing === undefined) {
    createIfMissing = true;
  }
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  var createdConfig = false;
  if (!sheet) {
    if (!createIfMissing) {
      return 'Missing sheet: ' + CONFIG_SHEET_NAME;
    }
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    createdConfig = true;
  }
  var headers = ['Key', 'Value'];
  if (!headerMatches(sheet, headers)) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (createdConfig) {
      seedConfig(sheet);
      return 'Created sheet: ' + CONFIG_SHEET_NAME;
    }
    seedConfig(sheet);
    return 'Reset headers on sheet: ' + CONFIG_SHEET_NAME;
  }
  seedConfig(sheet);
  if (createdConfig) {
    return 'Created sheet: ' + CONFIG_SHEET_NAME;
  }
  return '';
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
    var guardians = valueOrNA(payload.guardians);
    var dueMonth = valueOrNA(payload.dueMonth);
    var notes = valueOrNA(payload.notes);
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
      var client = buildClientObject(row);
      var spendInfo = computeClientSpend(client.clientId);
      var creditLimit = getContractCreditLimit();
      client.totalSpent = spendInfo.totalSpent;
      client.contractLimit = creditLimit;
      client.creditBalance = creditLimit - spendInfo.totalSpent;
      return client;
    }
  }
  return null;
}

function buildClientObject(row) {
  return {
    clientId: valueOrNA(row[0]),
    lastName: valueOrNA(row[1]),
    firstName: valueOrNA(row[2]),
    phone: valueOrNA(row[3]),
    status: valueOrNA(row[4]),
    reason: valueOrNA(row[5]),
    guardians: valueOrNA(row[6]),
    dueMonth: valueOrNA(row[7]),
    signature: valueOrNA(row[8]),
    lastActivity: valueOrNA(row[9]),
    createdAt: valueOrNA(row[10])
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

function computeClientSpend(clientId) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(PROVISIONS_LOG_SHEET);
  if (!sheet) {
    return { totalSpent: 0 };
  }
  var data = sheet.getDataRange().getValues();
  var total = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[1]) === String(clientId)) {
      var price = parseFloat(row[6]);
      if (!isNaN(price)) {
        total += price;
      }
    }
  }
  return { totalSpent: total };
}

function saveProvision(payload) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var clientId = cleanText(payload.clientId);
    var itemsGiven = cleanText(payload.itemsGiven);
    var clerk = cleanText(payload.clerk);
    var priceRaw = cleanText(payload.finalPrice);
    var finalPrice = 0;
    if (!clientId) {
      throw new Error('Client is required.');
    }
    if (!itemsGiven) {
      throw new Error('Items Given is required.');
    }
    if (!clerk) {
      throw new Error('Clerk initials are required.');
    }
    if (priceRaw === '') {
      throw new Error('Final price is required.');
    }
    if (priceRaw !== '') {
      var parsed = parseFloat(priceRaw);
      if (isNaN(parsed) || parsed < 0) {
        throw new Error('Final price must be a number 0 or higher.');
      }
      finalPrice = parsed;
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
      finalPrice: finalPrice,
      signature: '',
      notes: cleanText(payload.notes)
    });
    return getClient(clientId);
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
    valueOrNA(entry.itemsGiven),
    valueOrNA(entry.qtyDetails),
    valueOrNA(entry.clerk),
    entry.signature ? entry.signature : 'N/A',
    valueOrNA(entry.notes)
  ];
  sheet.appendRow(row);

  var ss = SpreadsheetApp.getActive();
  if (entry.actionType === 'Provision') {
    var provisionSheet = ensureSheet(ss, PROVISIONS_LOG_SHEET, PROVISION_LOG_HEADERS) ? ss.getSheetByName(PROVISIONS_LOG_SHEET) : ss.getSheetByName(PROVISIONS_LOG_SHEET);
    provisionSheet.appendRow([
      entry.timestamp,
      entry.clientId,
      entry.clientName,
      entry.packType,
      valueOrNA(entry.itemsGiven),
      valueOrNA(entry.qtyDetails),
      valueOrNA(entry.finalPrice || ''),
      valueOrNA(entry.clerk),
      entry.signature ? entry.signature : 'N/A',
      valueOrNA(entry.notes)
    ]);
  }

  if (entry.actionType === 'Close Case' || entry.actionType === 'Reopen Case' || entry.actionType === 'Status Update') {
    var caseSheet = ensureSheet(ss, CASE_LOG_SHEET, CASE_LOG_HEADERS) ? ss.getSheetByName(CASE_LOG_SHEET) : ss.getSheetByName(CASE_LOG_SHEET);
    caseSheet.appendRow([
      entry.timestamp,
      entry.clientId,
      entry.clientName,
      entry.actionType,
      valueOrNA(entry.itemsGiven),
      valueOrNA(entry.notes),
      valueOrNA(entry.clerk),
      entry.signature ? entry.signature : 'N/A'
    ]);
  }

  if (entry.actionType === 'Close Case' || entry.actionType === 'Reopen Case') {
    var bcSheet = ensureSheet(ss, BIRTH_CLOSE_SHEET, BIRTH_CLOSE_HEADERS) ? ss.getSheetByName(BIRTH_CLOSE_SHEET) : ss.getSheetByName(BIRTH_CLOSE_SHEET);
    bcSheet.appendRow([
      entry.timestamp,
      entry.clientId,
      entry.clientName,
      entry.actionType,
      valueOrNA(entry.itemsGiven),
      valueOrNA(entry.clerk),
      valueOrNA(entry.notes)
    ]);
  }
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

function valueOrNA(value) {
  var cleaned = cleanText(value);
  return cleaned ? cleaned : 'N/A';
}

function mtb_generateSheets() {
  var ss = SpreadsheetApp.getActive();
  ensureMtbMainSheet(ss);
  ensureMtbContractSheet(ss);
  ensureMtbItemSheet(ss);
  ensureConfigSheet(ss, true);
  SpreadsheetApp.getActive().toast('Mommy to Be sheets created.');
  return 'Mommy to Be sheets created.';
}

function mtb_recalculateData() {
  var ss = SpreadsheetApp.getActive();
  var messages = [];
  var main = ss.getSheetByName(MTB_MAIN_SHEET);
  if (!main) {
    messages.push('Missing sheet: ' + MTB_MAIN_SHEET + '. Run Generate Needed Sheets.');
  } else {
    try {
      getMtbMainSheetValidated();
    } catch (err) {
      messages.push(err.message || err);
    }
  }

  var configMessage = ensureConfigSheet(ss, false);
  if (configMessage) {
    messages.push(configMessage);
  }

  if (!messages.length) {
    SpreadsheetApp.flush();
    SpreadsheetApp.getActive().toast('Recalculated client index and history.');
    return 'Recalculated client index and history.';
  }

  var summary = messages.join('\n');
  SpreadsheetApp.getActive().toast(summary);
  return summary;
}

function mtb_openDashboard() {
  var html = HtmlService.createHtmlOutputFromFile('MommyToBe')
    .setWidth(1100)
    .setHeight(900);
  SpreadsheetApp.getUi().showModalDialog(html, 'Mommy-To-Be Dashboard');
}

function mtb_rebuildIndex() {
  var ss = SpreadsheetApp.getActive();
  var messages = [];
  var mainCheck = ss.getSheetByName(MTB_MAIN_SHEET);
  if (!mainCheck) {
    messages.push('Mommy to Be sheet is missing. Run Generate Needed Sheets.');
  } else {
    getMtbMainSheetValidated();
  }
  var configCheck = ensureConfigSheet(ss, false);
  if (configCheck) {
    messages.push(configCheck);
  }
  if (messages.length) {
    var summary = messages.join('\n');
    SpreadsheetApp.getActive().toast(summary);
    return summary;
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getActive().toast('Data index rebuilt and synced.');
  return 'Data index rebuilt and synced.';
}

function mtb_saveEntry(data) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var momFirstName = cleanText(data && data.momFirstName);
    var momLastName = cleanText(data && data.momLastName);
    var guardians = cleanText(data && data.guardians);
    var itemsGiven = cleanText(data && data.itemsGiven);
    var clerkInitials = cleanText(data && data.clerkInitials);
    var signature = cleanText(data && data.signature);

    if (!momFirstName || !momLastName) {
      throw new Error('Mom first and last name are required.');
    }
    if (!guardians) {
      throw new Error('Guardians are required.');
    }
    if (!itemsGiven) {
      throw new Error('Items Given is required.');
    }
    if (!clerkInitials) {
      throw new Error('Clerk initials are required.');
    }
    if (!signature) {
      throw new Error('Signature is required.');
    }

    var sheetInfo = getMtbMainSheetValidated();
    var sheet = sheetInfo.sheet;
    var timestamp = new Date();
    var dateValue = cleanText(data && data.date) ? new Date(data.date) : new Date();
    var dueDateValue = cleanText(data && data.dueDate);
    var babyStatus = cleanText(data && data.babyStatus) || 'Active';
    var spentRaw = cleanText(data && data.spent);
    var spentValue = '';
    if (spentRaw) {
      var parsed = parseFloat(spentRaw);
      if (isNaN(parsed)) {
        throw new Error('Spent must be a number.');
      }
      spentValue = parsed;
    }

    var row = [
      timestamp,
      dateValue,
      momLastName,
      momFirstName,
      guardians,
      dueDateValue || 'N/A',
      babyStatus,
      itemsGiven,
      spentValue === '' ? 'N/A' : spentValue,
      clerkInitials,
      signature
    ];
    sheet.appendRow(row);
    return { success: true, message: 'Entry saved.' };
  } finally {
    lock.releaseLock();
  }
}

function mtb_getHistory(fullName) {
  var name = cleanText(fullName);
  if (!name) {
    throw new Error('Full name is required to load history.');
  }
  var sheetInfo = getMtbMainSheetValidated();
  var sheet = sheetInfo.sheet;
  var data = sheet.getDataRange().getValues();
  var matches = [];
  var normalized = name.toLowerCase();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var first = String(row[3] || '').trim();
    var last = String(row[2] || '').trim();
    var combined = (first + ' ' + last).toLowerCase();
    var alternate = (last + ' ' + first).toLowerCase();
    if (combined === normalized || alternate === normalized) {
      matches.push({
        timestamp: row[0],
        date: row[1],
        momLastName: last,
        momFirstName: first,
        guardians: valueOrNA(row[4]),
        dueDate: valueOrNA(row[5]),
        babyStatus: valueOrNA(row[6]),
        itemsGiven: valueOrNA(row[7]),
        spent: valueOrNA(row[8]),
        clerkInitials: valueOrNA(row[9]),
        signature: valueOrNA(row[10])
      });
    }
  }
  matches.sort(function(a, b) {
    var aTime = new Date(a.timestamp).getTime();
    var bTime = new Date(b.timestamp).getTime();
    return bTime - aTime;
  });
  return matches;
}

function mtb_getContractSpend(fullName) {
  var name = cleanText(fullName);
  if (!name) {
    throw new Error('Full name is required to load spend.');
  }
  var sheetInfo = getMtbMainSheetValidated();
  var sheet = sheetInfo.sheet;
  var data = sheet.getDataRange().getValues();
  var normalized = name.toLowerCase();
  var total = 0;
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var first = String(row[3] || '').trim();
    var last = String(row[2] || '').trim();
    var combined = (first + ' ' + last).toLowerCase();
    var alternate = (last + ' ' + first).toLowerCase();
    if (combined === normalized || alternate === normalized) {
      var spentVal = parseFloat(row[8]);
      if (!isNaN(spentVal)) {
        total += spentVal;
      }
      count++;
    }
  }
  return { totalSpent: total, entries: count };
}

function ensureMtbMainSheet(ss) {
  var sheet = ss.getSheetByName(MTB_MAIN_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MTB_MAIN_SHEET);
  }
  ensureSheetColumns(sheet, MTB_MAIN_HEADERS);
  return sheet;
}

function ensureMtbContractSheet(ss) {
  var sheet = ss.getSheetByName(MTB_CONTRACTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MTB_CONTRACTS_SHEET);
  }
  ensureSheetColumns(sheet, MTB_CONTRACT_HEADERS);
  return sheet;
}

function ensureMtbItemSheet(ss) {
  var sheet = ss.getSheetByName(MTB_ITEMS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(MTB_ITEMS_SHEET);
  }
  ensureSheetColumns(sheet, MTB_ITEM_HEADERS);
  if (sheet.getLastRow() < 2) {
    sheet.getRange(2, 1, 4, 2).setValues([
      ['Diapers', ''],
      ['Wipes', ''],
      ['Formula', ''],
      ['Clothing (0–3m, 3–6m, etc.)', '']
    ]);
  }
  return sheet;
}

function ensureSheetColumns(sheet, headers) {
  var needed = headers.length;
  var maxColumns = sheet.getMaxColumns();
  if (maxColumns < needed) {
    sheet.insertColumnsAfter(maxColumns, needed - maxColumns);
  }
  maxColumns = sheet.getMaxColumns();
  var totalColumns = Math.max(maxColumns, headers.length);
  var row = [];
  for (var i = 0; i < totalColumns; i++) {
    row[i] = headers[i] || '';
  }
  sheet.getRange(1, 1, 1, totalColumns).setValues([row]);
}

function getContractCreditLimit() {
  var ss = SpreadsheetApp.getActive();
  var config = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!config) {
    return DEFAULT_CONTRACT_CREDIT;
  }
  var data = config.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === 'CONTRACT_CREDIT') {
      var val = parseFloat(data[i][1]);
      if (!isNaN(val)) {
        return val;
      }
    }
  }
  return DEFAULT_CONTRACT_CREDIT;
}

function getMtbMainSheetValidated() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(MTB_MAIN_SHEET);
  if (!sheet) {
    throw new Error('Mommy to Be sheet is missing. Please run Generate Required Sheets.');
  }
  var headerRange = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), MTB_MAIN_HEADERS.length));
  var headers = headerRange.getValues()[0];
  var spentIndex = -1;
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').trim() === 'Spent') {
      spentIndex = i;
      break;
    }
  }
  if (spentIndex === -1) {
    throw new Error("Cannot find 'Spent' column in Mommy to Be sheet.");
  }
  return { sheet: sheet, spentIndex: spentIndex };
}
