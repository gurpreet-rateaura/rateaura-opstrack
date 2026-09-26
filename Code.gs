/**
 * RATEAURA FINTRACK — V1
 * Backend (Google Apps Script) — connects the web frontend to this Google Sheet.
 *
 * SETUP:
 * 1. Open your Google Sheet.
 * 2. Extensions > Apps Script.
 * 3. Delete any starter code and paste this whole file in.
 * 4. Set the PASSWORD constant below to whatever you want to log in with.
 * 5. Run the "setup" function once (Run > select "setup" > Run) to create all
 *    sheets, headers and default master lists. Approve the permissions it asks for.
 * 6. Deploy > New deployment > type "Web app".
 *      - Execute as: Me
 *      - Who has access: Anyone
 *    Click Deploy, authorize, and copy the Web App URL — you'll paste it into
 *    the API_URL constant near the top of index.html.
 * Full instructions are in README.md.
 */

// ⚠️ CHANGE THIS to your own password before deploying.
const PASSWORD = 'CHANGE_ME_123';

const ENTRY_SHEET = 'Entries';
const AGENTS_SHEET = 'Agents';
const SUPPLIERS_SHEET = 'Suppliers';
const NATURE_SHEET = 'NatureOfProfit';
const CONCLUSION_SHEET = 'FinalConclusion';

const ENTRY_HEADERS = [
  'ID', 'CreatedAt', 'ReferenceNo', 'AgentName', 'SupplierName',
  'SupplierConfirmationNo', 'AgentReferenceNo', 'ServiceName',
  'CheckInDate', 'CheckOutDate', 'SupplierBuyingAmount', 'AgentSellingAmount',
  'Status', 'NewBuyingAmount', 'NewSellingAmount', 'ConditionOfCalculation',
  'NatureOfProfit', 'Remarks', 'FinalConclusion', 'ResultAmount', 'ResultType',
  'UpdatedAt'
];

// ---------- One-time setup ----------
function setup() {
  const entrySheet = getOrCreateSheet(ENTRY_SHEET);
  ensureHeaders(entrySheet, ENTRY_HEADERS);

  seedMaster(AGENTS_SHEET, ['Sample Agent 1', 'Sample Agent 2']);
  seedMaster(SUPPLIERS_SHEET, ['Sample Supplier 1', 'Sample Supplier 2']);
  seedMaster(NATURE_SHEET, ['Rate Difference', 'Cancellation Charge', 'No Show', 'Amendment', 'Currency Fluctuation', 'Commission Adjustment', 'Other']);
  seedMaster(CONCLUSION_SHEET, ['Settled', 'Recovered from Supplier', 'Recovered from Agent', 'Written Off', 'Under Dispute', 'Disputed - Resolved']);

  SpreadsheetApp.getActiveSpreadsheet().toast('Setup complete. You can now deploy the Web App.');
}

function seedMaster(name, defaults) {
  const sheet = getOrCreateSheet(name);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, defaults.length, 1).setValues(defaults.map(v => [v]));
  }
}

function getOrCreateSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function ensureHeaders(sheet, headers) {
  const existing = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (existing.join('') === '') {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

// ---------- Web entry points ----------
function doGet(e) {
  const action = (e.parameter.action || 'getData');
  let out = { success: true };
  try {
    if (e.parameter.password !== PASSWORD) throw new Error('Invalid password');
    if (action === 'getData') {
      out.entries = getAllEntries();
      out.agents = getMasterList(AGENTS_SHEET);
      out.suppliers = getMasterList(SUPPLIERS_SHEET);
      out.natureOfProfit = getMasterList(NATURE_SHEET);
      out.finalConclusion = getMasterList(CONCLUSION_SHEET);
    } else {
      out = { success: false, error: 'Unknown action' };
    }
  } catch (err) {
    out = { success: false, error: err.message };
  }
  return jsonOut(out);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let out = { success: true };
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.password !== PASSWORD) throw new Error('Invalid password');

    switch (body.action) {
      case 'login':
        out = { success: true };
        break;
      case 'createEntry':
        out.entry = createEntry(body.data);
        break;
      case 'updateEntry':
        out.entry = updateEntry(body.data);
        break;
      case 'deleteEntry':
        deleteEntry(body.id);
        break;
      case 'addMaster':
        addMaster(body.list, body.value);
        break;
      case 'deleteMaster':
        deleteMaster(body.list, body.value);
        break;
      default:
        throw new Error('Unknown action');
    }
  } catch (err) {
    out = { success: false, error: err.message };
  } finally {
    lock.releaseLock();
  }
  return jsonOut(out);
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Entries ----------
function getAllEntries() {
  const sheet = getOrCreateSheet(ENTRY_SHEET);
  ensureHeaders(sheet, ENTRY_HEADERS);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const data = sheet.getRange(2, 1, lastRow - 1, ENTRY_HEADERS.length).getValues();
  return data
    .map(row => {
      const obj = {};
      ENTRY_HEADERS.forEach((h, i) => {
        let v = row[i];
        if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        obj[h] = v;
      });
      return obj;
    })
    .filter(o => o.ID); // ignore any stray blank rows
}

function findRowIndexById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

function createEntry(data) {
  const sheet = getOrCreateSheet(ENTRY_SHEET);
  ensureHeaders(sheet, ENTRY_HEADERS);
  const id = Utilities.getUuid();
  const now = new Date();
  const record = {
    ID: id,
    CreatedAt: now,
    UpdatedAt: now,
    Status: 'Open',
    ReferenceNo: data.ReferenceNo || '',
    AgentName: data.AgentName || '',
    SupplierName: data.SupplierName || '',
    SupplierConfirmationNo: data.SupplierConfirmationNo || '',
    AgentReferenceNo: data.AgentReferenceNo || '',
    ServiceName: data.ServiceName || '',
    CheckInDate: data.CheckInDate || '',
    CheckOutDate: data.CheckOutDate || '',
    SupplierBuyingAmount: Number(data.SupplierBuyingAmount) || 0,
    AgentSellingAmount: Number(data.AgentSellingAmount) || 0,
    NewBuyingAmount: '',
    NewSellingAmount: '',
    ConditionOfCalculation: '',
    NatureOfProfit: '',
    Remarks: '',
    FinalConclusion: '',
    ResultAmount: '',
    ResultType: ''
  };
  const row = ENTRY_HEADERS.map(h => record[h]);
  sheet.appendRow(row);
  return getEntryById(id);
}

function updateEntry(data) {
  const sheet = getOrCreateSheet(ENTRY_SHEET);
  const rowIndex = findRowIndexById(sheet, data.ID);
  if (rowIndex === -1) throw new Error('Entry not found (it may have been deleted elsewhere)');

  const currentRow = sheet.getRange(rowIndex, 1, 1, ENTRY_HEADERS.length).getValues()[0];
  const current = {};
  ENTRY_HEADERS.forEach((h, i) => (current[h] = currentRow[i]));

  // Editable original-entry fields
  const editableFields = [
    'ReferenceNo', 'AgentName', 'SupplierName', 'SupplierConfirmationNo',
    'AgentReferenceNo', 'ServiceName', 'CheckInDate', 'CheckOutDate',
    'SupplierBuyingAmount', 'AgentSellingAmount'
  ];
  editableFields.forEach(f => {
    if (data[f] !== undefined) {
      current[f] = f.indexOf('Amount') !== -1 ? Number(data[f]) || 0 : data[f];
    }
  });

  // Revised entry
  if (data.isRevision) {
    current.NewBuyingAmount = Number(data.NewBuyingAmount) || 0;
    current.NewSellingAmount = Number(data.NewSellingAmount) || 0;
    current.ConditionOfCalculation = data.ConditionOfCalculation || '';
    current.NatureOfProfit = data.NatureOfProfit || '';
    current.Remarks = data.Remarks || '';
    current.FinalConclusion = data.FinalConclusion || '';

    const result = calculateResult(
      Number(current.SupplierBuyingAmount) || 0,
      Number(current.AgentSellingAmount) || 0,
      current.NewBuyingAmount,
      current.NewSellingAmount,
      current.ConditionOfCalculation
    );
    current.ResultAmount = result.amount;
    current.ResultType = result.type;
    current.Status = 'Closed';
  }

  current.UpdatedAt = new Date();
  const newRow = ENTRY_HEADERS.map(h => current[h]);
  sheet.getRange(rowIndex, 1, 1, ENTRY_HEADERS.length).setValues([newRow]);
  return getEntryById(data.ID);
}

// Core business calculation — see README for the rules this implements.
function calculateResult(originalSupplier, originalAgent, newSupplier, newAgent, condition) {
  let base = 0;
  if (condition.indexOf('Agent') !== -1) {
    base = originalAgent - newAgent;
  } else {
    base = originalSupplier - newSupplier;
  }
  const type = condition.indexOf('Profit') !== -1 ? 'Profit' : 'Loss';
  return { amount: Math.abs(base), type: type };
}

function getEntryById(id) {
  return getAllEntries().find(e => e.ID === id);
}

function deleteEntry(id) {
  const sheet = getOrCreateSheet(ENTRY_SHEET);
  const rowIndex = findRowIndexById(sheet, id);
  if (rowIndex === -1) throw new Error('Entry not found (it may already be deleted)');
  sheet.deleteRow(rowIndex);
}

// ---------- Masters ----------
function getMasterList(name) {
  const sheet = getOrCreateSheet(name);
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  return sheet.getRange(1, 1, lastRow, 1).getValues().map(r => r[0]).filter(String);
}

function addMaster(list, value) {
  if (!value) throw new Error('Value required');
  const sheet = getOrCreateSheet(list);
  const existing = getMasterList(list);
  if (existing.indexOf(value) === -1) sheet.appendRow([value]);
}

function deleteMaster(list, value) {
  const sheet = getOrCreateSheet(list);
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) return;
  const values = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === value) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}
