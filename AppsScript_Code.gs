/**
 * TS Finance Dashboard — Automated Data Refresh
 * ================================================
 * Apps Script port of the "TS Finance Dashboard - Data Refresh" Colab notebook.
 * Runs under your own Google account (or Felita's/Tomekia's after handoff) —
 * no GCP project, no service account, no JSON key.
 *
 * ONE-TIME SETUP:
 * 1. Go to script.google.com -> New Project. Paste this whole file in as Code.gs.
 * 2. In the editor, click "Services" (+ icon in left sidebar) -> add "Drive API"
 *    (this lets us convert Felita's raw .xlsx into a temporary Google Sheet so
 *    we can read it — Apps Script can't open .xlsx files directly).
 * 3. The Chart of Accounts is restricted data, so its ID is NOT in this file
 *    (this file lives in a public repo). Set it via Project Settings (gear
 *    icon, left sidebar) -> Script Properties -> Add property:
 *      CHART_OF_ACCOUNTS_SHEET_ID = the ID from the sheet's URL
 *      CHART_OF_ACCOUNTS_TAB_NAME = exact tab name (only if not the first tab)
 *    It's read directly under your own Google account's existing access —
 *    never copied anywhere, never committed to the repo.
 * 4. Fill in CONFIG below with your real file/sheet IDs (defaults here are the
 *    known ones from the project).
 * 5. Run `refreshAll` once manually (Run button) — Google will ask you to
 *    authorize; click Allow. Check the Execution log for errors.
 * 6. Click the clock icon (Triggers) -> Add Trigger -> refreshAll -> Time-driven
 *    -> Minutes timer -> Every 5 minutes.
 *
 * OPTIONAL: also push the refreshed CSVs into the GitHub repo so the website
 * reads them same-origin instead of through the Drive+proxy path. See
 * pushCsvsToGitHub_() at the bottom — needs a GitHub Personal Access Token
 * stored in Script Properties (Project Settings -> Script Properties), never
 * pasted into this code.
 */

// ===================== CONFIG =====================
const CONFIG = {
  rawExcelId: '1bPBamW_Tui3_ty6FImaqANNTsAy2KYXZ',              // Felita's raw Excel
  contractsSheetId: '1F9gtr5UnzdwNHqXJ7kRoQW1jUSCd-e-lJiSjFpw_Iw8',
  tsPosSheetId: '1Gb9Lo0DHTHcsHNoKFOWdw-kcjQWH-SSDk0wFmChplXM',
  telecomSheetId: '1rRUBJP4EKW9MAeXyKwUNo8j98_0cMI-0StdVd-inE50',
  networkSheetId: '1MQryERd37hWGRiJwdZkimcAlbXQuaj4-kmI51eynIjc',
  jvSheetId: '1VpswBHv8Zx2g-xLY9K0lU9HoQsquuuXquIVu5VlekEk',
  fy27SheetId: '1NE9_S_7Y1csvLmb3ZlNI1JTBnbBlklIglVwXWBNmsLw',
  // Chart of Accounts is restricted data — its ID lives in Script Properties,
  // never in this file (this file is committed to a public repo). Set it via
  // Project Settings (gear icon) -> Script Properties -> Add property:
  //   CHART_OF_ACCOUNTS_SHEET_ID = <the sheet's ID from its URL>
  //   CHART_OF_ACCOUNTS_TAB_NAME = <exact tab name, if not the first tab>
  chartOfAccountsSheetId: null,
  chartOfAccountsTabName: null,
  outputFolderName: 'TS Finance Dashboard - Data Sources',
};

// ===================== ENTRY POINT =====================
function refreshAll() {
  const props = PropertiesService.getScriptProperties();
  CONFIG.chartOfAccountsSheetId = props.getProperty('CHART_OF_ACCOUNTS_SHEET_ID');
  CONFIG.chartOfAccountsTabName = props.getProperty('CHART_OF_ACCOUNTS_TAB_NAME');

  const folder = getOutputFolder_();
  const orgMap = loadChartOfAccounts_(); // Org code -> {dept, subdept, mbu}

  const results = {};
  const steps = [
    ['expenditures', () => processExpenditures_(folder, orgMap)],
    ['contracts', () => processContracts_(folder)],
    ['ts_pos', () => processTsPos_(folder, orgMap)],
    ['telecom', () => processTelecom_(folder)],
    ['network', () => processNetwork_(folder, orgMap)],
    ['jv', () => processJv_(folder)],
    ['fy27', () => processFy27_(folder, orgMap)],
  ];

  steps.forEach(([name, fn]) => {
    try {
      fn();
      results[name] = 'OK';
    } catch (e) {
      results[name] = 'FAILED: ' + e.message;
      Logger.log('Step "%s" failed: %s\n%s', name, e.message, e.stack);
    }
  });

  Logger.log(JSON.stringify(results, null, 2));
  // Uncomment once GitHub push is configured (see bottom of file):
  // pushCsvsToGitHub_(folder);
}

// ===================== SHARED HELPERS =====================
function getOutputFolder_() {
  const it = DriveApp.getFoldersByName(CONFIG.outputFolderName);
  if (!it.hasNext()) throw new Error('Output folder not found: ' + CONFIG.outputFolderName);
  return it.next();
}

function parseNum_(v) {
  if (v === null || v === undefined || v === '') return 0;
  const n = parseFloat(String(v).replace(/[$,\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function csvEscape_(v) {
  const s = (v === null || v === undefined) ? '' : String(v);
  return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
}

function rowsToCsv_(headers, rows) {
  const lines = [headers.join(',')];
  rows.forEach(row => {
    lines.push(headers.map(h => csvEscape_(row[h])).join(','));
  });
  return lines.join('\n');
}

function writeCsv_(folder, filename, headers, rows) {
  const csv = rowsToCsv_(headers, rows);
  const existing = folder.getFilesByName(filename);
  if (existing.hasNext()) {
    existing.next().setContent(csv);
  } else {
    folder.createFile(filename, csv, MimeType.CSV);
  }
  Logger.log('Wrote %s (%s rows)', filename, rows.length);
}

// Converts a raw .xlsx Drive file to a temporary Google Sheet so we can read
// it with SpreadsheetApp, then deletes the temp copy. Requires the "Drive API"
// advanced service enabled (Services -> + -> Drive API).
function readXlsxAsSheet_(fileId, sheetNameFilter) {
  const copy = Drive.Files.copy({ title: 'TEMP_xlsx_read_' + new Date().getTime() }, fileId);
  const ss = SpreadsheetApp.openById(copy.id);
  try {
    const out = {};
    ss.getSheets().forEach(sheet => {
      const name = sheet.getName();
      if (!sheetNameFilter || sheetNameFilter.includes(name)) {
        out[name] = sheet.getDataRange().getValues();
      }
    });
    return out;
  } finally {
    Drive.Files.remove(copy.id); // clean up the temp copy either way
  }
}

function openLiveSheet_(sheetId, tabName) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = tabName ? ss.getSheetByName(tabName) : ss.getSheets()[0];
  if (!sheet) return null;
  return sheet.getDataRange().getValues();
}

// ===================== CHART OF ACCOUNTS (Org -> Dept mapping) =====================
// Replaces the hand-typed dept_map / po_dept_map / network_dept_map / division_map
// dictionaries from the notebook with a lookup against VCU's actual chart of accounts.
function loadChartOfAccounts_() {
  if (!CONFIG.chartOfAccountsSheetId) {
    Logger.log('Chart of Accounts not configured (see Script Properties setup in the file header) — falling back to raw Org codes as department names.');
    return {};
  }
  const rows = openLiveSheet_(CONFIG.chartOfAccountsSheetId, CONFIG.chartOfAccountsTabName);
  const headers = rows[0];
  const orgCol = headers.indexOf('Org');       // code column (odd Org appears twice: code, then desc)
  const orgDescCol = orgCol + 1;
  const deptCol = headers.indexOf('Dept') + 1; // Dept desc (the code column is one before)
  const subdeptCol = headers.indexOf('Subdept') + 1;
  const mbuCol = headers.indexOf('Major Budget Unit') + 1;

  const map = {};
  for (let i = 1; i < rows.length; i++) {
    const org = String(rows[i][orgCol]).trim();
    if (!org) continue;
    map[org] = {
      orgDesc: rows[i][orgDescCol],
      dept: rows[i][deptCol],
      subdept: rows[i][subdeptCol],
      mbu: rows[i][mbuCol],
    };
  }
  return map;
}

// Known typos in the source Chart of Accounts, fixed for display only —
// the underlying VCU data is untouched, this just corrects what shows up
// on the dashboard.
const DEPT_NAME_FIXUPS = {
  'Network Infrastruture Maintenance': 'Network Infrastructure Maintenance',
};

function lookupDept_(orgMap, code, fallback) {
  const entry = orgMap[String(code).trim()];
  if (!entry) return fallback || String(code);
  return DEPT_NAME_FIXUPS[entry.subdept] || entry.subdept;
}

// The Chart of Accounts' "Dept" field (~9 values, VCU's own grouping) is NOT
// the same thing as this dashboard's 6 custom divisions (Central Services,
// Network & Infrastructure, Technical Support Services, Application Services,
// Information Security, Academic Technology) — that grouping was designed for
// this dashboard specifically and has no equivalent in VCU's chart of accounts.
// This maps the former onto the latter.
const COA_DEPT_TO_DIVISION = {
  'Telecommunications Services': 'Network & Infrastructure',
  'Network Services': 'Network & Infrastructure',
  'Technology Services': 'Central Services',
  'Central Services': 'Central Services',
  'Technology Support Services': 'Technical Support Services',
  'Administrative Systems': 'Application Services',
  'Application Services': 'Application Services',
  'Information Security': 'Information Security',
  'Academic Technologies': 'Academic Technology',
};

function lookupDivision_(orgMap, code, fallback) {
  const entry = orgMap[String(code).trim()];
  if (!entry) return fallback || 'Other';
  return COA_DEPT_TO_DIVISION[entry.dept] || fallback || 'Other';
}

// ===================== 1. EXPENDITURES (raw Excel) =====================
const FY_SHEET_MAP = {
  'TS E&G FY24': 'FY24',
  'TS E&G FY25': 'FY25',
  'TS E&G FY26': 'FY26',
};

const SKIP_CATEGORIES = new Set([
  'total budget', 'personal services', 'central services', 'total fy',
  'non-personal services', 'non-persoanl services', 'telecommunications',
  'network services', 'university computer center', 'administrative systems',
  'application services', 'academic technology', 'information security',
  'technical support services', 'ts strategic communications', 'nan',
]);

function groupCategory_(cat) {
  const c = String(cat).toLowerCase();
  const has = (...keys) => keys.some(k => c.indexOf(k) !== -1);
  if (has('salary', 'salaries', 'hourly', 'wages', 'faculty', 'it salaries', 'it/univ')) return 'Salaries';
  if (has('fringe')) return 'Fringes';
  if (has('bonus', 'on call', 'oncall', 'ot/', '/ot', 'vsdp', 'overtime')) return 'Bonus/OT/VSDP';
  if (has('software', 'maintenance', 'cloud', 'maint')) return 'Software & Maintenance';
  if (has('training', 'travel', 'membership')) return 'Training & Travel';
  if (has('supplies', 'postage', 'printing', 'shipping')) return 'Supplies & Printing';
  if (has('telecomm', 'telephone', 'network', 'tele ', 'wireless', 'windstream', 'segra', 'level3', 'comcast')) return 'Telecomm & Network';
  if (has('recovery', 'recoveries', 'revenue', 'internal charge', 'internal serv', 'sla')) return 'Recoveries & Revenue';
  if (has('equipment', 'hardware', 'computer', 'ramtech', 'laptop', 'computing', 'electronic')) return 'Equipment & Hardware';
  if (has('contractual', 'consulting', 'technical services', 'skilled services')) return 'Contractual & Services';
  return 'Other';
}

function getFiscalMonth_(date) {
  const m = date.getMonth() + 1; // 1-12
  return m >= 7 ? m - 6 : m + 6;
}
const MONTH_NAMES_ = { 1: 'Jul', 2: 'Aug', 3: 'Sep', 4: 'Oct', 5: 'Nov', 6: 'Dec', 7: 'Jan', 8: 'Feb', 9: 'Mar', 10: 'Apr', 11: 'May', 12: 'Jun' };

function processExpenditures_(folder, orgMap) {
  const sheets = readXlsxAsSheet_(CONFIG.rawExcelId, Object.keys(FY_SHEET_MAP));
  const allRows = [];

  Object.keys(sheets).forEach(sheetName => {
    const fy = FY_SHEET_MAP[sheetName];
    if (!fy) return;
    const data = sheets[sheetName];

    // Find columns that look like "a/o M/D/YY" month headers
    const monthCols = [];
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const val = String(data[r][c]);
        if (val.indexOf('a/o') !== -1 && /\d+\/\d+\/\d+/.test(val)) {
          if (!monthCols.some(m => m[0] === c)) monthCols.push([c, val.trim()]);
        }
      }
    }

    let currentDept = 'Unknown';
    for (let r = 0; r < data.length; r++) {
      for (const checkCol of [0, 1]) {
        const cellVal = String(data[r][checkCol]).trim();
        if (orgMap[cellVal]) { currentDept = cellVal; break; } // keep the Org code; resolve name at write-time
      }

      monthCols.forEach(([monthCol, monthLabel]) => {
        let category = null;
        for (const offset of [-3, -2, -1]) {
          const catCol = monthCol + offset;
          if (catCol < 0) continue;
          const potential = data[r][catCol];
          if (potential !== '' && potential !== null && String(potential).trim() !== 'nan') {
            if (isNaN(parseFloat(potential))) { category = String(potential).trim(); break; }
          }
        }
        if (!category || SKIP_CATEGORIES.has(category.toLowerCase())) return;

        const expenditure = parseFloat(data[r][monthCol]);
        if (isNaN(expenditure) || expenditure === 0) return;
        const permBudget = monthCol >= 2 ? parseFloat(data[r][monthCol - 2]) || 0 : 0;
        const currentBudget = monthCol >= 1 ? parseFloat(data[r][monthCol - 1]) || 0 : 0;
        const balance = parseFloat(data[r][monthCol + 3]) || 0;

        const dateMatch = monthLabel.replace('a/o', '').trim();
        const date = new Date(dateMatch);
        if (isNaN(date)) return;

        allRows.push({
          FY: fy,
          Department: lookupDept_(orgMap, currentDept, currentDept),
          Division: lookupDivision_(orgMap, currentDept, 'Other'),
          Month: monthLabel,
          Category: category,
          Category_Group: groupCategory_(category),
          Perm_Budget: permBudget,
          Current_Budget: currentBudget,
          Expenditure: expenditure,
          Balance: balance,
          Date: Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          Fiscal_Month: getFiscalMonth_(date),
          Month_Name: MONTH_NAMES_[getFiscalMonth_(date)],
        });
      });
    }
  });

  const headers = ['FY', 'Department', 'Division', 'Month', 'Category', 'Category_Group', 'Perm_Budget', 'Current_Budget', 'Expenditure', 'Balance', 'Date', 'Fiscal_Month', 'Month_Name'];
  writeCsv_(folder, 'expenditures_monthly.csv', headers, allRows);

  // Last-month-only view per FY/Department, for budget comparison
  const lastDateByKey = {};
  allRows.forEach(r => {
    const key = r.FY + '|' + r.Department;
    if (!lastDateByKey[key] || r.Date > lastDateByKey[key]) lastDateByKey[key] = r.Date;
  });
  const lastRows = allRows.filter(r => r.Date === lastDateByKey[r.FY + '|' + r.Department]);
  writeCsv_(folder, 'expenditures_clean.csv', headers, lastRows);
}

// ===================== 2. CONTRACTS =====================
const CONTRACT_DEPT_SHEETS = ['InfoSec/EndPoint', 'Application Svcs', 'TS Administration',
  'Technology Support Svcs', 'Academic Technologies', 'Network/UCC',
  'Administrative Systems', 'Central Maintenance'];

function getContractStatus_(renewalDate) {
  if (!renewalDate) return 'Unknown';
  const days = Math.floor((renewalDate - new Date()) / 86400000);
  if (days < 0) return 'Expired';
  if (days <= 90) return 'Expiring Soon';
  return 'Active';
}

function processContracts_(folder) {
  const ss = SpreadsheetApp.openById(CONFIG.contractsSheetId);
  const junk = new Set(['NOTES:', 'Total Cost:', 'MARIA', 'Facility Support', '']);
  const rows = [];

  CONTRACT_DEPT_SHEETS.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const vendor = String(row[0] || '').trim();
      if (!vendor || junk.has(vendor)) continue;
      const cost = parseNum_(row[3]);
      if (cost <= 0) continue;
      let renewalDate = null;
      if (row[4]) {
        const d = new Date(row[4]);
        if (!isNaN(d)) renewalDate = d;
      }
      rows.push({
        Department: sheetName,
        Vendor: vendor,
        Product: row[1] || '',
        Annual_Cost: cost,
        Renewal_Date: renewalDate ? Utilities.formatDate(renewalDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
        Funding_Source: row[5] || '',
        Status: getContractStatus_(renewalDate),
        Days_Until_Expiry: renewalDate ? Math.floor((renewalDate - new Date()) / 86400000) : '',
      });
    }
  });

  writeCsv_(folder, 'contracts_clean.csv',
    ['Department', 'Vendor', 'Product', 'Annual_Cost', 'Renewal_Date', 'Funding_Source', 'Status', 'Days_Until_Expiry'],
    rows);
}

// ===================== 3. TS PURCHASE ORDERS =====================
function processTsPos_(folder, orgMap) {
  const ss = SpreadsheetApp.openById(CONFIG.tsPosSheetId);
  const rows = [];

  ss.getSheets().forEach(sheet => {
    const tabName = sheet.getName().trim();
    const dept = lookupDept_(orgMap, tabName.replace(/^1-|^2-|^3-|^R-/, ''), tabName);
    const data = sheet.getDataRange().getValues();
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].includes('Date') && data[i].includes('Vendor')) { headerRow = i; break; }
    }
    if (headerRow === -1) return;

    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const amount = parseNum_(row[4]);
      if (amount === 0) continue;
      const vendor = String(row[3] || '').trim();
      if (!vendor) continue;
      rows.push({
        Department: dept,
        Date: row[0],
        PR_Number: row[1] || '',
        PO_Number: row[2] || '',
        Vendor: vendor,
        Amount: amount,
        Amount_Received: parseNum_(row[5]),
        Ending_Balance: parseNum_(row[6]),
        Receiver_Date: row[7] || '',
        Budget_Code: row[8] || '',
        Account_Code: row[9] || '',
        Comment: row[10] || '',
      });
    }
  });

  writeCsv_(folder, 'purchase_orders_clean.csv',
    ['Department', 'Date', 'PR_Number', 'PO_Number', 'Vendor', 'Amount', 'Amount_Received', 'Ending_Balance', 'Receiver_Date', 'Budget_Code', 'Account_Code', 'Comment'],
    rows);
}

// ===================== 4. TELECOM POs (column layout varies by FY) =====================
const TELECOM_SHEET_CONFIG = {
  FY2027: { date: 0, pr: 1, po: 2, contract: 3, vendor: 4, amount: 5, received: 6, balance: 7, receiver: 8, budget: 9 },
  FY2026: { date: 0, pr: 1, po: null, contract: 2, vendor: 3, amount: 4, received: 5, balance: 6, receiver: 7, budget: 8 },
  FY2025: { date: 0, pr: 1, po: null, contract: 2, vendor: 3, amount: 4, received: 5, balance: 6, receiver: 7, budget: 8 },
  FY2024: { date: 0, pr: 1, po: null, contract: 2, vendor: 3, amount: 4, received: 5, balance: 6, receiver: 7, budget: 8 },
};

function processTelecom_(folder) {
  const ss = SpreadsheetApp.openById(CONFIG.telecomSheetId);
  const rows = [];

  Object.keys(TELECOM_SHEET_CONFIG).forEach(fy => {
    const cols = TELECOM_SHEET_CONFIG[fy];
    const sheet = ss.getSheetByName(fy);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].some(v => String(v).toUpperCase() === 'DATE')) { headerRow = i; break; }
    }
    if (headerRow === -1) return;

    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const amount = parseNum_(row[cols.amount]);
      if (amount === 0) continue;
      const vendor = String(row[cols.vendor] || '').trim();
      if (!vendor) continue;
      rows.push({
        FY: fy,
        Department: 'Telecommunications',
        Date: row[cols.date],
        PR_Number: row[cols.pr] || '',
        Vendor: vendor,
        Amount: amount,
        Amount_Received: parseNum_(row[cols.received]),
        Ending_Balance: parseNum_(row[cols.balance]),
        Receiver_Date: row[cols.receiver] || '',
        Budget_Code: row[cols.budget] || '',
      });
    }
  });

  writeCsv_(folder, 'telecom_pos_clean.csv',
    ['FY', 'Department', 'Date', 'PR_Number', 'Vendor', 'Amount', 'Amount_Received', 'Ending_Balance', 'Receiver_Date', 'Budget_Code'],
    rows);
}

// ===================== 5. NETWORK POs =====================
function processNetwork_(folder, orgMap) {
  const ss = SpreadsheetApp.openById(CONFIG.networkSheetId);
  const rows = [];

  ss.getSheets().forEach(sheet => {
    const tabName = sheet.getName().trim();
    const dept = lookupDept_(orgMap, tabName, tabName);
    const data = sheet.getDataRange().getValues();
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].some(v => String(v).toLowerCase() === 'date')) { headerRow = i; break; }
    }
    if (headerRow === -1) return;

    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const amount = parseNum_(row[5]);
      if (amount === 0) continue;
      const vendor = String(row[4] || '').trim();
      if (!vendor) continue;
      rows.push({
        Department: dept,
        Date: row[0],
        Job_Number: row[1] || '',
        PR_Number: row[2] || '',
        PO_Number: row[3] || '',
        Vendor: vendor,
        Amount: amount,
        Amount_Received: parseNum_(row[6]),
        Ending_Balance: parseNum_(row[7]),
        Receiver_Date: row[8] || '',
      });
    }
  });

  writeCsv_(folder, 'network_pos_clean.csv',
    ['Department', 'Date', 'Job_Number', 'PR_Number', 'PO_Number', 'Vendor', 'Amount', 'Amount_Received', 'Ending_Balance', 'Receiver_Date'],
    rows);
}

// ===================== 6. JV LOG =====================
function processJv_(folder) {
  const ss = SpreadsheetApp.openById(CONFIG.jvSheetId);
  const rows = [];

  ['FY2027', 'FY2026', 'FY2025', 'FY2024'].forEach(fy => {
    const sheet = ss.getSheetByName(fy);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let headerRow = -1;
    for (let i = 0; i < data.length; i++) {
      if (data[i].some(v => String(v).toLowerCase() === 'date')) { headerRow = i; break; }
    }
    if (headerRow === -1) return;

    for (let i = headerRow + 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      const amount = parseNum_(row[3]);
      if (amount === 0) continue;
      const dept = String(row[2] || '').trim();
      if (!dept) continue;
      rows.push({
        FY: fy,
        Date: row[0],
        HD_Number: row[1] || '',
        Department: dept,
        Amount: amount,
        JV_Number: row[4] || '',
        Budget_Code: row[5] || '',
        Comment: row[6] || '',
      });
    }
  });

  writeCsv_(folder, 'jv_log_clean.csv',
    ['FY', 'Date', 'HD_Number', 'Department', 'Amount', 'JV_Number', 'Budget_Code', 'Comment'],
    rows);
}

// ===================== 7. FY27 BUDGET (+ projection) =====================
function processFy27_(folder, orgMap) {
  const ss = SpreadsheetApp.openById(CONFIG.fy27SheetId);
  const sheet = ss.getSheets()[0];
  const data = sheet.getDataRange().getValues();
  const rows = [];

  for (let i = 7; i < data.length; i++) { // matches notebook's data[7:44] slice
    const row = data[i];
    if (!row[0] || row[0] === 'Total') continue;
    const budget = parseNum_(row[1]);
    const expenses = parseNum_(row[2]);
    const commitments = parseNum_(row[3]);
    const difference = parseNum_(row[4]);
    const dept = String(row[5] || '').trim();
    if (!dept) continue;
    rows.push({
      Index: String(row[0]).trim(),
      Department: dept,
      FY: 'FY27',
      Budget: budget,
      Expenses: expenses,
      Commitments: commitments,
      Difference: difference,
      As_Of: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      Utilization_Pct: budget > 0 ? Math.round((expenses / budget) * 1000) / 10 : 0,
    });
  }

  writeCsv_(folder, 'fy27_budget.csv',
    ['Index', 'Department', 'FY', 'Budget', 'Expenses', 'Commitments', 'Difference', 'As_Of', 'Utilization_Pct'],
    rows);
}

// ===================== OPTIONAL: push CSVs into the GitHub repo =====================
// Lets the website read same-origin files instead of the Drive+CORS-proxy path.
// One-time setup: Project Settings (gear icon) -> Script Properties -> Add:
//   GITHUB_TOKEN = a GitHub Personal Access Token with "repo" scope
function pushCsvsToGitHub_(folder) {
  const token = PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
  if (!token) { Logger.log('GITHUB_TOKEN not set — skipping GitHub push.'); return; }

  const owner = 'Tejesh18';
  const repo = 'vcu-finance-dashboard';
  const files = folder.getFilesByType(MimeType.CSV);

  while (files.hasNext()) {
    const file = files.next();
    const path = 'data/' + file.getName();
    const content = Utilities.base64Encode(file.getBlob().getBytes());

    const apiUrl = `https://api.github.com/repos/${owner}/${repo}/contents/${path}`;
    let sha = null;
    try {
      const getResp = UrlFetchApp.fetch(apiUrl, {
        headers: { Authorization: 'token ' + token },
        muteHttpExceptions: true,
      });
      if (getResp.getResponseCode() === 200) sha = JSON.parse(getResp.getContentText()).sha;
    } catch (e) { /* file doesn't exist yet — that's fine */ }

    UrlFetchApp.fetch(apiUrl, {
      method: 'put',
      contentType: 'application/json',
      headers: { Authorization: 'token ' + token },
      payload: JSON.stringify({
        message: 'Automated data refresh: ' + file.getName(),
        content: content,
        sha: sha || undefined,
      }),
      muteHttpExceptions: true,
    });
    Logger.log('Pushed %s to GitHub', file.getName());
  }
}
