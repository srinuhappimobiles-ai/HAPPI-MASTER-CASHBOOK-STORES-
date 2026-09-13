import React, {
  useEffect,
  useRef,
  useState,
} from 'react';

import { DEFAULT_BRANCHES } from '../config/branches';
import { supabase } from '../lib/supabase';
import * as XLSX from 'xlsx-js-style';

const STORAGE_KEY = 'happi_master_cashbook_v1';
const FINANCE_REPORT_STORAGE_KEY = 'happi_finance_report_v1';
const OPENING_UNDO_KEY = 'opening';
const DEPOSIT_UNDO_KEY = 'deposit';
const ADDINS_UNDO_KEY = 'addings';
const DENOMINATION_UNDO_KEY = 'denomination';

const FINANCE_CC_EMAILS = [
  'dfm@happimobiles.com',
  'jfm@happimobiles.com',
  'agmfinance@happimobiles.com',
  'financesupport@happimobiles.com',
  'financesupport1@happimobiles.com',
  'financesupport3@happimobiles.com',
  'affordability@happimobiles.com',
];

const SR_CC_EMAILS = [
  'doatg@happimobiles.com',
  'doaap@happimobiles.com',
  'dfm@happimobiles.com',
  'affordability@happimobiles.com',
];

const SR_REPORT_STORAGE_KEY = 'happi_sr_report_v1';
const LOW_CASH_STORAGE_KEY = 'happi_low_cash_report_v1';

// Cloud sync runs only once every 10 minutes.
// LocalStorage is updated immediately, so typing/editing never waits for Supabase.
const AUTO_SYNC_INTERVAL = 10 * 60 * 1000;

const HEADERS = [
  'SL.No.',
  'CODE',
  'BRANCH',
  'OPENING BALANCE',
  'DEPOSIT',
  'DENOMINATION',
  'AddinGS',
  'PENDING APPRVLS',
  'FINANCE AMNT',
  'SR',
  'SWEEPER SALARY',
  'EDITS',
  'APX SHORTAGE',
  "(KSP)'Sir's Approvals",
  'CLOSING BALANCE',
  'REMARKS',
];

// All master numeric cells use the same Excel-like edit/navigation
// behavior: F2/double-click, Backspace one character at a time,
// Delete clears the whole cell, arrow-key navigation, Enter/blur,
// Escape, and direct entry for empty cells.
const NUMBER_FIELDS = [
  'opening',
  'deposit',
  'denomination',
  'addings',
  'pendingApprovals',
  'finance',
  'sr',
  'sweeperSalary',
  'edits',
  'apxShortage',
  'kspApprovals',
];

/* -----------------------------------------
   LOCAL STORAGE
----------------------------------------- */

function getSavedData() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);

    if (!saved) {
      return {};
    }

    const parsed = JSON.parse(saved);

    if (!Array.isArray(parsed)) {
      return {};
    }

    const map = {};

    for (const row of parsed) {
      if (row && row.code) {
        map[row.code] = row;
      }
    }

    return map;
  } catch (error) {
    console.error('Cashbook local load error:', error);
    return {};
  }
}

/* -----------------------------------------
   CREATE ROWS
----------------------------------------- */

function createRows(savedMap = {}) {
  return DEFAULT_BRANCHES.map((branch, index) => {
    const saved = savedMap[branch.CODE];

    return {
      slNo: index + 1,
      code: branch.CODE,
      branch: branch.BRANCH,

      opening: saved?.opening ?? '',
      deposit: saved?.deposit ?? '',
      denomination: saved?.denomination ?? '',
      addings: saved?.addings ?? '',
      pendingApprovals: saved?.pendingApprovals ?? '',
      finance: saved?.finance ?? '',
      sr: saved?.sr ?? '',
      sweeperSalary: saved?.sweeperSalary ?? '',
      edits: saved?.edits ?? '',
      apxShortage: saved?.apxShortage ?? '',
      kspApprovals: saved?.kspApprovals ?? '',

      remarks: saved?.remarks ?? '',
    };
  });
}

/* -----------------------------------------
   DATABASE ROW → WEBSITE ROW
----------------------------------------- */

function databaseRowToWebsiteRow(dbRow, index, fallbackBranch) {
  return {
    slNo: index + 1,

    code: dbRow?.code || fallbackBranch.CODE,

    branch: dbRow?.branch || fallbackBranch.BRANCH,

    opening: dbRow?.opening_balance ?? '',
    deposit: dbRow?.deposit ?? '',
    denomination: dbRow?.denomination ?? '',
    addings: dbRow?.addings ?? '',
    pendingApprovals: dbRow?.pending_apprvls ?? '',
    finance: dbRow?.finance_amnt ?? '',
    sr: dbRow?.sr ?? '',
    sweeperSalary: dbRow?.sweeper_salary ?? '',
    edits: dbRow?.edits ?? '',
    apxShortage: dbRow?.apx_shortage ?? '',
    kspApprovals: dbRow?.ksp_approvals ?? '',

    remarks: dbRow?.remarks ?? '',

    // Keep Supabase ID internally.
    databaseId: dbRow?.id ?? null,
  };
}

/* -----------------------------------------
   FORMULA / AMOUNT CALCULATOR

   Supports:
   600
   -425
   =600+6000
   600+6000
   =10000+2700+500
   100-50
   (1000+500)-200

   A leading "=" is treated like an Excel-style
   formula marker. Only arithmetic characters are
   accepted; no JavaScript evaluation is used.
----------------------------------------- */

function calculateAmount(value) {
  if (value === null || value === undefined) {
    return 0;
  }

  let expression = String(value).trim();

  if (expression === '') {
    return 0;
  }

  // Support Excel-style formulas such as =600+6000
  if (expression.startsWith('=')) {
    expression = expression.slice(1).trim();
  }

  if (expression === '') {
    return 0;
  }

  // Plain number
  if (/^-?\d+(?:\.\d+)?$/.test(expression)) {
    return Number(expression);
  }

  // Remove spaces inside expressions
  expression = expression.replace(/\s+/g, '');

  // Safe arithmetic characters only
  if (!/^[0-9.+\-*/()]+$/.test(expression)) {
    return 0;
  }

  let position = 0;

  function parseExpression() {
    let value = parseTerm();

    while (position < expression.length) {
      const operator = expression[position];

      if (operator !== '+' && operator !== '-') {
        break;
      }

      position++;

      const nextValue = parseTerm();

      if (operator === '+') {
        value += nextValue;
      } else {
        value -= nextValue;
      }
    }

    return value;
  }

  function parseTerm() {
    let value = parseFactor();

    while (position < expression.length) {
      const operator = expression[position];

      if (operator !== '*' && operator !== '/') {
        break;
      }

      position++;

      const nextValue = parseFactor();

      if (operator === '*') {
        value *= nextValue;
      } else {
        if (nextValue === 0) {
          throw new Error('Division by zero');
        }

        value /= nextValue;
      }
    }

    return value;
  }

  function parseFactor() {
    // Unary + / -
    if (
      expression[position] === '+' ||
      expression[position] === '-'
    ) {
      const operator = expression[position];
      position++;

      const value = parseFactor();

      return operator === '-' ? -value : value;
    }

    // Parentheses
    if (expression[position] === '(') {
      position++;

      const value = parseExpression();

      if (expression[position] !== ')') {
        throw new Error('Missing closing parenthesis');
      }

      position++;

      return value;
    }

    // Number
    const start = position;

    while (
      position < expression.length &&
      /[0-9.]/.test(expression[position])
    ) {
      position++;
    }

    if (start === position) {
      throw new Error('Invalid number');
    }

    const number = Number(
      expression.slice(start, position)
    );

    if (!Number.isFinite(number)) {
      throw new Error('Invalid number');
    }

    return number;
  }

  try {
    const result = parseExpression();

    // Reject incomplete/invalid expressions such as 100+
    if (position !== expression.length) {
      return 0;
    }

    return Number.isFinite(result) ? result : 0;
  } catch {
    return 0;
  }
}


/* -----------------------------------------
   CLOSING BALANCE
----------------------------------------- */

function calculateClosing(row) {
  return (
    calculateAmount(row.opening) -
    calculateAmount(row.deposit) -
    calculateAmount(row.denomination) -
    calculateAmount(row.addings) -
    calculateAmount(row.pendingApprovals) -
    calculateAmount(row.finance) -
    calculateAmount(row.sr) -
    calculateAmount(row.sweeperSalary) -
    calculateAmount(row.edits) -
    calculateAmount(row.apxShortage) -
    calculateAmount(row.kspApprovals)
  );
}


/* -----------------------------------------
   FORMAT AMOUNT
----------------------------------------- */

function formatAmount(value) {
  if (!Number.isFinite(value)) {
    return '';
  }

  return value.toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

/* -----------------------------------------
   DISPLAY VALUE

   Stored values such as =600+6000 are shown
   as 6600 inside the input cell.
----------------------------------------- */

function getDisplayInputValue(value) {
  if (value === null || value === undefined) {
    return '';
  }

  const text = String(value).trim();

  if (text === '') {
    return '';
  }

  if (
    text.startsWith('=') ||
    /[+\-*/()]/.test(text)
  ) {
    const result = calculateAmount(text);

    if (Number.isFinite(result)) {
      return String(result);
    }
  }

  return text;
}

function getStoredInputValue(input) {
  return input.dataset.rawValue !== undefined
    ? input.dataset.rawValue
    : input.value;
}

/* -----------------------------------------
   FEATURE HELPERS
----------------------------------------- */

function normalizeBranchName(value) {
  return String(value ?? '')
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function roundRupee(value) {
  const number = Number(value);

  if (!Number.isFinite(number)) {
    return null;
  }

  const sign = number < 0 ? -1 : 1;
  const absolute = Math.abs(number);

  // Required rule: below .500 stays down; .500 and above goes up.
  return sign * Math.floor(absolute + 0.5);
}

function parseAmountNumber(value) {
  if (value === null || value === undefined) {
    return null;
  }

  const text = String(value)
    .replace(/₹/g, '')
    .replace(/,/g, '')
    .trim();

  if (!text) return null;

  // Excel opening-balance values can arrive like:
  // "6728.404 Cr."
  // "5828.940 Dr."
  // "12,437.200 Cr."
  //
  // Extract the first actual numeric portion instead of trying
  // to strip all non-numeric characters, because the trailing
  // period in "Cr." would otherwise produce "6728.404.".
  const match = text.match(
    /[+-]?\s*\d+(?:\.\d+)?/
  );

  if (!match) return null;

  const number = Number(
    match[0].replace(/\s+/g, '')
  );

  return Number.isFinite(number)
    ? number
    : null;
}

function parseOpeningBalance(value) {
  if (value === null || value === undefined || String(value).trim() === '') {
    return null;
  }

  const text = String(value).trim();
  const upper = text.toUpperCase();
  const amount = parseAmountNumber(text);

  if (amount === null) {
    return null;
  }

  const rounded = roundRupee(Math.abs(amount));

  if (rounded === null) {
    return null;
  }

  if (upper.includes('CR')) {
    return -Math.abs(rounded);
  }

  return Math.abs(rounded);
}

function parseVoucherValue(value) {
  const amount = parseAmountNumber(value);
  if (amount === null) return null;
  return roundRupee(Math.abs(amount));
}

function getLocalDateString(date = new Date()) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function calculateAgeing(todayString, billDateString) {
  if (!todayString || !billDateString) {
    return '';
  }

  const today = new Date(`${todayString}T00:00:00`);
  const billDate = new Date(`${billDateString}T00:00:00`);

  if (
    Number.isNaN(today.getTime()) ||
    Number.isNaN(billDate.getTime())
  ) {
    return '';
  }

  const diff = Math.floor(
    (today.getTime() - billDate.getTime()) /
      (24 * 60 * 60 * 1000)
  );

  return Math.max(0, diff);
}

function getFinanceStorageMap() {
  try {
    const saved = localStorage.getItem(
      FINANCE_REPORT_STORAGE_KEY
    );

    if (!saved) return {};

    const parsed = JSON.parse(saved);
    return parsed && typeof parsed === 'object'
      ? parsed
      : {};
  } catch (error) {
    console.error(
      'Finance report local load error:',
      error
    );
    return {};
  }
}

function saveFinanceStorageMap(map) {
  try {
    localStorage.setItem(
      FINANCE_REPORT_STORAGE_KEY,
      JSON.stringify(map)
    );
  } catch (error) {
    console.error(
      'Finance report local save error:',
      error
    );
  }
}

function branchToEmail(branchName) {
  const localPart = normalizeBranchName(branchName)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, '');

  return `${localPart}@happimobiles.com`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

/* -----------------------------------------
   WEBSITE ROW → DATABASE ROW
----------------------------------------- */

function websiteRowToDatabaseRow(row, databaseId = null) {
  const payload = {
    code: row.code,
    branch: row.branch,

    opening_balance: row.opening,
    deposit: row.deposit,
    denomination: row.denomination,
    addings: row.addings,
    pending_apprvls: row.pendingApprovals,
    finance_amnt: row.finance,
    sr: row.sr,
    sweeper_salary: row.sweeperSalary,
    edits: row.edits,
    apx_shortage: row.apxShortage,
    ksp_approvals: row.kspApprovals,

    remarks: row.remarks,

    updated_at: new Date().toISOString(),
  };

  if (databaseId !== null && databaseId !== undefined) {
    payload.id = databaseId;
  }

  return payload;
}

const featureButtonStyle = (background) => ({
  border: '1px solid rgba(0,0,0,0.10)',
  borderRadius: '5px',
  padding: '6px 10px',
  background,
  color: '#ffffff',
  fontSize: '11px',
  fontWeight: 700,
  cursor: 'pointer',
  whiteSpace: 'nowrap',
  boxShadow: '0 1px 2px rgba(0,0,0,0.12)',
});

const modalOverlayStyle = {
  position: 'fixed',
  inset: 0,
  zIndex: 2000,
  background: 'rgba(24, 36, 48, 0.42)',
  display: 'flex',
  alignItems: 'flex-start',
  justifyContent: 'center',
  padding: '70px 16px 16px',
};

const financeModalStyle = {
  width: 'min(1120px, 98vw)',
  maxHeight: '82vh',
  background: '#ffffff',
  borderRadius: '6px',
  border: '1px solid #b9c5cf',
  boxShadow: '0 10px 35px rgba(0,0,0,0.22)',
  overflow: 'hidden',
  display: 'flex',
  flexDirection: 'column',
};

const errorModalStyle = {
  width: 'min(680px, 96vw)',
  maxHeight: '80vh',
  background: '#ffffff',
  borderRadius: '6px',
  border: '1px solid #d9b0b0',
  boxShadow: '0 10px 35px rgba(0,0,0,0.22)',
  overflow: 'hidden',
  display: 'flex',
  flexDirection: 'column',
};

const modalHeaderStyle = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  minHeight: '42px',
  padding: '0 9px 0 12px',
  background: 'linear-gradient(90deg,#ffffff 0%,#f4f8fc 100%)',
  borderBottom: '1px solid #d5dee7',
};

const closeButtonStyle = {
  border: 0,
  background: 'transparent',
  color: '#7b8a98',
  fontSize: '24px',
  lineHeight: 1,
  cursor: 'pointer',
  padding: '2px 5px',
};

const financeTableWrapStyle = {
  overflow: 'auto',
  padding: '8px 10px 0',
};

const financeTableStyle = {
  width: '100%',
  minWidth: '920px',
  borderCollapse: 'collapse',
  fontSize: '11px',
  color: '#263746',
};

const financeThStyle = {
  position: 'sticky',
  top: 0,
  zIndex: 2,
  padding: '6px 7px',
  border: '1px solid #bdc9d3',
  background: '#f5c928',
  color: '#17212b',
  fontWeight: 800,
  textAlign: 'left',
  whiteSpace: 'nowrap',
};

const financeTdStyle = {
  padding: '4px 6px',
  border: '1px solid #cbd5df',
  background: '#ffffff',
};

const financeTdCenterStyle = {
  ...financeTdStyle,
  textAlign: 'center',
  whiteSpace: 'nowrap',
};

const financeAmountTdStyle = {
  ...financeTdStyle,
  textAlign: 'right',
  fontWeight: 700,
  whiteSpace: 'nowrap',
};

const financeInputStyle = {
  width: '100%',
  minWidth: '105px',
  height: '25px',
  padding: '3px 6px',
  border: '1px solid #b7c4cf',
  borderRadius: '3px',
  outline: 'none',
  fontSize: '11px',
  background: '#ffffff',
};

const financeDateInputStyle = {
  height: '25px',
  padding: '2px 4px',
  border: '1px solid #b7c4cf',
  borderRadius: '3px',
  fontSize: '11px',
  background: '#ffffff',
};

const financeFooterStyle = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  gap: '10px',
  padding: '8px 10px',
  borderTop: '1px solid #d9e1e8',
  background: '#f8fafc',
};

const financeActionButtonStyle = (background) => ({
  border: '1px solid rgba(0,0,0,0.10)',
  borderRadius: '4px',
  padding: '6px 10px',
  background,
  color: '#ffffff',
  fontSize: '11px',
  fontWeight: 700,
  cursor: 'pointer',
  whiteSpace: 'nowrap',
});

const financeCloseActionStyle = {
  border: '1px solid #9aa8b5',
  borderRadius: '4px',
  padding: '6px 10px',
  background: '#ffffff',
  color: '#405261',
  fontSize: '11px',
  fontWeight: 700,
  cursor: 'pointer',
};

const errorListStyle = {
  overflow: 'auto',
  padding: '8px 10px',
  maxHeight: '52vh',
  background: '#fffafa',
};

const errorRowStyle = {
  padding: '7px 8px',
  marginBottom: '5px',
  border: '1px solid #f0c4c4',
  borderRadius: '4px',
  background: '#fff4f4',
  fontSize: '11px',
  color: '#4a2020',
};

const simpleListModalStyle = {
  width: 'min(430px, 94vw)',
  maxHeight: '78vh',
  background: '#ffffff',
  borderRadius: '6px',
  border: '1px solid #b9c5cf',
  boxShadow: '0 10px 35px rgba(0,0,0,0.22)',
  overflow: 'hidden',
  display: 'flex',
  flexDirection: 'column',
};

const twoPaneModalStyle = {
  width: 'min(900px, 96vw)',
  maxHeight: '78vh',
  background: '#ffffff',
  borderRadius: '6px',
  border: '1px solid #b9c5cf',
  boxShadow: '0 10px 35px rgba(0,0,0,0.22)',
  overflow: 'hidden',
  display: 'flex',
  flexDirection: 'column',
};

const reportListStyle = {
  overflow: 'auto',
  maxHeight: '62vh',
};

const reportListRowStyle = {
  display: 'grid',
  gridTemplateColumns: '74px 1fr 92px',
  gap: '7px',
  alignItems: 'center',
  minHeight: '29px',
  padding: '4px 9px',
  borderBottom: '1px solid #e3e9ef',
  fontSize: '11px',
};

const reportCodeStyle = {
  fontWeight: 800,
  color: '#034b8f',
};

const reportBranchStyle = {
  color: '#3d5366',
  whiteSpace: 'nowrap',
  overflow: 'hidden',
  textOverflow: 'ellipsis',
};

const reportAmountStyle = {
  textAlign: 'right',
  fontWeight: 800,
  color: '#0066a6',
  whiteSpace: 'nowrap',
};

const twoPaneGridStyle = {
  display: 'grid',
  gridTemplateColumns: '1fr 1fr',
  minHeight: 0,
};

const paneStyle = {
  minWidth: 0,
  borderRight: '1px solid #d9e1e8',
};

const paneLastStyle = {
  minWidth: 0,
};

const paneTitleStyle = {
  display: 'flex',
  alignItems: 'center',
  justifyContent: 'space-between',
  padding: '8px 9px',
  borderBottom: '1px solid #d9e1e8',
  background: '#f6f9fc',
  fontSize: '12px',
  fontWeight: 800,
  color: '#253746',
};

const lowCashTableStyle = {
  width: '100%',
  minWidth: '640px',
  borderCollapse: 'collapse',
  fontSize: '11px',
};

const lowCashInputStyle = {
  width: '100%',
  minWidth: '125px',
  height: '26px',
  padding: '3px 6px',
  border: '1px solid #b7c4cf',
  borderRadius: '3px',
  fontSize: '11px',
  background: '#ffffff',
};

const lowCashDefaultStyle = {
  ...lowCashInputStyle,
  fontWeight: 700,
};

const undoButtonStyle = {
  border: '1px solid #64748b',
  borderRadius: '5px',
  padding: '6px 9px',
  background: '#ffffff',
  color: '#334155',
  fontSize: '11px',
  fontWeight: 700,
  cursor: 'pointer',
  whiteSpace: 'nowrap',
};

/* -----------------------------------------
   MAIN COMPONENT
----------------------------------------- */

export default function CashbookGrid() {
  const rowRefs = useRef([]);

  const databaseRowsRef = useRef({});

  // Prevent duplicate cloud sync requests.
  const autoSyncInFlightRef = useRef(false);

  // Only changed branches are sent to Supabase at the 10-minute sync.
  const dirtyCodesRef = useRef(new Set());

  // Tracks which amount cells are currently in F2/double-click edit mode.
  const editingCellRef = useRef(null);

  const [rows, setRows] = useState([]);
  const [cashierFilter, setCashierFilter] = useState('all');

  const [loading, setLoading] = useState(true);

  const [saveStatus, setSaveStatus] = useState(
    'Connecting...'
  );

  const rowsRef = useRef([]);
  const [financeReportOpen, setFinanceReportOpen] = useState(false);
  const [financeReportRows, setFinanceReportRows] = useState([]);
  const [financeToday, setFinanceToday] = useState(
    getLocalDateString()
  );

  const [srReportOpen, setSrReportOpen] = useState(false);
  const [srReportRows, setSrReportRows] = useState([]);
  const [srToday, setSrToday] = useState(getLocalDateString());

  const [editsOpen, setEditsOpen] = useState(false);
  const [pendingApprovalsOpen, setPendingApprovalsOpen] = useState(false);
  const [pendingStatusOpen, setPendingStatusOpen] = useState(false);
  const [lowCashOpen, setLowCashOpen] = useState(false);
  const [lowCashRows, setLowCashRows] = useState([]);
  const [uploadErrors, setUploadErrors] = useState([]);
  const [uploadErrorTitle, setUploadErrorTitle] = useState('Upload Errors');
  const [uploadErrorOpen, setUploadErrorOpen] = useState(false);
  const [lastClearField, setLastClearField] = useState(null);
  const [clipboardNotice, setClipboardNotice] = useState('');

  const openingFileInputRef = useRef(null);
  const addinsFileInputRef = useRef(null);

  const undoSnapshotsRef = useRef({
    [OPENING_UNDO_KEY]: null,
    [DEPOSIT_UNDO_KEY]: null,
    [ADDINS_UNDO_KEY]: null,
    [DENOMINATION_UNDO_KEY]: null,
  });

  // Global Ctrl+Z history for accidental edits/deletes.
  // Each item stores the complete row state before a user edit.
  const historyRef = useRef([]);
  const historyApplyingRef = useRef(false);

  const pushHistorySnapshot = (snapshot) => {
    if (historyApplyingRef.current) return;

    historyRef.current.push(
      JSON.parse(JSON.stringify(snapshot))
    );

    // Keep history bounded so it cannot grow forever.
    if (historyRef.current.length > 50) {
      historyRef.current.shift();
    }
  };

  useEffect(() => {
    const handleGlobalUndo = (event) => {
      const isUndo =
        (event.ctrlKey || event.metaKey) &&
        event.key.toLowerCase() === 'z';

      if (!isUndo) return;

      // Let text inputs behave normally while actively editing.
      // Ctrl+Z inside an edit session should undo text editing first.
      const activeElement =
        document.activeElement;

      if (
        activeElement &&
        activeElement.dataset &&
        activeElement.dataset.numberField ===
          'true' &&
        !activeElement.readOnly
      ) {
        return;
      }

      const history =
        historyRef.current;

      if (history.length === 0) {
        return;
      }

      event.preventDefault();

      const snapshot =
        history.pop();

      if (!snapshot) return;

      historyApplyingRef.current =
        true;

      try {
        if (
          snapshot.type === 'row' &&
          snapshot.row
        ) {
          const previousRows =
            rowsRef.current.map(
              (row, index) =>
                index === snapshot.rowIndex
                  ? snapshot.row
                  : row
            );

          const changedCodes =
            snapshot.row.code
              ? [snapshot.row.code]
              : [];

          rowsRef.current =
            previousRows;

          setRows(
            previousRows
          );

          persistRowsLocally(
            previousRows,
            changedCodes
          );

          // Restore the visible cells after React updates the DOM.
          requestAnimationFrame(() => {
            const rowIndex =
              previousRows.findIndex(
                (row) =>
                  row.code ===
                  snapshot.row.code
              );

            if (rowIndex >= 0) {
              const row =
                previousRows[rowIndex];

              for (
                const field
                of NUMBER_FIELDS
              ) {
                const rowElement =
                  rowRefs.current[
                    rowIndex
                  ];

                const input =
                  rowElement?.querySelector(
                    `input[data-number-field="true"][data-field="${field}"]`
                  );

                if (!input) continue;

                const rawValue =
                  row[field] ?? '';

                input.dataset.rawValue =
                  rawValue;

                input.value =
                  getDisplayInputValue(
                    rawValue
                  );

                input.readOnly = true;
              }

              updateClosingBalance(
                rowIndex
              );
            }

            setSaveStatus(
              'Undo ✓'
            );

            historyApplyingRef.current =
              false;
          });
        } else {
          historyApplyingRef.current =
            false;
        }
      } catch (error) {
        historyApplyingRef.current =
          false;

        console.error(
          'Ctrl+Z undo error:',
          error
        );
      }
    };

    window.addEventListener(
      'keydown',
      handleGlobalUndo
    );

    return () => {
      window.removeEventListener(
        'keydown',
        handleGlobalUndo
      );
    };
  }, []);

  useEffect(() => {
    rowsRef.current = rows;
  }, [rows]);

  /* -----------------------------------------
     LOAD SUPABASE DATA
  ----------------------------------------- */

  useEffect(() => {
    let cancelled = false;

    async function loadCashbook() {
      try {
        setSaveStatus('Loading cloud data...');

        const { data, error } = await supabase
          .from('master_cashbook')
          .select('*');

        if (error) {
          throw error;
        }

        if (cancelled) {
          return;
        }

        const cloudMap = {};

        if (Array.isArray(data)) {
          for (const dbRow of data) {
            if (dbRow?.code) {
              cloudMap[dbRow.code] = dbRow;
            }
          }
        }

        databaseRowsRef.current = cloudMap;

        /*
         * Supabase is the primary source.
         * If a branch is not present in Supabase,
         * fall back to LocalStorage for that branch.
         */

        const localMap = getSavedData();

        const finalRows = DEFAULT_BRANCHES.map(
          (branch, index) => {
            const cloudRow = cloudMap[branch.CODE];

            if (cloudRow) {
              return databaseRowToWebsiteRow(
                cloudRow,
                index,
                branch
              );
            }

            const localRow = localMap[branch.CODE];

            return {
              slNo: index + 1,
              code: branch.CODE,
              branch: branch.BRANCH,

              opening: localRow?.opening ?? '',
              deposit: localRow?.deposit ?? '',
              denomination:
                localRow?.denomination ?? '',
              addings: localRow?.addings ?? '',
              pendingApprovals:
                localRow?.pendingApprovals ?? '',
              finance: localRow?.finance ?? '',
              sr: localRow?.sr ?? '',
              sweeperSalary:
                localRow?.sweeperSalary ?? '',
              edits: localRow?.edits ?? '',
              apxShortage:
                localRow?.apxShortage ?? '',
              kspApprovals:
                localRow?.kspApprovals ?? '',

              remarks: localRow?.remarks ?? '',

              databaseId: null,
            };
          }
        );

        setRows(finalRows);

        /*
         * Keep local cache synchronized with
         * the cloud data that was loaded.
         */

        try {
          localStorage.setItem(
            STORAGE_KEY,
            JSON.stringify(
              finalRows.map((row) => ({
                code: row.code,
                branch: row.branch,
                opening: row.opening,
                deposit: row.deposit,
                denomination: row.denomination,
                addings: row.addings,
                pendingApprovals:
                  row.pendingApprovals,
                finance: row.finance,
                sr: row.sr,
                sweeperSalary:
                  row.sweeperSalary,
                edits: row.edits,
                apxShortage:
                  row.apxShortage,
                kspApprovals:
                  row.kspApprovals,
                remarks: row.remarks,
              }))
            )
          );
        } catch (localError) {
          console.warn(
            'Local cache update failed:',
            localError
          );
        }

        setSaveStatus(
          `Cloud Ready • ${data?.length || 0} records`
        );
      } catch (error) {
        console.error(
          'Supabase load error:',
          error
        );

        if (cancelled) {
          return;
        }

        /*
         * Supabase failed.
         * Website still works using LocalStorage.
         */

        const localRows = createRows(
          getSavedData()
        );

        databaseRowsRef.current = {};

        setRows(localRows);

        setSaveStatus(
          'Offline • Local data'
        );
      } finally {
        if (!cancelled) {
          setLoading(false);
        }
      }
    }

    loadCashbook();

    return () => {
      cancelled = true;
    };
  }, []);

  /* -----------------------------------------
     FEATURE DATA HELPERS
  ----------------------------------------- */

  const persistRowsLocally = (nextRows, dirtyCodes = []) => {
    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify(
          nextRows.map((row) => ({
            code: row.code,
            branch: row.branch,
            opening: row.opening,
            deposit: row.deposit,
            denomination: row.denomination,
            addings: row.addings,
            pendingApprovals: row.pendingApprovals,
            finance: row.finance,
            sr: row.sr,
            sweeperSalary: row.sweeperSalary,
            edits: row.edits,
            apxShortage: row.apxShortage,
            kspApprovals: row.kspApprovals,
            remarks: row.remarks,
          }))
        )
      );
    } catch (error) {
      console.error('Feature local save error:', error);
    }

    for (const code of dirtyCodes) {
      dirtyCodesRef.current.add(code);
    }

    if (dirtyCodes.length > 0) {
      setSaveStatus('Unsaved changes');
    }
  };

  const updateRowsWithField = (field, updater, successMessage = 'Updated') => {
    const nextRows = rowsRef.current.map((row) => ({
      ...row,
      [field]: updater(row),
    }));

    const changedCodes = nextRows.map((row) => row.code);

    persistRowsLocally(nextRows, changedCodes);
    rowsRef.current = nextRows;
    setRows(nextRows);
    setSaveStatus(successMessage);
  };

  const syncMasterFieldInputsToDom = (
    field,
    nextRows
  ) => {
    // React inputs use defaultValue so normal typing is not
    // controlled. Therefore a bulk Clear/Undo must explicitly
    // update the already-mounted DOM inputs too.
    nextRows.forEach((row, rowIndex) => {
      const rowElement =
        rowRefs.current[rowIndex];

      if (!rowElement) return;

      const input =
        rowElement.querySelector(
          `input[data-number-field="true"][data-field="${field}"]`
        );

      if (!input) return;

      const rawValue = row[field] ?? '';

      input.dataset.rawValue = rawValue;
      input.value =
        getDisplayInputValue(rawValue);

      input.readOnly = true;
    });

    // Closing balances depend on the cleared field.
    requestAnimationFrame(() => {
      for (
        let rowIndex = 0;
        rowIndex < nextRows.length;
        rowIndex++
      ) {
        updateClosingBalance(rowIndex);
      }
    });
  };

  const clearMasterField = (
    field,
    undoKey,
    label
  ) => {
    const currentRows = rowsRef.current;

    const snapshot = currentRows.map((row) => ({
      code: row.code,
      value: row[field] ?? '',
    }));

    undoSnapshotsRef.current[undoKey] =
      snapshot;

    setLastClearField(undoKey);

    const nextRows = currentRows.map((row) => ({
      ...row,
      [field]: '',
    }));

    const dirtyCodes = nextRows.map(
      (row) => row.code
    );

    // Update both the React state/local cache and the
    // visible input elements immediately.
    syncMasterFieldInputsToDom(
      field,
      nextRows
    );

    persistRowsLocally(
      nextRows,
      dirtyCodes
    );

    rowsRef.current = nextRows;
    setRows(nextRows);
    setSaveStatus(
      `Cleared ${label}`
    );
  };

  const confirmAndClearMasterField = (
    field,
    undoKey,
    label
  ) => {
    const confirmed = window.confirm(
      `Are you sure you want to clear all ${label}?\n\nAll 107 branches will be cleared. The values will become blank, not 0.`
    );

    if (!confirmed) return;

    clearMasterField(field, undoKey, label);
  };

  const undoClearMasterField = (
    field,
    undoKey,
    label
  ) => {
    const snapshot =
      undoSnapshotsRef.current[
        undoKey
      ];

    if (!snapshot) return;

    const snapshotMap = new Map(
      snapshot.map((item) => [
        item.code,
        item.value,
      ])
    );

    const nextRows =
      rowsRef.current.map((row) => ({
        ...row,
        [field]:
          snapshotMap.get(row.code) ??
          '',
      }));

    const dirtyCodes = nextRows.map(
      (row) => row.code
    );

    // Restore the actual visible input values too.
    syncMasterFieldInputsToDom(
      field,
      nextRows
    );

    persistRowsLocally(
      nextRows,
      dirtyCodes
    );

    rowsRef.current = nextRows;
    setRows(nextRows);

    undoSnapshotsRef.current[
      undoKey
    ] = null;

    setLastClearField(null);

    setSaveStatus(
      `${label} restored`
    );
  };

  const uploadExcelFile = async (file, mode) => {
    if (!file) return;

    try {
      const buffer = await file.arrayBuffer();

      const workbook = XLSX.read(buffer, {
        type: 'array',
        cellDates: true,
      });

      const sheetName =
        workbook.SheetNames[0];

      const sheet =
        workbook.Sheets[sheetName];

      if (!sheet) {
        throw new Error(
          'No worksheet found in the uploaded Excel file.'
        );
      }

      /*
       * Some source Excel files have report/title rows above
       * the actual table header.
       *
       * Opening Balance sample:
       *   Header is on the first row.
       *
       * Voucher Approval / Addins sample:
       *   Header is around row 10.
       *
       * Detect the real header row instead of assuming row 1.
       */
      const rawMatrix =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            header: 1,
            defval: '',
            raw: false,
          }
        );

      const normalizedCell = (value) =>
        String(value ?? '')
          .replace(/\\u00a0/g, ' ')
          .trim()
          .toLowerCase();

      const headerRowIndex =
        rawMatrix.findIndex((row) => {
          const headers = row.map(
            normalizedCell
          );

          const hasBranch =
            headers.includes('branch');

          const hasValue =
            mode === 'opening'
              ? headers.includes(
                  'opening balance'
                )
              : headers.includes(
                  'voucher value'
                );

          return (
            hasBranch &&
            hasValue
          );
        });

      if (headerRowIndex < 0) {
        throw new Error(
          mode === 'opening'
            ? 'Required columns not found. Expected: Branch, Opening Balance.'
            : 'Required columns not found. Expected: Branch, Voucher Value.'
        );
      }

      const jsonRows =
        XLSX.utils.sheet_to_json(
          sheet,
          {
            range: headerRowIndex,
            defval: '',
            raw: false,
          }
        );

      if (!jsonRows.length) {
        throw new Error(
          'The uploaded Excel sheet has no data rows.'
        );
      }

      const normalizedRows =
        jsonRows.map((item) => {
          const out = {};

          for (
            const [key, value]
            of Object.entries(item)
          ) {
            out[
              String(key)
                .replace(/\\u00a0/g, ' ')
                .trim()
                .toLowerCase()
            ] = value;
          }

          return out;
        });

      const branchColumn =
        Object.keys(
          normalizedRows[0]
        ).find(
          (key) => key === 'branch'
        );

      const valueColumn =
        mode === 'opening'
          ? Object.keys(
              normalizedRows[0]
            ).find(
              (key) =>
                key ===
                'opening balance'
            )
          : Object.keys(
              normalizedRows[0]
            ).find(
              (key) =>
                key ===
                'voucher value'
            );

      if (
        !branchColumn ||
        !valueColumn
      ) {
        throw new Error(
          mode === 'opening'
            ? 'Required columns not found. Expected: Branch, Opening Balance.'
            : 'Required columns not found. Expected: Branch, Voucher Value.'
        );
      }

      const websiteBranchMap =
        new Map();

      for (
        const row of rowsRef.current
      ) {
        websiteBranchMap.set(
          normalizeBranchName(
            row.branch
          ),
          row.code
        );
      }

      const errors = [];
      const openingValues =
        new Map();
      const addinGroups =
        new Map();

      normalizedRows.forEach(
        (item) => {
          const excelBranchName =
            String(
              item[
                branchColumn
              ] ?? ''
            ).trim();

          // Ignore blank/template rows.
          if (!excelBranchName) {
            return;
          }

          const code =
            websiteBranchMap.get(
              normalizeBranchName(
                excelBranchName
              )
            );

          // Ignore branches/stores that are not in our 107-branch master.
          if (!code) {
            return;
          }

          const rawValue =
            item[valueColumn];

          if (
            rawValue === null ||
            rawValue === undefined ||
            String(rawValue).trim() === ''
          ) {
            // For Addins and Opening Balance, a matching branch
            // with a blank amount has nothing to upload.
            return;
          }

          if (mode === 'opening') {
            const parsed =
              parseOpeningBalance(
                rawValue
              );

            if (parsed === null) {
              errors.push(
                `${excelBranchName} — Invalid Opening Balance: ${rawValue}`
              );
              return;
            }

            openingValues.set(
              code,
              parsed
            );
            return;
          }

          const parsed =
            parseVoucherValue(
              rawValue
            );

          if (parsed === null) {
            errors.push(
              `${excelBranchName} — Invalid Voucher Value: ${rawValue}`
            );
            return;
          }

          if (
            !addinGroups.has(code)
          ) {
            addinGroups.set(
              code,
              []
            );
          }

          addinGroups
            .get(code)
            .push(parsed);
        }
      );

      const currentRows =
        rowsRef.current;

      let updatedRows;

      if (mode === 'opening') {
        updatedRows =
          currentRows.map(
            (row) => {
              if (
                !openingValues.has(
                  row.code
                )
              ) {
                return row;
              }

              return {
                ...row,
                opening:
                  String(
                    openingValues.get(
                      row.code
                    )
                  ),
              };
            }
          );
      } else {
        updatedRows =
          currentRows.map(
            (row) => {
              if (
                !addinGroups.has(
                  row.code
                )
              ) {
                return row;
              }

              const values =
                addinGroups.get(
                  row.code
                );

              const formula =
                values.length === 1
                  ? `=${values[0]}`
                  : `=${values.join(
                      '+'
                    )}`;

              return {
                ...row,
                addings: formula,
              };
            }
          );
      }

      const changedCodes =
        updatedRows
          .filter(
            (row, index) =>
              row !==
              currentRows[index]
          )
          .map(
            (row) =>
              row.code
          );

      if (
        changedCodes.length > 0
      ) {
        // Update visible mounted cells immediately because
        // amount inputs intentionally use defaultValue.
        syncMasterFieldInputsToDom(
          mode === 'opening'
            ? 'opening'
            : 'addings',
          updatedRows
        );

        persistRowsLocally(
          updatedRows,
          changedCodes
        );

        rowsRef.current =
          updatedRows;

        setRows(
          updatedRows
        );

        // A new upload invalidates the old Clear/Undo snapshot
        // for that same field. After uploading new Opening Balance
        // or Addins data, the related Undo button must disappear.
        if (mode === 'opening') {
          undoSnapshotsRef.current[
            OPENING_UNDO_KEY
          ] = null;
        } else {
          undoSnapshotsRef.current[
            ADDINS_UNDO_KEY
          ] = null;
        }

        setLastClearField(
          (current) => {
            const invalidatedKey =
              mode === 'opening'
                ? OPENING_UNDO_KEY
                : ADDINS_UNDO_KEY;

            return current ===
              invalidatedKey
              ? null
              : current;
          }
        );
      }

      if (
        errors.length > 0
      ) {
        setUploadErrorTitle(
          mode === 'opening'
            ? 'Opening Balance Upload Errors'
            : 'Addins Upload Errors'
        );

        setUploadErrors(
          errors
        );

        setUploadErrorOpen(
          true
        );
      } else {
        window.alert(
          mode === 'opening'
            ? `Opening Balance upload completed. ${changedCodes.length} matched branches updated.`
            : `Addins upload completed. ${changedCodes.length} matched branches updated.`
        );
      }
    } catch (error) {
      console.error(
        'Excel upload error:',
        error
      );

      setUploadErrorTitle(
        'Excel Upload Error'
      );

      setUploadErrors([
        error?.message ||
          'Unable to read the Excel file.',
      ]);

      setUploadErrorOpen(
        true
      );
    } finally {
      if (
        mode === 'opening' &&
        openingFileInputRef.current
      ) {
        openingFileInputRef.current.value =
          '';
      }

      if (
        mode === 'addins' &&
        addinsFileInputRef.current
      ) {
        addinsFileInputRef.current.value =
          '';
      }
    }
  };

  const openOpeningUpload = () => {
    openingFileInputRef.current?.click();
  };

  const openAddinsUpload = () => {
    addinsFileInputRef.current?.click();
  };

  const buildFinanceRows = () => {
    const stored = getFinanceStorageMap();
    const today = getLocalDateString();

    return rowsRef.current
      .filter((row) => String(row.finance ?? '').trim() !== '')
      .map((row, index) => {
        const saved = stored[row.code] || {};
        const billDate = saved.billDate || today;

        return {
          id: row.code,
          slNo: index + 1,
          branch: row.branch,
          amount: calculateAmount(row.finance),
          billNo: saved.billNo || '',
          remarks: saved.remarks || 'CASH TO CARD',
          date: today,
          billDate,
        };
      });
  };

  const openFinanceReport = () => {
    const today = getLocalDateString();
    setFinanceToday(today);
    setFinanceReportRows(buildFinanceRows());
    setClipboardNotice('');
    setFinanceReportOpen(true);
  };

  const updateFinanceReportRow = (id, patch) => {
    setFinanceReportRows((current) => {
      const updated = current.map((item) =>
        item.id === id
          ? { ...item, ...patch }
          : item
      );

      const storage = getFinanceStorageMap();

      for (const item of updated) {
        storage[item.id] = {
          billNo: item.billNo,
          remarks: item.remarks,
          billDate: item.billDate,
        };
      }

      saveFinanceStorageMap(storage);

      return updated;
    });
  };

  const downloadFinanceExcel = () => {
    const exportRows = financeReportRows.map((item) => ({
      'Sl.no': item.slNo,
      Branch: item.branch,
      Amount: item.amount,
      'Bill No.': item.billNo,
      Remarks: item.remarks,
      Date: item.date,
      'Bill Date': item.billDate,
      Ageing: calculateAgeing(
        financeToday,
        item.billDate
      ),
    }));

    const worksheet = XLSX.utils.json_to_sheet(
      exportRows
    );

    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 24 },
      { wch: 14 },
      { wch: 18 },
      { wch: 28 },
      { wch: 14 },
      { wch: 14 },
      { wch: 10 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Finance Report'
    );

    XLSX.writeFile(
      workbook,
      'Finance and Card Amount Pending in Cash Book.xlsx'
    );
  };

  const getFinanceTableHtml = () => {
    const header = [
      'Sl.no',
      'Branch',
      'Amount',
      'Bill No.',
      'Remarks',
      'Date',
      'Bill Date',
      'Ageing',
    ];

    const rowsHtml = financeReportRows
      .map((item) => {
        const cells = [
          item.slNo,
          item.branch,
          item.amount,
          item.billNo,
          item.remarks,
          item.date,
          item.billDate,
          calculateAgeing(
            financeToday,
            item.billDate
          ),
        ];

        return `<tr>${cells
          .map(
            (cell) =>
              `<td style="border:1px solid #b8c4cf;padding:5px 7px">${escapeHtml(cell)}</td>`
          )
          .join('')}</tr>`;
      })
      .join('');

    return `<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:12px"><thead><tr>${header
      .map(
        (cell) =>
          `<th style="border:1px solid #7e8b96;padding:5px 7px;background:#eaf0f5;font-weight:700">${escapeHtml(cell)}</th>`
      )
      .join('')}</tr></thead><tbody>${rowsHtml}</tbody></table>`;
  };

  const getFinanceTableText = () => {
    const lines = [];

    lines.push(
      [
        'Sl.no',
        'Branch',
        'Amount',
        'Bill No.',
        'Remarks',
        'Date',
        'Bill Date',
        'Ageing',
      ].join('\t')
    );

    for (const item of financeReportRows) {
      lines.push(
        [
          item.slNo,
          item.branch,
          item.amount,
          item.billNo,
          item.remarks,
          item.date,
          item.billDate,
          calculateAgeing(
            financeToday,
            item.billDate
          ),
        ].join('\t')
      );
    }

    return lines.join('\n');
  };

  const copyFinanceTableToClipboard = async () => {
    const html = getFinanceTableHtml();
    const text = getFinanceTableText();

    try {
      if (
        navigator.clipboard &&
        window.ClipboardItem
      ) {
        const item = new ClipboardItem({
          'text/html': new Blob([html], {
            type: 'text/html',
          }),
          'text/plain': new Blob([text], {
            type: 'text/plain',
          }),
        });

        await navigator.clipboard.write([item]);
      } else if (navigator.clipboard) {
        await navigator.clipboard.writeText(text);
      } else {
        throw new Error('Clipboard API unavailable');
      }

      setClipboardNotice(
        'Table copied — Ctrl+V can paste it into Outlook.'
      );
      return true;
    } catch (error) {
      console.warn(
        'Finance table clipboard copy failed:',
        error
      );
      setClipboardNotice(
        'Clipboard copy was not available. Outlook will still open with the report text.'
      );
      return false;
    }
  };

  const sendFinanceMail = async () => {
    const to = financeReportRows
      .map((item) => branchToEmail(item.branch))
      .filter(Boolean)
      .join(',');

    const cc = FINANCE_CC_EMAILS.join(',');
    const subject =
      'Finance and Card Amount Pending in Cash Book';

    // Copy the table first. Outlook body intentionally stays blank.
    await copyFinanceTableToClipboard();

    const mailto =
      `mailto:${encodeURIComponent(to)}` +
      `?cc=${encodeURIComponent(cc)}` +
      `&subject=${encodeURIComponent(subject)}`;

    window.location.href = mailto;
  };


  /* -----------------------------------------
     SR REPORT
  ----------------------------------------- */

  const getSRStorageMap = () => {
    try {
      const saved = localStorage.getItem(
        SR_REPORT_STORAGE_KEY
      );
      if (!saved) return {};
      const parsed = JSON.parse(saved);
      return parsed && typeof parsed === 'object'
        ? parsed
        : {};
    } catch {
      return {};
    }
  };

  const saveSRStorageMap = (map) => {
    try {
      localStorage.setItem(
        SR_REPORT_STORAGE_KEY,
        JSON.stringify(map)
      );
    } catch (error) {
      console.error('SR report local save error:', error);
    }
  };

  const buildSRRows = () => {
    const stored = getSRStorageMap();
    const today = getLocalDateString();

    return rowsRef.current
      .filter((row) => String(row.sr ?? '').trim() !== '')
      .map((row, index) => {
        const saved = stored[row.code] || {};
        const billDate = saved.billDate || today;

        return {
          id: row.code,
          slNo: index + 1,
          branch: row.branch,
          amount: calculateAmount(row.sr),
          billNo: saved.billNo || '',
          remarks: saved.remarks || 'CASH TO CARD',
          date: today,
          billDate,
        };
      });
  };

  const openSRReport = () => {
    const today = getLocalDateString();
    setSrToday(today);
    setSrReportRows(buildSRRows());
    setSrReportOpen(true);
    setClipboardNotice('');
  };

  const updateSRReportRow = (id, patch) => {
    setSrReportRows((current) => {
      const updated = current.map((item) =>
        item.id === id
          ? { ...item, ...patch }
          : item
      );

      const storage = getSRStorageMap();
      for (const item of updated) {
        storage[item.id] = {
          billNo: item.billNo,
          remarks: item.remarks,
          billDate: item.billDate,
        };
      }
      saveSRStorageMap(storage);
      return updated;
    });
  };

  const downloadSRExcel = () => {
    const exportRows = srReportRows.map((item) => ({
      'Sl.no': item.slNo,
      Branch: item.branch,
      Amount: item.amount,
      'Bill No.': item.billNo,
      Remarks: item.remarks,
      Date: item.date,
      'Bill Date': item.billDate,
      Ageing: calculateAgeing(
        srToday,
        item.billDate
      ),
    }));

    const worksheet =
      XLSX.utils.json_to_sheet(exportRows);

    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 24 },
      { wch: 14 },
      { wch: 18 },
      { wch: 28 },
      { wch: 14 },
      { wch: 14 },
      { wch: 10 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'SR Report'
    );

    XLSX.writeFile(
      workbook,
      'SRN Pending Stores.xlsx'
    );
  };

  const getSRTableHtml = () => {
    const header = [
      'Sl.no',
      'Branch',
      'Amount',
      'Bill No.',
      'Remarks',
      'Date',
      'Bill Date',
      'Ageing',
    ];

    const rowsHtml = srReportRows
      .map((item) => {
        const cells = [
          item.slNo,
          item.branch,
          item.amount,
          item.billNo,
          item.remarks,
          item.date,
          item.billDate,
          calculateAgeing(
            srToday,
            item.billDate
          ),
        ];

        return `<tr>${cells
          .map(
            (cell) =>
              `<td style="border:1px solid #b8c4cf;padding:5px 7px">${escapeHtml(cell)}</td>`
          )
          .join('')}</tr>`;
      })
      .join('');

    return `<table style="border-collapse:collapse;font-family:Segoe UI,Arial,sans-serif;font-size:12px"><thead><tr>${header
      .map(
        (cell) =>
          `<th style="border:1px solid #7e8b96;padding:5px 7px;background:#eaf0f5;font-weight:700">${escapeHtml(cell)}</th>`
      )
      .join('')}</tr></thead><tbody>${rowsHtml}</tbody></table>`;
  };

  const getSRTableText = () => {
    const lines = [];
    lines.push(
      [
        'Sl.no',
        'Branch',
        'Amount',
        'Bill No.',
        'Remarks',
        'Date',
        'Bill Date',
        'Ageing',
      ].join('\t')
    );

    for (const item of srReportRows) {
      lines.push(
        [
          item.slNo,
          item.branch,
          item.amount,
          item.billNo,
          item.remarks,
          item.date,
          item.billDate,
          calculateAgeing(
            srToday,
            item.billDate
          ),
        ].join('\t')
      );
    }
    return lines.join('\n');
  };

  const copySRTableToClipboard = async () => {
    const html = getSRTableHtml();
    const text = getSRTableText();

    try {
      if (
        navigator.clipboard &&
        window.ClipboardItem
      ) {
        const item = new ClipboardItem({
          'text/html': new Blob([html], {
            type: 'text/html',
          }),
          'text/plain': new Blob([text], {
            type: 'text/plain',
          }),
        });

        await navigator.clipboard.write([item]);
      } else if (navigator.clipboard) {
        await navigator.clipboard.writeText(text);
      } else {
        throw new Error('Clipboard API unavailable');
      }

      setClipboardNotice(
        'Table copied — Ctrl+V can paste it into Outlook.'
      );
      return true;
    } catch (error) {
      console.warn('SR table clipboard copy failed:', error);
      setClipboardNotice('Table copy was not available.');
      return false;
    }
  };

  const sendSRMail = async () => {
    const to = srReportRows
      .map((item) => branchToEmail(item.branch))
      .filter(Boolean)
      .join(',');

    await copySRTableToClipboard();

    const mailto =
      `mailto:${encodeURIComponent(to)}` +
      `?cc=${encodeURIComponent(
        SR_CC_EMAILS.join(',')
      )}` +
      `&subject=${encodeURIComponent(
        'SRN Pending Stores'
      )}`;

    window.location.href = mailto;
  };

  /* -----------------------------------------
     SIMPLE LIST REPORTS
  ----------------------------------------- */

  const buildAmountList = (field) =>
    rowsRef.current
      .filter(
        (row) =>
          String(row[field] ?? '').trim() !== '' &&
          calculateAmount(row[field]) !== 0
      )
      .map((row) => ({
        code: row.code,
        branch: row.branch,
        amount: calculateAmount(row[field]),
      }));

  // Pending Approvals needs to show each component of a stored
  // additive formula as a separate row.
  //
  // Example:
  //   =1000+300
  // becomes:
  //   AMEERPET  1000
  //   AMEERPET   300
  //
  // The store-count on the button remains a store count, not a
  // component-row count.
  const buildPendingApprovalList = () => {
    const output = [];

    rowsRef.current.forEach((row) => {
      const raw =
        String(
          row.pendingApprovals ?? ''
        ).trim();

      if (!raw) return;

      const calculated =
        calculateAmount(raw);

      if (calculated === 0) return;

      const expression =
        raw.replace(/^=/, '').trim();

      // Pending Approvals formulas are additive, e.g. =1000+300+500.
      // Split the original expression into its individual voucher amounts.
      const parts =
        expression
          .split('+')
          .map((part) => part.trim())
          .filter(Boolean);

      if (parts.length > 1) {
        parts.forEach((part) => {
          const value = Number(part);

          if (
            Number.isFinite(value) &&
            value !== 0
          ) {
            output.push({
              code: row.code,
              branch: row.branch,
              amount: value,
            });
          }
        });

        return;
      }

      // Single value / non-formula value.
      output.push({
        code: row.code,
        branch: row.branch,
        amount: calculated,
      });
    });

    return output;
  };

  const getReportStoreCount = (field) =>
    buildAmountList(field).length;

  const getPendingStatusStoreCount = () =>
    getPendingCashbooks().length +
    getPendingDepositSlips().length;

  const getLowNoCashStoreCount = () =>
    buildLowCashRows().length;

  const openPendingApprovals = () => {
    setPendingApprovalsOpen(true);
  };

  const openEdits = () => {
    setEditsOpen(true);
  };

  const downloadSimpleListExcel = (
    title,
    sheetName,
    filename,
    field
  ) => {
    const list =
      field === 'pendingApprovals'
        ? buildPendingApprovalList()
        : buildAmountList(field);

    const data = list.map((item, index) => ({
      'Sl.No.': index + 1,
      CODE: item.code,
      BRANCH: item.branch,
      AMOUNT: item.amount,
    }));

    const worksheet =
      XLSX.utils.json_to_sheet(data);

    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 12 },
      { wch: 28 },
      { wch: 16 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      sheetName
    );
    XLSX.writeFile(workbook, filename);
  };

  /* -----------------------------------------
     PENDING STATUS
  ----------------------------------------- */

  const isPositiveAmount = (value) =>
    calculateAmount(value) > 0;

  const isEmptyOrZeroAmount = (value) =>
    String(value ?? '').trim() === '' ||
    calculateAmount(value) === 0;

  const getPendingCashbooks = () =>
    rowsRef.current.filter(
      (row) =>
        isEmptyOrZeroAmount(row.denomination) &&
        isEmptyOrZeroAmount(row.deposit)
    );

  const getPendingDepositSlips = () =>
    rowsRef.current.filter(
      (row) =>
        isPositiveAmount(row.denomination) &&
        isEmptyOrZeroAmount(row.deposit)
    );

  /* -----------------------------------------
     LOW / NO CASH
  ----------------------------------------- */

  const getLowCashStorageMap = () => {
    try {
      const saved = localStorage.getItem(
        LOW_CASH_STORAGE_KEY
      );
      if (!saved) return {};
      const parsed = JSON.parse(saved);
      return parsed && typeof parsed === 'object'
        ? parsed
        : {};
    } catch {
      return {};
    }
  };

  const saveLowCashStorageMap = (map) => {
    try {
      localStorage.setItem(
        LOW_CASH_STORAGE_KEY,
        JSON.stringify(map)
      );
    } catch (error) {
      console.error('Low cash local save error:', error);
    }
  };

  const getLowCashDefaultRemark = (row) => {
    const amount = calculateAmount(
      row.denomination
    );

    if (amount <= 0) {
      return 'No Cash';
    }

    if (amount >= 1 && amount <= 500) {
      return 'Low Cash';
    }

    return '';
  };

  const buildLowCashRows = () => {
    const stored = getLowCashStorageMap();

    const reportRows = rowsRef.current.map((row) => {
      const denomination =
        calculateAmount(row.denomination);

      const defaultRemark =
        getLowCashDefaultRemark(row);

      return {
        code: row.code,
        branch: row.branch,
        denomination,
        status:
          stored[row.code] !== undefined
            ? stored[row.code]
            : defaultRemark,
      };
    });

    // Sorting:
    // 1) No Cash (blank / 0) first
    // 2) Low Cash (1–500) next
    // 3) 501+ last
    // Within the same group, lower denomination first.
    return reportRows.sort((a, b) => {
      const getGroup = (amount) => {
        if (amount <= 0) return 0;
        if (amount <= 500) return 1;
        return 2;
      };

      const groupA = getGroup(a.denomination);
      const groupB = getGroup(b.denomination);

      if (groupA !== groupB) {
        return groupA - groupB;
      }

      if (a.denomination !== b.denomination) {
        return a.denomination - b.denomination;
      }

      return a.branch.localeCompare(
        b.branch,
        undefined,
        { sensitivity: 'base' }
      );
    });
  };

  const openLowCashReport = () => {
    setLowCashRows(buildLowCashRows());
    setLowCashOpen(true);
  };

  const updateLowCashStatus = (code, value) => {
    setLowCashRows((current) =>
      current.map((item) =>
        item.code === code
          ? { ...item, status: value }
          : item
      )
    );

    const storage = getLowCashStorageMap();
    storage[code] = value;
    saveLowCashStorageMap(storage);
  };

  const downloadLowCashExcel = () => {
    const data = lowCashRows.map(
      (item, index) => ({
        'Sl.No.': index + 1,
        'Store Name': item.branch,
        'Denom Amt': item.denomination,
        'Status / Remarks': item.status,
      })
    );

    const worksheet =
      XLSX.utils.json_to_sheet(data);

    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 28 },
      { wch: 14 },
      { wch: 32 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Low Cash & No Cash'
    );

    XLSX.writeFile(
      workbook,
      'Low Cash & No Cash Report.xlsx'
    );
  };

  /* -----------------------------------------
     CONTRA SHEET
  ----------------------------------------- */

  const toYYYYMMDD = (date = new Date()) =>
    `${date.getFullYear()}${String(
      date.getMonth() + 1
    ).padStart(2, '0')}${String(
      date.getDate()
    ).padStart(2, '0')}`;

  const downloadContraSheet = () => {
    const now = new Date();
    const yyyymmdd = toYYYYMMDD(now);
    const day = String(now.getDate()).padStart(2, '0');

    const headers = [
      'Voucher Date',
      'Credit Account',
      'Debit Account',
      'Payment Mode',
      'Amount',
      'TXN No',
      'TXN Date',
      'Bank Narration',
      'Branch Short Code',
      'Ref Doc No',
      'Ref Doc Date',
      'Salesperson Name',
      'Product Category',
      'Deposited Date',
    ];

    const liveRows =
      rowsRef.current.map(
        (row, rowIndex) => {
          const liveRow =
            readRowFromDom(rowIndex);

          return liveRow
            ? {
                ...row,
                ...liveRow,
              }
            : row;
        }
      );

    const rowsForExport = liveRows
      .filter(
        (row) =>
          isPositiveAmount(row.deposit)
      )
      .map((row) => [
        yyyymmdd,
        'Cash A/c.',
        'HDFC BANK OD A/C: 57500001930168',
        'CASH',
        calculateAmount(row.deposit),
        day,
        yyyymmdd,
        'Being Entry Passing words Cash deposited into bank',
        row.code,
        day,
        yyyymmdd,
        '',
        '',
        yyyymmdd,
      ]);

    const data = [
      headers,
      headers.map(() => 'ABCD1234'),
      ...rowsForExport,
    ];

    const worksheet =
      XLSX.utils.aoa_to_sheet(data);

    worksheet['!cols'] = [
      { wch: 15 },
      { wch: 22 },
      { wch: 34 },
      { wch: 14 },
      { wch: 14 },
      { wch: 12 },
      { wch: 15 },
      { wch: 43 },
      { wch: 18 },
      { wch: 14 },
      { wch: 15 },
      { wch: 18 },
      { wch: 20 },
      { wch: 16 },
    ];

    // Keep the dummy row exactly under the header.
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'Contra Export'
    );

    XLSX.writeFile(
      workbook,
      `CONTRA_UPLOAD_SHEET ${yyyymmdd}.xlsx`
    );
  };

  /* -----------------------------------------
     MASTER CASHBOOK EXCEL
  ----------------------------------------- */

  const writeExcelCellStyle = (
    worksheet,
    address,
    fill,
    fontColor = '#17324d'
  ) => {
    if (!worksheet[address]) {
      worksheet[address] = {
        t: 's',
        v: '',
      };
    }

    worksheet[address].s = {
      fill: {
        patternType: 'solid',
        fgColor: { rgb: fill.replace('#', '').toUpperCase() },
      },
      font: {
        bold: true,
        color: {
          rgb: fontColor.replace('#', '').toUpperCase(),
        },
      },
      alignment: {
        vertical: 'center',
        horizontal: 'center',
        wrapText: true,
      },
      border: {
        top: { style: 'thin', color: { rgb: 'B8C7D4' } },
        bottom: { style: 'thin', color: { rgb: 'B8C7D4' } },
        left: { style: 'thin', color: { rgb: 'B8C7D4' } },
        right: { style: 'thin', color: { rgb: 'B8C7D4' } },
      },
    };
  };

  const toExcelValue = (raw) => {
    const value = String(raw ?? '').trim();

    if (value.startsWith('=')) {
      return {
        formula: value.slice(1),
        result: calculateAmount(value),
      };
    }

    const numeric = parseAmountNumber(value);
    if (numeric !== null && value !== '') {
      return numeric;
    }

    return value;
  };

  const downloadMasterCashbookExcel = () => {
    const headers = [...HEADERS];

    const data = [
      headers,
      ...rowsRef.current.map((row) => [
        row.slNo,
        row.code,
        row.branch,
        row.opening,
        row.deposit,
        row.denomination,
        row.addings,
        row.pendingApprovals,
        row.finance,
        row.sr,
        row.sweeperSalary,
        row.edits,
        row.apxShortage,
        row.kspApprovals,
        calculateClosing(row),
        row.remarks,
      ]),
    ];

    const worksheet =
      XLSX.utils.aoa_to_sheet(data);

    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 10 },
      { wch: 23 },
      { wch: 15 },
      { wch: 13 },
      { wch: 15 },
      { wch: 13 },
      { wch: 16 },
      { wch: 15 },
      { wch: 11 },
      { wch: 15 },
      { wch: 13 },
      { wch: 16 },
      { wch: 18 },
      { wch: 16 },
      { wch: 32 },
    ];

    // Preserve formula cells as actual Excel formulas.
    for (
      let rowIndex = 0;
      rowIndex < rowsRef.current.length;
      rowIndex++
    ) {
      const row = rowsRef.current[rowIndex];
      const dataRow = rowIndex + 2;

      NUMBER_FIELDS.forEach((field, fieldIndex) => {
        const colIndex =
          3 + fieldIndex; // D is index 3

        const address =
          XLSX.utils.encode_cell({
            r: dataRow - 1,
            c: colIndex,
          });

        const converted = toExcelValue(
          row[field]
        );

        if (
          converted &&
          typeof converted === 'object' &&
          converted.formula
        ) {
          worksheet[address] = {
            t: 'n',
            f: converted.formula,
            v: converted.result,
          };
        }
      });

      const closingAddress =
        XLSX.utils.encode_cell({
          r: dataRow - 1,
          c: 14,
        });

      worksheet[closingAddress] = {
        t: 'n',
        v: calculateClosing(row),
      };
    }

    // Colourful header groups, matching the website.
    const headerFills = [
      '2F80ED',
      '56B4E9',
      '56B4E9',
      'BFE8FF',
      'BFE8FF',
      'BFE8FF',
      'BDECCB',
      'BDECCB',
      'BDECCB',
      'FFE3AE',
      'FFE3AE',
      'FFE3AE',
      'FFD1D1',
      'FFD1D1',
      'C9F1D2',
      'E4D4FF',
    ];

    for (let col = 0; col < headers.length; col++) {
      const address =
        XLSX.utils.encode_cell({
          r: 0,
          c: col,
        });
      writeExcelCellStyle(
        worksheet,
        address,
        headerFills[col] || 'E9EEF5'
      );
    }

    for (
      let rowIndex = 1;
      rowIndex <= rowsRef.current.length;
      rowIndex++
    ) {
      for (
        let col = 0;
        col < headers.length;
        col++
      ) {
        const address =
          XLSX.utils.encode_cell({
            r: rowIndex,
            c: col,
          });

        if (worksheet[address]) {
          worksheet[address].s = {
            fill: {
              patternType: 'solid',
              fgColor: {
                rgb:
                  rowIndex % 2 === 0
                    ? 'F8FBFD'
                    : 'FFFFFF',
              },
            },
            alignment: {
              vertical: 'center',
              horizontal:
                col === 15
                  ? 'left'
                  : col === 2
                    ? 'left'
                    : 'center',
            },
            border: {
              top: { style: 'thin', color: { rgb: 'D7E0E8' } },
              bottom: { style: 'thin', color: { rgb: 'D7E0E8' } },
              left: { style: 'thin', color: { rgb: 'D7E0E8' } },
              right: { style: 'thin', color: { rgb: 'D7E0E8' } },
            },
          };
        }
      }
    }

    worksheet['!freeze'] = { xSplit: 3, ySplit: 1 };

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(
      workbook,
      worksheet,
      'MASTER REPORT'
    );

    XLSX.writeFile(
      workbook,
      `HAPPI MASTER CASHBOOK ${toYYYYMMDD()}.xlsx`
    );
  };

  /* -----------------------------------------
     UPDATE ONE CLOSING BALANCE
  ----------------------------------------- */

  const updateClosingBalance = (rowIndex) => {
    const rowElement =
      rowRefs.current[rowIndex];

    if (!rowElement) {
      return;
    }

    const inputs =
      rowElement.querySelectorAll(
        'input[data-number-field="true"]'
      );

    const values = {};

    let hasValue = false;

    inputs.forEach((input) => {
      const value =
        getStoredInputValue(input).trim();

      values[input.dataset.field] =
        value;

      if (value !== '') {
        hasValue = true;
      }
    });

    const closingElement =
      rowElement.querySelector(
        '.closing-value'
      );

    if (!closingElement) {
      return;
    }

    if (!hasValue) {
      closingElement.textContent = '';

      closingElement.classList.remove(
        'negative'
      );

      return;
    }

    const closing =
      calculateClosing(values);

    closingElement.textContent =
      formatAmount(closing);

    closingElement.classList.toggle(
      'negative',
      closing < 0
    );
  };

  /* -----------------------------------------
     UPDATE ALL CLOSING BALANCES
  ----------------------------------------- */

  const updateAllClosingBalances = () => {
    for (
      let rowIndex = 0;
      rowIndex < rows.length;
      rowIndex++
    ) {
      updateClosingBalance(rowIndex);
    }
  };

  useEffect(() => {
    if (!loading && rows.length > 0) {
      updateAllClosingBalances();
    }
  }, [loading, rows]);

  /* -----------------------------------------
     EDITING BEHAVIOUR

     Amount cells are NOT editable just because the
     keyboard/tab reaches them.

     Edit is allowed only by:
       1. F2
       2. Double-click

     Once editing starts, the original formula is shown.
     Example:
       display: 6600
       F2/double-click -> =6000+600
       blur -> 6600
  ----------------------------------------- */

  const getEditingKey = (rowIndex, field) =>
    `${rowIndex}:${field}`;

  const startNumberEdit = (event, rowIndex, field) => {
    const input = event.currentTarget;

    // Capture the row before entering an edit session so the user can
    // Ctrl+Z the whole edit back if they make a mistake.
    const beforeRow =
      rowsRef.current[rowIndex];

    if (
      beforeRow &&
      editingCellRef.current !==
        getEditingKey(rowIndex, field)
    ) {
      pushHistorySnapshot({
        type: 'row',
        rowIndex,
        row: beforeRow,
      });
    }

    editingCellRef.current = getEditingKey(
      rowIndex,
      field
    );

    if (input.dataset.rawValue !== undefined) {
      input.value = input.dataset.rawValue;
    }

    input.readOnly = false;
    input.style.caretColor = 'currentColor';
    input.focus();

    // Put the caret at the end, like Excel's F2 edit behaviour.
    const placeCaretAtEnd = () => {
      const length = input.value.length;

      try {
        input.setSelectionRange(
          length,
          length
        );
      } catch {
        // Ignore selection errors for unsupported input states.
      }
    };

    // Browsers can apply focus/readOnly changes asynchronously, so do it
    // once immediately and once on the next frame to guarantee the blinking
    // caret is visible.
    placeCaretAtEnd();
    requestAnimationFrame(placeCaretAtEnd);
  };

  const handleNumberFocus = (event) => {
    const input = event.currentTarget;

    // Focus alone must never switch the cell into edit mode.
    // It is intentionally left read-only until F2/double-click.
    if (!input.readOnly) {
      if (input.dataset.rawValue !== undefined) {
        input.value = input.dataset.rawValue;
      }
    }
  };

  const moveCellSelection = (
    rowIndex,
    field,
    rowDelta,
    fieldDelta
  ) => {
    const currentFieldIndex =
      NUMBER_FIELDS.indexOf(field);

    if (currentFieldIndex < 0) return;

    const nextRowIndex = Math.max(
      0,
      Math.min(
        rows.length - 1,
        rowIndex + rowDelta
      )
    );

    const nextFieldIndex = Math.max(
      0,
      Math.min(
        NUMBER_FIELDS.length - 1,
        currentFieldIndex + fieldDelta
      )
    );

    const nextField =
      NUMBER_FIELDS[nextFieldIndex];

    const nextKey = getEditingKey(
      nextRowIndex,
      nextField
    );

    const nextInput =
      rowRefs.current[nextRowIndex]?.querySelector(
        `input[data-field="${nextField}"][data-number-field="true"]`
      );

    if (!nextInput) return;

    editingCellRef.current = null;

    document
      .querySelectorAll(
        'input[data-number-field="true"]'
      )
      .forEach((el) => {
        el.readOnly = true;
      });

    nextInput.focus();

    // Selection/navigation must not enter edit mode.
    nextInput.readOnly = true;

    // Keep the visible calculated value when simply moving.
    const rawValue =
      nextInput.dataset.rawValue ?? '';

    nextInput.value =
      getDisplayInputValue(rawValue);

    return nextKey;
  };

  const clearNumberCell = (
    event,
    rowIndex,
    field
  ) => {
    const input = event.currentTarget;

    event.preventDefault();

    // Whole-cell Delete is allowed only when the cell is NOT in
    // explicit edit mode. In edit mode, Delete is the normal
    // character-delete key and must not wipe the whole cell.
    if (!input.readOnly) {
      return;
    }

    const beforeRow =
      rowsRef.current[rowIndex];

    if (beforeRow) {
      pushHistorySnapshot({
        type: 'row',
        rowIndex,
        row: beforeRow,
      });
    }

    input.dataset.rawValue = '';
    input.value = '';

    editingCellRef.current = null;
    input.readOnly = true;

    updateClosingBalance(rowIndex);
    markRowDirty(rowIndex);

    // Keep the cleared cell selected.
    requestAnimationFrame(() => {
      input.focus();
    });
  };

  const handleNumberKeyDown = (
    event,
    rowIndex,
    field
  ) => {
    const input = event.currentTarget;

    // Navigation mode:
    // Arrow keys move between cells only when the cell is NOT being edited.
    //
    // Edit mode:
    // Arrow keys are left completely to the text input so the caret moves
    // through the formula/value instead of jumping to the neighbouring cell.
    if (
      ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight']
        .includes(event.key)
    ) {
      if (!input.readOnly) {
        return;
      }

      event.preventDefault();

      const rowDelta =
        event.key === 'ArrowUp'
          ? -1
          : event.key === 'ArrowDown'
            ? 1
            : 0;

      const fieldDelta =
        event.key === 'ArrowLeft'
          ? -1
          : event.key === 'ArrowRight'
            ? 1
            : 0;

      moveCellSelection(
        rowIndex,
        field,
        rowDelta,
        fieldDelta
      );

      return;
    }

    // Delete:
    // • In edit mode -> let the browser delete the character after the caret.
    // • Outside edit mode -> clear the whole cell.
    if (event.key === 'Delete') {
      if (!input.readOnly) {
        return;
      }

      clearNumberCell(
        event,
        rowIndex,
        field
      );
      return;
    }

    // Backspace:
    // • In edit mode -> normal one-character-at-a-time editing.
    // • Outside edit mode -> do not clear the whole cell.
    if (event.key === 'Backspace') {
      if (input.readOnly) {
        event.preventDefault();
      }
      return;
    }

    // Existing data: F2 is an explicit edit command.
    if (event.key === 'F2') {
      event.preventDefault();
      startNumberEdit(
        event,
        rowIndex,
        field
      );
      return;
    }

    // EMPTY CELL:
    // Directly typing a number/formula starts editing.
    // F2 is NOT required for a blank cell.
    if (
      input.readOnly &&
      input.value.trim() === '' &&
      event.key.length === 1 &&
      !event.ctrlKey &&
      !event.altKey &&
      !event.metaKey
    ) {
      event.preventDefault();

      startNumberEdit(
        event,
        rowIndex,
        field
      );

      input.value = event.key;
      input.dataset.rawValue = event.key;

      requestAnimationFrame(() => {
        try {
          const length =
            input.value.length;
          input.setSelectionRange(
            length,
            length
          );
        } catch {
          // Ignore selection errors.
        }
      });

      return;
    }

    if (input.readOnly) {
      return;
    }

    if (event.key === 'Enter') {
      event.preventDefault();
      input.blur();
      return;
    }

    if (event.key === 'Escape') {
      event.preventDefault();

      const original =
        input.dataset.rawValue ?? '';

      input.value =
        getDisplayInputValue(original);

      editingCellRef.current = null;
      input.readOnly = true;
      input.blur();
    }
  };

  const handleNumberDoubleClick = (
    event,
    rowIndex,
    field
  ) => {
    event.preventDefault();
    startNumberEdit(event, rowIndex, field);
  };

  const handleNumberInput = (event, rowIndex) => {
    const input = event.currentTarget;

    // Safety guard: typing is accepted only while this cell
    // is explicitly in edit mode.
    if (input.readOnly) {
      return;
    }

    input.dataset.rawValue = input.value;

    updateClosingBalance(rowIndex);
    markRowDirty(rowIndex);
  };

  const handleNumberBlur = (event, rowIndex) => {
    const input = event.currentTarget;

    // If it was never explicitly edited, leave it untouched.
    if (input.readOnly) {
      return;
    }

    const rawValue = input.value.trim();

    input.dataset.rawValue = rawValue;

    if (rawValue !== '') {
      const isExpression =
        rawValue.startsWith('=') ||
        /[+\-*/()]/.test(rawValue);

      if (isExpression) {
        const calculatedValue =
          calculateAmount(rawValue);

        if (Number.isFinite(calculatedValue)) {
          input.value = String(calculatedValue);
        }
      }
    }

    input.readOnly = true;
    editingCellRef.current = null;

    updateClosingBalance(rowIndex);
    markRowDirty(rowIndex);
  };

  /* -----------------------------------------
     READ CURRENT ROW FROM DOM
  ----------------------------------------- */

  const readRowFromDom = (rowIndex) => {
    const rowElement = rowRefs.current[rowIndex];

    if (!rowElement || !rows[rowIndex]) {
      return null;
    }

    const row = {
      code: rows[rowIndex].code,
      branch: rows[rowIndex].branch,
      databaseId: rows[rowIndex].databaseId ?? null,
    };

    const numberInputs =
      rowElement.querySelectorAll(
        'input[data-number-field="true"]'
      );

    numberInputs.forEach((input) => {
      row[input.dataset.field] =
        getStoredInputValue(input);
    });

    const remarksInput =
      rowElement.querySelector('.remarks-input');

    row.remarks = remarksInput?.value || '';

    return row;
  };

  const saveRowToLocalCache = (rowIndex) => {
    const row = readRowFromDom(rowIndex);

    if (!row) return;

    try {
      const localMap = getSavedData();

      localMap[row.code] = row;

      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify(
          Object.values(localMap)
        )
      );
    } catch (error) {
      console.error(
        'Local cache save error:',
        error
      );
    }
  };

  /* -----------------------------------------
     MARK ROW DIRTY

     Every edit is saved locally immediately.
     Supabase is deliberately NOT called here.
  ----------------------------------------- */

  const markRowDirty = (rowIndex) => {
    const baseRow = rowsRef.current[rowIndex];

    if (!baseRow) return;

    // Read the live DOM value so reports/exports immediately see
    // the amount that the user just typed, without forcing a React
    // table re-render on every keystroke.
    const liveRow =
      readRowFromDom(rowIndex);

    if (liveRow) {
      rowsRef.current[rowIndex] = {
        ...baseRow,
        ...liveRow,
      };
    }

    const currentRow =
      rowsRef.current[rowIndex];

    // Save the latest row to LocalStorage immediately.
    try {
      const localMap = getSavedData();

      localMap[currentRow.code] =
        currentRow;

      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify(
          Object.values(localMap)
        )
      );
    } catch (error) {
      console.error(
        'Local cache save error:',
        error
      );
    }

    dirtyCodesRef.current.add(
      currentRow.code
    );

    setSaveStatus('Unsaved changes');
  };

  /* -----------------------------------------
     10-MINUTE CLOUD SYNC

     Only dirty rows are uploaded.
     Existing rows are updated in one request.
     New rows are inserted in one request.

     This avoids:
       - sync on every keystroke
       - 1.2-second sync delays
       - repeated requests while entering data
       - website freezing during frequent cloud writes
  ----------------------------------------- */

  const syncDirtyRows = async () => {
    if (autoSyncInFlightRef.current) {
      return;
    }

    const dirtyCodes =
      Array.from(dirtyCodesRef.current);

    if (dirtyCodes.length === 0) {
      return;
    }

    const dirtyCodeSet = new Set(dirtyCodes);

    const rowsToSync = [];

    for (
      let rowIndex = 0;
      rowIndex < rows.length;
      rowIndex++
    ) {
      const row = rows[rowIndex];

      if (row && dirtyCodeSet.has(row.code)) {
        const currentRow =
          readRowFromDom(rowIndex);

        if (currentRow) {
          rowsToSync.push(currentRow);
        }
      }
    }

    if (rowsToSync.length === 0) {
      return;
    }

    autoSyncInFlightRef.current = true;

    setSaveStatus(
      `Syncing ${rowsToSync.length} changed row${rowsToSync.length === 1 ? '' : 's'}...`
    );

    try {
      const existingRows = [];
      const newRows = [];

      for (const row of rowsToSync) {
        const existingDatabaseRow =
          databaseRowsRef.current[row.code];

        const payload =
          websiteRowToDatabaseRow(
            row,
            existingDatabaseRow?.id ??
              row.databaseId ??
              null
          );

        if (
          existingDatabaseRow?.id ||
          row.databaseId
        ) {
          existingRows.push(payload);
        } else {
          newRows.push(payload);
        }
      }

      // One request for all existing changed rows.
      if (existingRows.length > 0) {
        const { error } =
          await supabase
            .from('master_cashbook')
            .upsert(existingRows);

        if (error) {
          throw error;
        }
      }

      // One request for all new rows.
      if (newRows.length > 0) {
        const { data, error } =
          await supabase
            .from('master_cashbook')
            .insert(newRows)
            .select('*');

        if (error) {
          throw error;
        }

        if (Array.isArray(data)) {
          for (const dbRow of data) {
            if (dbRow?.code) {
              databaseRowsRef.current[
                dbRow.code
              ] = dbRow;
            }
          }
        }
      }

      // Remove only the rows that were successfully synced.
      for (const code of dirtyCodes) {
        dirtyCodesRef.current.delete(code);
      }

      setSaveStatus('Cloud Synced ✓');

      window.setTimeout(() => {
        if (
          !autoSyncInFlightRef.current &&
          dirtyCodesRef.current.size === 0
        ) {
          setSaveStatus('Cloud Ready');
        }
      }, 1500);
    } catch (error) {
      console.error(
        'Supabase 10-minute auto-sync error:',
        error
      );

      console.error(
        'Supabase 10-minute sync details:',
        {
          code: error?.code,
          message: error?.message,
          details: error?.details,
          hint: error?.hint,
        }
      );

      // Keep dirty rows. They will be retried at the next 10-minute cycle.
      setSaveStatus(
        'Cloud Failed • Local Saved • Will retry'
      );
    } finally {
      autoSyncInFlightRef.current = false;
    }
  };

  /* -----------------------------------------
     10-MINUTE TIMER

     No cloud request is made while typing.
     The interval checks once every 10 minutes.
  ----------------------------------------- */

  useEffect(() => {
    const timer =
      window.setInterval(() => {
        syncDirtyRows();
      }, AUTO_SYNC_INTERVAL);

    return () => {
      window.clearInterval(timer);
    };
  }, [rows]);

  /* -----------------------------------------
     SAVE DATA
  ----------------------------------------- */

  const saveData = async () => {
    if (rows.length === 0) {
      return;
    }

    setSaveStatus(
      'Saving...'
    );

    const dataToSave = [];

    for (
      let rowIndex = 0;
      rowIndex < rows.length;
      rowIndex++
    ) {
      const code = rows[rowIndex].code;

      const savedRow =
        readRowFromDom(rowIndex);

      if (savedRow) {
        dataToSave.push(savedRow);
      }
    }

    /* ---------------------------------------
       1. SAVE LOCAL CACHE FIRST
    --------------------------------------- */

    try {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify(dataToSave)
      );
    } catch (error) {
      console.error(
        'Local save error:',
        error
      );
    }

    /* ---------------------------------------
       2. PREPARE SUPABASE PAYLOAD
    --------------------------------------- */

    const existingRows = [];
    const newRows = [];

    for (const row of dataToSave) {
      const existingDatabaseRow =
        databaseRowsRef.current[row.code];

      const payload =
        websiteRowToDatabaseRow(
          row,
          existingDatabaseRow?.id ?? null
        );

      if (existingDatabaseRow?.id) {
        existingRows.push(payload);
      } else {
        newRows.push(payload);
      }
    }

    try {
      /* -------------------------------------
         UPDATE EXISTING RECORDS
         ONE BULK REQUEST
      ------------------------------------- */

      if (existingRows.length > 0) {
        /*
         * Bulk update existing records by primary-key id.
         * We do not request returned rows here because the
         * existing database IDs are already known from the
         * initial Supabase load.
         */
        const { error } =
          await supabase
            .from('master_cashbook')
            .upsert(existingRows);

        if (error) {
          throw error;
        }
      }

      /* -------------------------------------
         INSERT NEW RECORDS
         ONE BULK REQUEST
      ------------------------------------- */

      if (newRows.length > 0) {
        const { data, error } =
          await supabase
            .from('master_cashbook')
            .insert(newRows)
            .select('*');

        if (error) {
          throw error;
        }

        if (Array.isArray(data)) {
          for (const dbRow of data) {
            if (dbRow?.code) {
              databaseRowsRef.current[
                dbRow.code
              ] = dbRow;
            }
          }
        }
      }

      dirtyCodesRef.current.clear();

      setSaveStatus(
        'Cloud Saved ✓'
      );

      window.setTimeout(() => {
        setSaveStatus('Cloud Ready');
      }, 1800);
    } catch (error) {
      console.error(
        'Supabase save error:',
        error
      );

      console.error(
        'Supabase error details:',
        {
          code: error?.code,
          message: error?.message,
          details: error?.details,
          hint: error?.hint,
        }
      );

      setSaveStatus(
        `Cloud Failed • ${error?.message || 'Save error'}`
      );
    }
  };

  const visibleRows =
    cashierFilter === 'cashier1'
      ? rows.slice(0, 54)
      : cashierFilter === 'cashier2'
        ? rows.slice(54, 107)
        : rows;

  /* -----------------------------------------
     LOADING SCREEN
  ----------------------------------------- */

  if (loading) {
    return (
      <main className="cashbook-page">
        <header className="cashbook-topbar">
          <div className="brand-box">

            <div className="brand-icon">
              ₹
            </div>

            <div>
              <div className="brand-title">
                HAPPI MOBILES
              </div>

              <div className="brand-subtitle">
                MASTER CASHBOOK
              </div>
            </div>

          </div>

          <div className="cashbook-actions">
            <div className="status-box">
              <span className="status-dot" />
              Connecting to Cloud...
            </div>
          </div>
        </header>

        <section
          className="cashbook-grid-wrap"
          style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            minHeight: '300px',
            fontWeight: 600,
          }}
        >
          Loading Master Cashbook...
        </section>
      </main>
    );
  }

  /* -----------------------------------------
     UI
  ----------------------------------------- */

  return (
    <main className="cashbook-page">

      {/* TOP BAR */}

      <header className="cashbook-topbar">

        <div className="brand-box">

          <div className="brand-icon">
            ₹
          </div>

          <div>

            <div className="brand-title">
              HAPPI MOBILES
            </div>

            <div className="brand-subtitle">
              MASTER CASHBOOK
            </div>

          </div>

        </div>

        <div className="cashbook-actions">

          <button
            type="button"
            className="save-button"
            onClick={saveData}
          >
            <span className="save-icon">
              💾
            </span>

            Save Data
          </button>

          <div className="status-box">

            <span
              className={
                saveStatus ===
                'Unsaved changes'
                  ? 'status-dot unsaved'
                  : 'status-dot'
              }
            />

            {saveStatus}

          </div>

        </div>

      </header>

      {/* FEATURE TOOLBAR */}

      <div
        style={{
          flex: '0 0 auto',
          display: 'flex',
          alignItems: 'center',
          gap: '6px',
          flexWrap: 'wrap',
          padding: '6px 10px',
          background: 'linear-gradient(90deg,#f7fbff 0%,#eef6ff 100%)',
          borderBottom: '1px solid #c7d2dc',
          boxShadow: '0 1px 2px rgba(0,0,0,0.05)',
          zIndex: 90,
        }}
      >
        <div
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: '7px',
            marginRight: '2px',
            padding: '0 2px',
          }}
        >
          <span style={{ fontSize: '12px', fontWeight: 800, whiteSpace: 'nowrap' }}>
            👤 Cashier
          </span>
          <select
            value={cashierFilter}
            onChange={(event) => setCashierFilter(event.target.value)}
            title="Display filter only. Cloud data is shared."
            style={{
              height: '34px',
              minWidth: '138px',
              border: '1px solid #cbd5e1',
              borderRadius: '8px',
              padding: '0 10px',
              background: '#fff',
              fontWeight: 700,
              cursor: 'pointer',
            }}
          >
            <option value="all">All Stores (107)</option>
            <option value="cashier1">Cashier 1 (54)</option>
            <option value="cashier2">Cashier 2 (53)</option>
          </select>
        </div>

        <button
          type="button"
          onClick={openOpeningUpload}
          style={featureButtonStyle('#1476bd')}
        >
          📤 Upload Opening Balance
        </button>

        <button
          type="button"
          onClick={openAddinsUpload}
          style={featureButtonStyle('#238b63')}
        >
          📤 Upload Addins
        </button>

        <button
          type="button"
          onClick={() =>
            confirmAndClearMasterField(
              'opening',
              OPENING_UNDO_KEY,
              'Opening Balances'
            )
          }
          style={featureButtonStyle('#d97706')}
        >
          🧹 Clear Opening Balances
        </button>

        {lastClearField === OPENING_UNDO_KEY &&
          undoSnapshotsRef.current[OPENING_UNDO_KEY] && (
            <button
              type="button"
              onClick={() =>
                undoClearMasterField(
                  'opening',
                  OPENING_UNDO_KEY,
                  'Opening Balances'
                )
              }
              style={undoButtonStyle}
            >
              ↶ Undo Opening
            </button>
          )}

        <button
          type="button"
          onClick={() =>
            confirmAndClearMasterField(
              'deposit',
              DEPOSIT_UNDO_KEY,
              'Deposits'
            )
          }
          style={featureButtonStyle('#c2410c')}
        >
          🧹 Clear Deposits
        </button>

        {lastClearField === DEPOSIT_UNDO_KEY &&
          undoSnapshotsRef.current[DEPOSIT_UNDO_KEY] && (
            <button
              type="button"
              onClick={() =>
                undoClearMasterField(
                  'deposit',
                  DEPOSIT_UNDO_KEY,
                  'Deposits'
                )
              }
              style={undoButtonStyle}
            >
              ↶ Undo Deposits
            </button>
          )}

        <button
          type="button"
          onClick={() =>
            confirmAndClearMasterField(
              'addings',
              ADDINS_UNDO_KEY,
              'Addins'
            )
          }
          style={featureButtonStyle('#be123c')}
        >
          🧹 Clear Addins
        </button>

        {lastClearField === ADDINS_UNDO_KEY &&
          undoSnapshotsRef.current[ADDINS_UNDO_KEY] && (
            <button
              type="button"
              onClick={() =>
                undoClearMasterField(
                  'addings',
                  ADDINS_UNDO_KEY,
                  'Addins'
                )
              }
              style={undoButtonStyle}
            >
              ↶ Undo Addins
            </button>
          )}

        <button
          type="button"
          onClick={() =>
            confirmAndClearMasterField(
              'denomination',
              DENOMINATION_UNDO_KEY,
              'Denomination'
            )
          }
          style={featureButtonStyle('#9f1239')}
        >
          🧹 Clear Denomination
        </button>

        {lastClearField === DENOMINATION_UNDO_KEY &&
          undoSnapshotsRef.current[DENOMINATION_UNDO_KEY] && (
            <button
              type="button"
              onClick={() =>
                undoClearMasterField(
                  'denomination',
                  DENOMINATION_UNDO_KEY,
                  'Denomination'
                )
              }
              style={undoButtonStyle}
            >
              ↶ Undo Denomination
            </button>
          )}

        <button
          type="button"
          onClick={openFinanceReport}
          style={featureButtonStyle('#7c3aed')}
        >
          📊 Finance Report ({getReportStoreCount('finance')})
        </button>

        <button
          type="button"
          onClick={openSRReport}
          style={featureButtonStyle('#0f766e')}
        >
          📋 SR Report ({getReportStoreCount('sr')})
        </button>

        <button
          type="button"
          onClick={openPendingApprovals}
          style={featureButtonStyle('#0369a1')}
        >
          📌 Pending Approvals ({getReportStoreCount('pendingApprovals')})
        </button>

        <button
          type="button"
          onClick={openEdits}
          style={featureButtonStyle('#7e22ce')}
        >
          ✏️ Edits ({getReportStoreCount('edits')})
        </button>

        <button
          type="button"
          onClick={downloadContraSheet}
          style={featureButtonStyle('#475569')}
        >
          ⬇ Contra Sheet
        </button>

        <button
          type="button"
          onClick={() => setPendingStatusOpen(true)}
          style={featureButtonStyle('#b45309')}
        >
          📬 Pending Status ({getPendingStatusStoreCount()})
        </button>

        <button
          type="button"
          onClick={openLowCashReport}
          style={featureButtonStyle('#d97706')}
        >
          ⚠ Low / No Cash ({getLowNoCashStoreCount()})
        </button>

        <button
          type="button"
          onClick={downloadMasterCashbookExcel}
          style={featureButtonStyle('#166534')}
        >
          📥 Master Cash Book Excel
        </button>

        <input
          ref={openingFileInputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          style={{ display: 'none' }}
          onChange={(event) =>
            uploadExcelFile(
              event.target.files?.[0],
              'opening'
            )
          }
        />

        <input
          ref={addinsFileInputRef}
          type="file"
          accept=".xlsx,.xls,.csv"
          style={{ display: 'none' }}
          onChange={(event) =>
            uploadExcelFile(
              event.target.files?.[0],
              'addins'
            )
          }
        />
      </div>

      {/* GRID */}

      <section className="cashbook-grid-wrap">

        <table className="cashbook-table">

          <colgroup>

            <col className="col-sl" />
            <col className="col-code" />
            <col className="col-branch" />

            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />
            <col className="col-number" />

            <col className="col-closing" />
            <col className="col-remarks" />

          </colgroup>

          <thead>

            <tr>

              {HEADERS.map(
                (header) => (
                  <th key={header}>
                    {header}
                  </th>
                )
              )}

            </tr>

          </thead>

          <tbody>

            {visibleRows.map(
              (row) => {
                const rowIndex = rows.findIndex(
                  (item) => item.code === row.code
                );

                return (

                <tr
                  key={row.code}
                  ref={(element) => {
                    rowRefs.current[
                      rowIndex
                    ] = element;
                  }}
                >

                  <td className="fixed-cell sl-cell">
                    {row.slNo}
                  </td>

                  <td className="fixed-cell code-cell">
                    {row.code}
                  </td>

                  <td className="fixed-cell branch-cell">
                    {row.branch}
                  </td>

                  {NUMBER_FIELDS.map(
                    (field) => (

                      <td
                        key={field}
                        className="amount-cell"
                      >

                        <input
                          type="text"
                          inputMode="decimal"
                          autoComplete="off"
                          spellCheck="false"
                          readOnly
                          style={{
                            caretColor: 'currentColor',
                          }}
                          defaultValue={getDisplayInputValue(row[field])}
                          data-raw-value={row[field] ?? ''}
                          data-field={field}
                          data-number-field="true"
                          onKeyDown={(event) =>
                            handleNumberKeyDown(
                              event,
                              rowIndex,
                              field
                            )
                          }
                          onDoubleClick={(event) =>
                            handleNumberDoubleClick(
                              event,
                              rowIndex,
                              field
                            )
                          }
                          onFocus={handleNumberFocus}
                          onInput={(event) =>
                            handleNumberInput(
                              event,
                              rowIndex
                            )
                          }
                          onBlur={(event) =>
                            handleNumberBlur(
                              event,
                              rowIndex
                            )
                          }
                        />

                      </td>

                    )
                  )}

                  <td className="closing-cell">

                    <span className="closing-value">
                      {formatAmount(
                        calculateClosing(
                          row
                        )
                      )}
                    </span>

                  </td>

                  <td className="remarks-cell">

                    <input
                      type="text"
                      autoComplete="off"
                      spellCheck="false"
                      defaultValue={
                        row.remarks
                      }
                      className="remarks-input"
                      onInput={() =>
                        markRowDirty(
                          rowIndex
                        )
                      }
                    />

                  </td>

                </tr>

                );
              }
            )}

          </tbody>

        </table>

      </section>

      {/* SR REPORT MODAL */}

      {srReportOpen && (
        <div style={modalOverlayStyle}>
          <div style={financeModalStyle}>
            <div style={modalHeaderStyle}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontSize: '16px' }}>📋</span>
                <strong style={{ color: '#263746', fontSize: '14px' }}>
                  SRN Pending Stores
                </strong>
              </div>
              <button
                type="button"
                onClick={() => setSrReportOpen(false)}
                style={closeButtonStyle}
                aria-label="Close SR Report"
              >
                ×
              </button>
            </div>

            <div style={financeTableWrapStyle}>
              <table style={financeTableStyle}>
                <thead>
                  <tr>
                    <th style={financeThStyle}>Sl.no</th>
                    <th style={financeThStyle}>Branch</th>
                    <th style={financeThStyle}>Amount</th>
                    <th style={financeThStyle}>Bill No.</th>
                    <th style={financeThStyle}>Remarks</th>
                    <th style={financeThStyle}>Date</th>
                    <th style={financeThStyle}>Bill Date</th>
                    <th style={financeThStyle}>Ageing</th>
                  </tr>
                </thead>
                <tbody>
                  {srReportRows.map((item) => (
                    <tr key={item.id}>
                      <td style={financeTdCenterStyle}>{item.slNo}</td>
                      <td style={financeTdStyle}>{item.branch}</td>
                      <td style={financeAmountTdStyle}>₹{formatAmount(item.amount)}</td>
                      <td style={financeTdStyle}>
                        <input
                          type="text"
                          placeholder="Bill No..."
                          value={item.billNo}
                          onChange={(event) =>
                            updateSRReportRow(item.id, {
                              billNo: event.target.value,
                            })
                          }
                          style={financeInputStyle}
                        />
                      </td>
                      <td style={financeTdStyle}>
                        <input
                          type="text"
                          value={item.remarks}
                          onChange={(event) =>
                            updateSRReportRow(item.id, {
                              remarks: event.target.value,
                            })
                          }
                          style={financeInputStyle}
                        />
                      </td>
                      <td style={financeTdCenterStyle}>
                        <input
                          type="date"
                          value={item.date}
                          readOnly
                          style={{
                            ...financeDateInputStyle,
                            background: '#f5f7fa',
                            cursor: 'not-allowed',
                          }}
                        />
                      </td>
                      <td style={financeTdCenterStyle}>
                        <input
                          type="date"
                          value={item.billDate}
                          onChange={(event) =>
                            updateSRReportRow(item.id, {
                              billDate: event.target.value,
                            })
                          }
                          style={financeDateInputStyle}
                        />
                      </td>
                      <td style={{ ...financeTdCenterStyle, fontWeight: 800, color: '#c62828' }}>
                        {calculateAgeing(srToday, item.billDate)}
                      </td>
                    </tr>
                  ))}

                  {srReportRows.length === 0 && (
                    <tr>
                      <td colSpan={8} style={{ padding: '20px', textAlign: 'center', color: '#64748b', fontWeight: 600 }}>
                        No SR amount is pending in the Master Cashbook.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div style={financeFooterStyle}>
              <div style={{ fontSize: '11px', color: '#5f6f7d' }}>
                {clipboardNotice || `${srReportRows.length} pending branch${srReportRows.length === 1 ? '' : 'es'}. Outlook body will stay blank.`}
              </div>
              <div style={{ display: 'flex', gap: '7px' }}>
                <button
                  type="button"
                  onClick={downloadSRExcel}
                  style={financeActionButtonStyle('#16803c')}
                >
                  ⬇ Download Excel
                </button>
                <button
                  type="button"
                  onClick={sendSRMail}
                  style={financeActionButtonStyle('#2867d8')}
                >
                  ✉ Send Mail (Outlook)
                </button>
                <button
                  type="button"
                  onClick={() => setSrReportOpen(false)}
                  style={financeCloseActionStyle}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* PENDING APPROVALS MODAL */}

      {pendingApprovalsOpen && (
        <div style={modalOverlayStyle}>
          <div style={simpleListModalStyle}>
            <div style={modalHeaderStyle}>
              <strong style={{ color: '#263746', fontSize: '13px' }}>
                📌 Pending Approvals List
              </strong>
              <button
                type="button"
                onClick={() => setPendingApprovalsOpen(false)}
                style={closeButtonStyle}
              >
                ×
              </button>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', padding: '5px 8px', borderBottom: '1px solid #d9e1e8' }}>
              <button
                type="button"
                onClick={() =>
                  downloadSimpleListExcel(
                    'Pending Approvals',
                    'Pending Approvals',
                    'Pending Approvals.xlsx',
                    'pendingApprovals'
                  )
                }
                style={financeActionButtonStyle('#16803c')}
              >
                ⬇ Excel
              </button>
            </div>

            <div style={reportListStyle}>
              {buildPendingApprovalList().map((item, index) => (
                <div
                  key={`${item.code}-${index}`}
                  style={{
                    ...reportListRowStyle,
                    gridTemplateColumns: '72px 1fr 105px',
                    alignItems: 'center',
                  }}
                >
                  <div style={reportCodeStyle}>
                    {item.code}
                  </div>

                  <div style={reportBranchStyle}>
                    {item.branch}
                  </div>

                  <div
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: '9px',
                      fontWeight: 700,
                      color: '#0066b3',
                    }}
                  >
                    <span>
                      ₹{formatAmount(item.amount)}
                    </span>

                    <input
                      type="checkbox"
                      aria-label={`Mark ${item.branch} ₹${formatAmount(item.amount)}`}
                      style={{
                        width: '17px',
                        height: '17px',
                        margin: 0,
                        cursor: 'pointer',
                        accentColor: '#15803d',
                      }}
                    />
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* EDITS MODAL */}

      {editsOpen && (
        <div style={modalOverlayStyle}>
          <div style={simpleListModalStyle}>
            <div style={modalHeaderStyle}>
              <strong style={{ color: '#263746', fontSize: '13px' }}>
                ✏️ Edits List
              </strong>
              <button
                type="button"
                onClick={() => setEditsOpen(false)}
                style={closeButtonStyle}
              >
                ×
              </button>
            </div>

            <div style={{ display: 'flex', justifyContent: 'flex-end', padding: '5px 8px', borderBottom: '1px solid #d9e1e8' }}>
              <button
                type="button"
                onClick={() =>
                  downloadSimpleListExcel(
                    'Edits',
                    'Edits',
                    'Edits.xlsx',
                    'edits'
                  )
                }
                style={financeActionButtonStyle('#16803c')}
              >
                ⬇ Excel
              </button>
            </div>

            <div style={reportListStyle}>
              {buildAmountList('edits').map((item) => (
                <div key={item.code} style={reportListRowStyle}>
                  <div style={reportCodeStyle}>{item.code}</div>
                  <div style={reportBranchStyle}>{item.branch}</div>
                  <div style={reportAmountStyle}>₹{formatAmount(item.amount)}</div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* PENDING STATUS MODAL */}

      {pendingStatusOpen && (
        <div style={modalOverlayStyle}>
          <div style={twoPaneModalStyle}>
            <div style={modalHeaderStyle}>
              <strong style={{ color: '#263746', fontSize: '14px' }}>
                Pending Status
              </strong>
              <button
                type="button"
                onClick={() => setPendingStatusOpen(false)}
                style={closeButtonStyle}
              >
                ×
              </button>
            </div>

            <div style={twoPaneGridStyle}>
              <section style={paneStyle}>
                <div style={paneTitleStyle}>
                  <span>📌 Pending Cashbooks</span>
                  <span>{getPendingCashbooks().length}</span>
                </div>
                <div style={reportListStyle}>
                  {getPendingCashbooks().map((row) => (
                    <div key={row.code} style={{ ...reportListRowStyle, gridTemplateColumns: '72px 1fr' }}>
                      <div style={reportCodeStyle}>{row.code}</div>
                      <div style={reportBranchStyle}>{row.branch}</div>
                    </div>
                  ))}
                </div>
              </section>

              <section style={paneLastStyle}>
                <div style={paneTitleStyle}>
                  <span>📄 Pending Deposit Slips (Denom Amt)</span>
                  <span>{getPendingDepositSlips().length}</span>
                </div>
                <div style={reportListStyle}>
                  {getPendingDepositSlips().map((row) => (
                    <div key={row.code} style={reportListRowStyle}>
                      <div style={reportCodeStyle}>{row.code}</div>
                      <div style={reportBranchStyle}>{row.branch}</div>
                      <div style={reportAmountStyle}>₹{formatAmount(calculateAmount(row.denomination))}</div>
                    </div>
                  ))}
                </div>
              </section>
            </div>
          </div>
        </div>
      )}

      {/* LOW / NO CASH MODAL */}

      {lowCashOpen && (
        <div style={modalOverlayStyle}>
          <div style={financeModalStyle}>
            <div style={modalHeaderStyle}>
              <strong style={{ color: '#263746', fontSize: '14px' }}>
                ⚠ Low Cash &amp; No Cash Report
              </strong>
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <button
                  type="button"
                  onClick={downloadLowCashExcel}
                  style={financeActionButtonStyle('#16803c')}
                >
                  ⬇ Download Excel
                </button>
                <button
                  type="button"
                  onClick={() => setLowCashOpen(false)}
                  style={closeButtonStyle}
                >
                  ×
                </button>
              </div>
            </div>

            <div style={{ padding: '7px 10px', fontSize: '11px', color: '#5f6f7d', borderBottom: '1px solid #d9e1e8' }}>
              Blank / 0 = No Cash • ₹1–₹500 = Low Cash • ₹501+ = blank editable remarks.
            </div>

            <div style={{ overflow: 'auto', padding: '8px 10px 12px' }}>
              <table style={lowCashTableStyle}>
                <thead>
                  <tr>
                    <th style={financeThStyle}>Store Name</th>
                    <th style={financeThStyle}>Denom Amt</th>
                    <th style={financeThStyle}>Status / Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  {lowCashRows.map((item) => (
                    <tr key={item.code}>
                      <td style={financeTdStyle}>{item.branch}</td>
                      <td style={financeAmountTdStyle}>₹{formatAmount(item.denomination)}</td>
                      <td style={financeTdStyle}>
                        <input
                          type="text"
                          value={item.status}
                          placeholder="Type remarks..."
                          onChange={(event) =>
                            updateLowCashStatus(
                              item.code,
                              event.target.value
                            )
                          }
                          style={{
                            ...(getLowCashDefaultRemark({
                              denomination: item.denomination,
                            })
                              ? lowCashDefaultStyle
                              : lowCashInputStyle),
                            background:
                              item.denomination <= 0
                                ? '#ffdada'
                                : item.denomination <= 500
                                  ? '#fff0c2'
                                  : '#ffffff',
                          }}
                        />
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}

      {/* FINANCE REPORT MODAL */}

      {financeReportOpen && (
        <div style={modalOverlayStyle}>
          <div style={financeModalStyle}>
            <div style={modalHeaderStyle}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <span style={{ fontSize: '16px' }}>💳</span>
                <strong style={{ color: '#263746', fontSize: '14px' }}>
                  Finance and Card Amount Pending in Cash Book
                </strong>
              </div>

              <button
                type="button"
                onClick={() => setFinanceReportOpen(false)}
                aria-label="Close Finance Report"
                style={closeButtonStyle}
              >
                ×
              </button>
            </div>

            <div
              style={{
                padding: '7px 10px',
                fontSize: '11px',
                color: '#5f6f7d',
                borderBottom: '1px solid #d9e1e8',
              }}
            >
              Bill No. and Remarks are editable. Date is today and cannot be changed. Bill Date is editable with calendar. Ageing = Current Date − Bill Date. Outlook opens with a blank body; the report table is copied for Ctrl+V.
            </div>

            <div style={financeTableWrapStyle}>
              <table style={financeTableStyle}>
                <thead>
                  <tr>
                    <th style={financeThStyle}>Sl.no</th>
                    <th style={financeThStyle}>Branch</th>
                    <th style={financeThStyle}>Amount</th>
                    <th style={financeThStyle}>Bill No.</th>
                    <th style={financeThStyle}>Remarks</th>
                    <th style={financeThStyle}>Date</th>
                    <th style={financeThStyle}>Bill Date</th>
                    <th style={financeThStyle}>Ageing</th>
                  </tr>
                </thead>

                <tbody>
                  {financeReportRows.map((item) => (
                    <tr key={item.id}>
                      <td style={financeTdCenterStyle}>{item.slNo}</td>
                      <td style={financeTdStyle}>{item.branch}</td>
                      <td style={financeAmountTdStyle}>₹{formatAmount(item.amount)}</td>

                      <td style={financeTdStyle}>
                        <input
                          type="text"
                          placeholder="Bill No..."
                          value={item.billNo}
                          onChange={(event) =>
                            updateFinanceReportRow(item.id, {
                              billNo: event.target.value,
                            })
                          }
                          style={financeInputStyle}
                        />
                      </td>

                      <td style={financeTdStyle}>
                        <input
                          type="text"
                          value={item.remarks}
                          onChange={(event) =>
                            updateFinanceReportRow(item.id, {
                              remarks: event.target.value,
                            })
                          }
                          style={financeInputStyle}
                        />
                      </td>

                      <td style={financeTdCenterStyle}>
                        <input
                          type="date"
                          value={item.date}
                          readOnly
                          aria-label="Report date"
                          style={{
                            ...financeDateInputStyle,
                            background: '#f5f7fa',
                            cursor: 'not-allowed',
                          }}
                        />
                      </td>

                      <td style={financeTdCenterStyle}>
                        <input
                          type="date"
                          value={item.billDate}
                          onChange={(event) =>
                            updateFinanceReportRow(item.id, {
                              billDate: event.target.value,
                            })
                          }
                          style={financeDateInputStyle}
                        />
                      </td>

                      <td style={{ ...financeTdCenterStyle, fontWeight: 800, color: '#c62828' }}>
                        {calculateAgeing(financeToday, item.billDate)}
                      </td>
                    </tr>
                  ))}

                  {financeReportRows.length === 0 && (
                    <tr>
                      <td
                        colSpan={8}
                        style={{
                          padding: '20px',
                          textAlign: 'center',
                          color: '#64748b',
                          fontWeight: 600,
                        }}
                      >
                        No Finance Amount is pending in the Master Cashbook.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <div style={financeFooterStyle}>
              <div style={{ fontSize: '11px', color: '#5f6f7d' }}>
                {clipboardNotice || `${financeReportRows.length} pending branch${financeReportRows.length === 1 ? '' : 'es'}.`}
              </div>

              <div style={{ display: 'flex', gap: '7px' }}>
                <button
                  type="button"
                  onClick={downloadFinanceExcel}
                  style={financeActionButtonStyle('#16803c')}
                >
                  ⬇ Download Excel
                </button>

                <button
                  type="button"
                  onClick={sendFinanceMail}
                  style={financeActionButtonStyle('#2867d8')}
                >
                  ✉ Send Mail (Outlook)
                </button>

                <button
                  type="button"
                  onClick={() => setFinanceReportOpen(false)}
                  style={financeCloseActionStyle}
                >
                  Close
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* UPLOAD ERROR MODAL */}

      {uploadErrorOpen && (
        <div style={modalOverlayStyle}>
          <div style={{ ...errorModalStyle, maxWidth: '680px' }}>
            <div style={modalHeaderStyle}>
              <strong style={{ color: '#a61b1b', fontSize: '14px' }}>
                ⚠ {uploadErrorTitle}
              </strong>

              <button
                type="button"
                onClick={() => setUploadErrorOpen(false)}
                aria-label="Close Upload Errors"
                style={closeButtonStyle}
              >
                ×
              </button>
            </div>

            <div style={errorListStyle}>
              {uploadErrors.map((message, index) => (
                <div key={`${message}-${index}`} style={errorRowStyle}>
                  <span style={{ fontWeight: 800, color: '#b42318', marginRight: '6px' }}>
                    {index + 1}.
                  </span>
                  {message}
                </div>
              ))}
            </div>

            <div style={financeFooterStyle}>
              <span style={{ fontSize: '11px', color: '#667085' }}>
                Matched branches were updated. These entries were not matched/processed.
              </span>

              <button
                type="button"
                onClick={() => setUploadErrorOpen(false)}
                style={financeCloseActionStyle}
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}

    </main>
  );
}