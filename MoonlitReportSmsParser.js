/**
 * Parse whole SMS-like report and fill MoonlitReport.
 * - AM paste in B27 -> writes into B* cells, pullouts -> N
 * - PM paste in C27 -> writes into C* cells, pullouts -> O
 * - NEW: copy cashier to F8 (AM) or F14 (PM)
 * - NEW: copy “Total: ####” to F13 (AM) or F20 (PM)
 * - Mineral count goes to F5 only for PM (as you asked previously)
 */
function parseAndFillSmsReport_(sheet, raw, isPM) {
  const t = String(raw).replace(/\r/g, '').trim();

  // ---------------- Cashier ----------------
  // Example: "Cashier: Anna - 6pm"
  const cashierMatch =
    t.match(/Cashier\s*:\s*([^\n-]+?)\s*-\s*.*$/mi) ||
    t.match(/Cashier\s*:\s*([^\n]+)$/mi);
  if (cashierMatch) {
    const name = cashierMatch[1].trim();
    // Duty (left side)
    setCellString_(sheet, isPM ? 'C8' : 'B8', name, true);
    // NEW: copy to F8 (AM) / F14 (PM)
    setCellString_(sheet, isPM ? 'F14' : 'F8', name, true);
  }

  // ---------------- NEW: Total: #### ----------------
  // Example line: "Total: 1358"
  const totalMatch = t.match(/^Total\s*:\s*([\d,]+)/mi);
  if (totalMatch) {
    const totalNum = parseInt(totalMatch[1].replace(/,/g, ''), 10);
    sheet.getRange(isPM ? 'F20' : 'F13').setValue(totalNum);
  }

  // ---------------- Scalar fields (left columns) ----------------
  // Row map (A8:A20 labels): Duty, SD Sale, TSTD STK, NS Sale, NS Stk, SB, Coins,
  //                           Oil, PlasticSB, PlasticLoaf, #3, #6, Tiny
  const L = (addrAM, addrPM, forceText = false) =>
    (isPM ? { a1: addrPM, text: forceText } : { a1: addrAM, text: forceText });

  const picks = [
    { re: /SD\s*=\s*([^\n]+)/i,           ...L('B9',  'C9')  },
    { re: /Toasted\s*=\s*([^\n]+)/i,      ...L('B10', 'C10') }, // TSTD STK
    { re: /NSSale\s*=\s*([^\n]+)/i,       ...L('B11', 'C11') },
    { re: /NSStocks\s*=\s*([^\n]+)/i,     ...L('B12', 'C12') },
    { re: /SB\s*=\s*([^\n]+)/i,           ...L('B13', 'C13') },
    { re: /Coins\s*=\s*([^\n]+)/i,        ...L('B14', 'C14') },
    // keep as TEXT to preserve things like "16.4k 200"
    { re: /Mantika\s*=\s*([^\n]+)/i,      ...L('B15', 'C15', true) },
    { re: /Plastic\s*SB\s*=\s*([^\n]+)/i, ...L('B16', 'C16') },
    { re: /Plastic\s*Loaf\s*=\s*([^\n]+)/i,...L('B17', 'C17') },
    { re: /Loaf(?:\s*bread)?\s*=\s*([^\n]+)/i, ...L('B18', 'C18', true) },
    // keep as TEXT so "500+500" doesn’t get coerced
    { re: /plastic[_\s-]*No3\s*=\s*([^\n]+)/i, ...L('B19', 'C19', true) },
    { re: /plastic[_\s-]*No6\s*=\s*([^\n]+)/i, ...L('B20', 'C20', true) },
    { re: /plastic[_\s-]*Tiny\s*=\s*([^\n]+)/i,...L('B21', 'C21', true) },
    { re: /Plastic\s*Medium\s*=\s*([^\n]+)/i,     ...L('B22', 'C22') },
    { re: /Plastic\s*Large\s*=\s*([^\n]+)/i,      ...L('B23', 'C23') },
  ];

  picks.forEach(p => {
    const m = t.match(p.re);
    if (m) setCellString_(sheet, p.a1, m[1].trim(), !!p.text);
  });

  // ---------------- Mineral (PM only) ----------------
  if (isPM) {
    const m = t.match(/Mineral\s*=\s*([0-9]+)\s*x/i);
    if (m) sheet.getRange('F5').setValue(parseInt(m[1], 10));
  }

  // ---------------- Pullouts block ----------------
  // Between "Pullouts =" and next blank line OR "Accounts =" / "Workers =" / "Expensis" etc.
  const pullBlock =
    (t.match(/Pullouts\s*=\s*\n([\s\S]*?)(?:\n\s*\n|Accounts\s*=|Workers\s*=|Expensis|Expenses|^-End-)/i) || [])[1];

  // Clear old list then write the new one
  sheet.getRange(isPM ? 'O2:O200' : 'N2:N200').clearContent();
  if (pullBlock) {
    const lines = pullBlock
      .split('\n')
      .map(s => s.trim())
      .filter(Boolean);
    if (lines.length) {
      const start = isPM ? 'O2' : 'N2';
      sheet.getRange(start).offset(0, 0, lines.length, 1)
           .setValues(lines.map(l => [l]));
    }
  }
}

/** helper: write as string, optionally forcing "Plain text" format */
function setCellString_(sheet, a1, value, forceText) {
  const r = sheet.getRange(a1);
  if (forceText) r.setNumberFormat('@'); // keep things like "500+500" / "16.4k 200" as text
  r.setValue(value);
}


/**
 * Paste any multi-line IN OUTS text into B28.
 * Appends to column L starting at L2, leaving ONE blank row between blocks.
 * Clears B28 after processing. Scripted clears do NOT retrigger onEdit.
 */
function handlePasteToColumnCell_(sheet, raw, COL, START) {
  const lines = String(raw)
    .replace(/\r/g, '')
    .split('\n')
    .map(s => s.trim())
    .filter(Boolean);

  if (!lines.length) return;

  //const COL = 12;      // column L
  //const START = 2;       // start at L2

  // Find the last non-empty row in column L (from L2 down),
  // then start writing TWO rows below it to leave exactly one blank row.
  const lastRow = Math.max(sheet.getLastRow(), START);
  const values = sheet.getRange(START, COL, lastRow - START + 1, 1).getValues();

  let lastFilledRow = 0; // 0 means no content found
  for (let i = values.length - 1; i >= 0; i--) {
    if (String(values[i][0]).trim() !== '') {
      lastFilledRow = START + i;
      break;
    }
  }

  const writeRow = lastFilledRow ? (lastFilledRow + 2) : START; // +2 = one blank line gap
  sheet.getRange(writeRow, COL, lines.length, 1).setValues(lines.map(l => [l]));
}

/**
 * Append a numbered block to a target column.
 * - Starts at 1. if column has no numbers yet
 * - Else next = (max existing number + 1)
 * - Adds one blank line between blocks if needed
 * - Splits text by newline into row-per-line
 */
function appendNumberedBlock_(sheet, targetCol, text) {
  const firstDataRow = 2; // your examples start at row 2 (L2, M2, etc.)
  const lastRow = Math.max(sheet.getLastRow(), firstDataRow);
  const numRows = Math.max(lastRow - firstDataRow + 1, 1);

  // Read current column (from row 2 down)
  const colVals = sheet.getRange(firstDataRow, targetCol, numRows, 1)
    .getValues()
    .map(r => String(r[0] ?? '').trim());

  // Find last non-empty row index in this column
  let lastIdx = colVals.length - 1;
  while (lastIdx >= 0 && colVals[lastIdx] === '') lastIdx--;

  // Determine next number = highest "^\d+\.$" found + 1 (or 1 if none)
  let nextNum = 1;
  for (let i = 0; i <= lastIdx; i++) {
    const m = colVals[i].match(/^(\d+)\.\s*$/);
    if (m) nextNum = Math.max(nextNum, parseInt(m[1], 10) + 1);
  }

  // Compute the row to start writing the new block
  let writeRow = firstDataRow + lastIdx + 1; // next free row
  // Ensure exactly one blank line before a new block if prior row isn't blank
  if (lastIdx >= 0 && colVals[lastIdx] !== '') {
    sheet.getRange(writeRow, targetCol).setValue('');
    writeRow += 1;
  }

  // Build output: ["N."], then each non-empty line
  const lines = String(text).split(/\r?\n/).filter(s => s.trim() !== '');
  const out = [[nextNum + '.']];
  lines.forEach(line => out.push([line]));

  sheet.getRange(writeRow, targetCol, out.length, 1).setValues(out);
}

/**
 * INTERNAL: scan a column for the last header number.
 * Accepts both "N." and "N. some text" (or even bare numeric N).
 * Returns { lastNonEmptyIdx, lastHeaderNum } relative to the scanned range startRow.
 */
function _scanNumberedColumn_(sheet, colIndex, startRow) {
  const lastRow = Math.max(sheet.getLastRow(), startRow);
  const readRows = Math.max(lastRow - startRow + 1, 1);
  const vals = sheet.getRange(startRow, colIndex, readRows, 1).getValues();

  let lastNonEmptyIdx = -1;
  let lastHeaderNum = 0;

  for (let i = 0; i < vals.length; i++) {
    const raw = vals[i][0];
    const s = (raw == null ? '' : String(raw)).trim();
    if (s !== '') lastNonEmptyIdx = i;

    // Match "12." OR "12. text" OR bare "12"
    const m = s.match(/^\s*(\d+)\.(?:\s|$)/) || s.match(/^\s*(\d+)\s*$/);
    if (m) lastHeaderNum = Math.max(lastHeaderNum, parseInt(m[1], 10));
    if (typeof raw === 'number' && Number.isFinite(raw)) {
      lastHeaderNum = Math.max(lastHeaderNum, Math.floor(raw));
    }
  }
  return { lastNonEmptyIdx, lastHeaderNum };
}

/**
 * Append a numbered block to a target column.
 * - Pastes the first line WITH the header on the SAME row: "N. <first line>"
 * - Leaves exactly ONE blank row between blocks (if the column already has text)
 * - If inputA1 === 'C30' (Expenses PM → column K), and column K is empty,
 *   this will CONTINUE the numbering from column J (Expenses AM) if present.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @param {number} targetColIndex  1-based column index (L=12, M=13, J=10, K=11, ...)
 * @param {string}  inputA1        A1 address of the input cell (e.g. 'B28','B29','B30','C30')
 */
function appendNumberedBlock_(sheet, targetColIndex, inputA1) {
  const text = String(sheet.getRange(inputA1).getValue() || '').trim();
  if (!text) return;

  // Split into lines (keep order; drop purely blank lines at ends)
  const lines = text.split(/\r?\n/).map(s => s.trim()).filter((s, i, arr) => {
    if (s !== '') return true;
    // keep inner blanks, but drop leading/trailing empties
    const head = arr.findIndex(x => x.trim() !== '');
    const tail = arr.length - 1 - [...arr].reverse().findIndex(x => x.trim() !== '');
    return i > head && i < tail;
  });
  if (!lines.length) return;

  const firstDataRow = 2;

  // 1) Scan the target column
  const { lastNonEmptyIdx, lastHeaderNum } =
    _scanNumberedColumn_(sheet, targetColIndex, firstDataRow);

  // 2) Special rule: Expenses PM (inputA1 === 'C30' → targetColIndex K=11)
  //    If K is empty, continue from J (Expenses AM).
  let nextNumBase = lastHeaderNum;
  if (inputA1 === 'C30' && lastNonEmptyIdx < 0) {
    const AM = _scanNumberedColumn_(sheet, 10 /* J */, firstDataRow);
    nextNumBase = Math.max(nextNumBase, AM.lastHeaderNum);
  }

  const nextNum = Math.max(1, nextNumBase + 1);

  // 3) Decide the first row where we’ll write
  let insertRow = firstDataRow;
  const out = [];

  if (lastNonEmptyIdx >= 0) {
    insertRow = firstDataRow + lastNonEmptyIdx + 1; // next free row
    // ensure exactly one blank spacer between blocks if the last row wasn't blank
    const lastVal = sheet.getRange(insertRow - 1, targetColIndex).getDisplayValue().trim();
    if (lastVal !== '') out.push(['']); // spacer
  }

  // 4) Build output
  // FIRST row = "N. <first-line>"
  out.push([`${nextNum}. ${lines[0]}`]);
  // Subsequent rows = remaining lines (no number prefix)
  for (let i = 1; i < lines.length; i++) out.push([lines[i]]);

  // Force plain text so "1. " stays as-is
  const r = sheet.getRange(insertRow, targetColIndex, out.length, 1);
  r.setNumberFormat('@');
  r.setValues(out);

  // Clear the input cell (scripted clears won't re-trigger simple onEdit)
  sheet.getRange(inputA1).clearContent();
}

/**
 * Dispatcher for MoonlitReport sheet inputs:
 *  - B28 → Column L (12)  IN OUTS
 *  - B29 → Column M (13)  Bills
 *  - B30 → Column J (10)  Expenses AM
 *  - C30 → Column K (11)  Expenses PM (continues numbering from J if J already has items)
 */
function handleNumberedInputs_(e) {
  const sh = e.source.getActiveSheet();
  if (!sh || sh.getName() !== 'MoonlitReport') return;
  if (!e.range) return;
  const a1 = e.range.getA1Notation();
  const val = e.value == null ? '' : String(e.value).trim();
  if (!val) return;

  // Exact inputs we support
  if (a1 === 'B28') {               // IN OUTS → L
    appendNumberedBlock_(sh, 12, 'B28');
  } else if (a1 === 'B29') {        // Bills → M
    appendNumberedBlock_(sh, 13, 'B29');
  } else if (a1 === 'B30') {        // Expenses AM → J
    appendNumberedBlock_(sh, 10, 'B30');
  } else if (a1 === 'C30') {        // Expenses PM → K (continue from J)
    appendNumberedBlock_(sh, 11, 'C30');
  }
}
