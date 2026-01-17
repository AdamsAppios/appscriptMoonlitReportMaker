/*** NEW: numbered pastes for IN OUTS, Bills, Expenses (AM/PM) ***/
function _handleNumberedPastes_(e) {
  const sh = e.source.getActiveSheet();
  if (!sh || sh.getName() !== 'MoonlitReport') return;
  const a1 = e.range.getA1Notation();

  // B28 -> IN OUTS -> Column L (12)
  if (a1 === 'B28') {
    appendNumberedBlock_(sh, 12 /*L*/, 'B28');
    return;
  }
  // B29 -> Bills -> Column M (13)
  if (a1 === 'B29') {
    appendNumberedBlock_(sh, 13 /*M*/, 'B29');
    return;
  }
  // B30 -> Expenses AM -> Column J (10)
  if (a1 === 'B30') {
    appendNumberedBlock_(sh, 10 /*J*/, 'B30');
    return;
  }
  // C30 -> Expenses PM -> Column K (11)
  if (a1 === 'C30') {
    appendNumberedBlock_(sh, 11 /*K*/, 'C30');
    return;
  }
}

function onEdit(e) {
  const sh  = e.source.getActiveSheet();
  const a1  = e.range.getA1Notation();
  const val = typeof e.value === 'string' ? e.value.trim() : '';
  
  _handleNumberedPastes_(e);

  // ────────────────────────────────────────────────────────────────────────────
  // NEW: MoonlitDisplay checkbox triggers (do NOT touch your other logic)
  // ────────────────────────────────────────────────────────────────────────────
  if (sh.getName() === 'MoonlitDisplay') {
    // Run only when the checkbox becomes checked (TRUE)
    if (a1 === 'C2' && e.range.getValue() === true) {
      updateBillsDisplay(e);            // in MoonlitDisplay.gs
      return;
    }
    if (a1 === 'D2' && e.range.getValue() === true) {
      updateMoonlitAccounts(e);         // in MoonlitDisplay.gs
      return;
    }
    if (a1 === 'E2' && e.range.getValue() === true) {
      updateSalaryDisplay(e);           // in MoonlitDisplay.gs
      return;
    }
    // If some other cell on MoonlitDisplay was edited, do nothing.
    return;
  }

  // ────────────────────────────────────────────────────────────────────────────
  // Your existing MoonlitReport logic (unchanged)
  // ────────────────────────────────────────────────────────────────────────────
  const sheetName = 'MoonlitReport';
  if (sh.getName() !== sheetName) return;

  // NEW: SMS paste (AM/PM)
  if (a1 === 'B27' || a1 === 'C27') {
    if (val) {
      const isPM = (a1 === 'C27');
      parseAndFillSmsReport_(sh, val, isPM);   // in MoonlitReportSmsParser.gs
    }
    // Clear the paste cell; scripted edits don’t fire the simple onEdit trigger.
    sh.getRange(a1).clearContent();
    return;
  }

  // The rest is exactly what you already had
  const dateCell         = 'B1';
  const checkboxUpdate   = 'B5';
  const checkboxRetrieve = 'B3';
  const checkboxClear    = 'B4';
  const depositCheckbox  = 'Q1';

  // 1) UPDATE checkbox (B5)
  if (a1 === checkboxUpdate) {
    if (e.range.getValue()) {
      updateMoonlitData(e);
      updateStatus("updated on " + getPhilippineTime(), "#ADD8E6");
    }
    return;
  }

  // 2) Deposit checkbox (Q1)
  if (a1 === depositCheckbox) {
    if (e.range.getValue()) {
      calculateDeposit(sh, 'Q8');
    }
    return;
  }

  // 3) RETRIEVE checkbox (B3) OR changing the date (B1)
  if (a1 === checkboxRetrieve || a1 === dateCell) {
    clearCellsForNewDate(e.source);
    retrieveDataFromMoonlitData(e);
    return;
  }

  // 4) CLEAR checkbox (B4)
  if (a1 === checkboxClear) {
    clearCellsForNewDate(e.source);
    updateStatus("cleared on " + getPhilippineTime(), "#FFFF00");
    return;
  }


}
