
/***********************************************************
 *  Code.js
 ***********************************************************/

/**
 * Returns the current date/time as a formatted string in Philippine time.
 * Format: "h:mma MMMM d,yyyy" in lower-case.
 */
function getPhilippineTime() {
  // Using Asia/Manila timezone
  let now = new Date();
  let formatted = Utilities.formatDate(now, "Asia/Manila", "h:mma MMMM d,yyyy");
  return formatted.toLowerCase();
}

/**
 * Writes a status message to cell B2 of MoonlitReport with the specified font color.
 */
function updateStatus(message, fontColor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("MoonlitReport");
  const statusCell = sheet.getRange("B2");
  statusCell.setValue(message);
  statusCell.setFontColor(fontColor);
}

/**
 * Searches for a matching date in MoonlitData (column A). 
 * If not found, automatically creates a new row with default values.
 * Returns the row number.
 */
function findOrCreateDateRow(dataSheet, dateValue) {
  const dateColumn = dataSheet.getRange('A:A').getValues();
  for (let i = 0; i < dateColumn.length; i++) {
    const currentDate = new Date(dateColumn[i][0]);
    if (currentDate.getTime() === dateValue.getTime()) {
      return i + 1;
    }
  }
  // Not found: create new row
  let newRow = dataSheet.getLastRow() + 1;
  dataSheet.getRange(newRow, 1).setValue(dateValue);
  // Set default values:
  const defaultDValue = 'expensesam=""; expensespm=""; inouts=""; bills=""; pulloutsam=""; pulloutspm="";';
  dataSheet.getRange(newRow, 4).setValue(defaultDValue);
  dataSheet.getRange(newRow, 6).setValue('duties=""');
  const defaultHValue =
    'totaldeposit="Sales AM: \\nSales PM: \\nCoffee Sales: \\nNS:\\n\\nTotal Sales:\\n\\nTotal Previous:\\nExpenses:\\nGrand Net Total:\\nDeposited (Y or N):"; absences=""; accountsam=""; accountspm=""; coffeesales="Sales: \\nCupsEnd: \\nCoffee: \\nChoco: \\nCaramel: \\nCappuccino: \\n3in1:"';
  dataSheet.getRange(newRow, 8).setValue(defaultHValue);
  return newRow;
}

/**
 * The UPDATED function called when B5 is toggled.
 * It pushes data from MoonlitReport into the matching row of MoonlitData.
 * If the date is not found, a new row is automatically created.
 */
function updateMoonlitData(e) {
  const sheetName = 'MoonlitReport';
  const dataSheetName = 'MoonlitData';
  const dateCell = 'B1';
  
  const sheet = e.source.getSheetByName(sheetName);
  const dateValue = new Date(sheet.getRange(dateCell).getValue());
  const dataSheet = e.source.getSheetByName(dataSheetName);
  
  // Find or create row in MoonlitData for the given date.
  let targetRow = findOrCreateDateRow(dataSheet, dateValue);
  
  // Build Column B string from B8:B20
  const keysBC = ['duty','sdsale','tstdstk','nssale','nsstk','sb','coins',
                  'oil','plasticsb','plasticloaf','loaf','#3','#6','tiny','medium', 'large'];
  const bVals = sheet.getRange(8, 2, keysBC.length, 1).getValues().map(r => r[0]);  // B8:B21
  const cVals = sheet.getRange(8, 3, keysBC.length, 1).getValues().map(r => r[0]);  // C8:C21
  let colBString = '';
  for (let i = 0; i < keysBC.length; i++) {
    const val = bVals[i] === undefined ? '' : bVals[i];
    colBString += `${keysBC[i]}=${val}; `;
  }
  colBString = colBString.trim();
  
  let colCString = '';
  for (let i = 0; i < keysBC.length; i++) {
    const val = cVals[i] === undefined ? '' : cVals[i];
    colCString += `${keysBC[i]}=${val}; `;
  }
  colCString = colCString.trim();
  
  // Build Column D string from J, K, L, M, N, O (preserving all blank lines)
  const expensesamString = gatherLinesKeepBlanks_(sheet, 10); // Column J
  const expensespmString = gatherLinesKeepBlanks_(sheet, 11); // Column K
  const inoutsString     = gatherLinesKeepBlanks_(sheet, 12); // Column L
  const billsString      = gatherLinesKeepBlanks_(sheet, 13); // Column M
  const pulloutsamString = gatherLinesKeepBlanks_(sheet, 14); // Column N
  const pulloutspmString = gatherLinesKeepBlanks_(sheet, 15); // Column O
  
  let colDString = '';
  colDString += `expensesam="${expensesamString}"; `;
  colDString += `expensespm="${expensespmString}"; `;
  colDString += `inouts="${inoutsString}"; `;
  colDString += `bills="${billsString}"; `;
  colDString += `pulloutsam="${pulloutsamString}"; `;
  colDString += `pulloutspm="${pulloutspmString}";`;
  
  // Build Column E string from F5 => "mineral=?; x15=?;"
  let f5Val = sheet.getRange('F5').getValue();
  let num = parseFloat(f5Val);
  let colEString = '';
  if (!isNaN(num)) {
    let x15 = num * 15;
    colEString = `mineral=${num}; x15=${x15};`;
  } else {
    colEString = `mineral=${f5Val}; x15=?;`;
  }
  
  // Build Column F string from F6 => duties="..."
  let f6Val = sheet.getRange('F6').getValue();
  if (f6Val === undefined) f6Val = '';
  let colFString = `duties="${f6Val}"`;
  
  // Build Column G string from F8:F20 (13 items), in order.
  const timeSalesKeys = ['dutyam','totalam','8pm','10pm','2am','6am',
                         'dutypm','totalpm','8am','10am','1pm','4pm','6pm'];
  const fVals = sheet.getRange('F8:F20').getValues().map(r => r[0]);
  let colGString = '';
  for (let i = 0; i < timeSalesKeys.length; i++) {
    const val = fVals[i] === undefined ? '' : fVals[i];
    colGString += `${timeSalesKeys[i]}=${val}; `;
  }
  colGString = colGString.trim();
  
  // Build Column H string from:
  // totaldeposit from P2:Q10, absences from F7,
  // accountsam from S2:S, accountspm from T2:T, coffeesales from F22:F28
  const totalDepositStr = buildTotalDepositString_(sheet);
  const absencesVal = sheet.getRange('F7').getValue() || '';
  const accountsamVal = gatherLinesKeepBlanks_(sheet, 19); // Column S
  const accountspmVal = gatherLinesKeepBlanks_(sheet, 20); // Column T
  const coffeeSalesVal = buildCoffeeSalesString_(sheet);
  
  let colHString = '';
  colHString += `totaldeposit="${totalDepositStr}"; `;
  colHString += `absences="${absencesVal}"; `;
  colHString += `accountsam="${accountsamVal}"; `;
  colHString += `accountspm="${accountspmVal}"; `;
  colHString += `coffeesales="${coffeeSalesVal}"`;
  
  // Write all data to MoonlitData row for this date.
  dataSheet.getRange(targetRow, 2).setValue(colBString);  // Column B
  dataSheet.getRange(targetRow, 3).setValue(colCString);  // Column C
  dataSheet.getRange(targetRow, 4).setValue(colDString);  // Column D
  dataSheet.getRange(targetRow, 5).setValue(colEString);  // Column E
  dataSheet.getRange(targetRow, 6).setValue(colFString);  // Column F
  dataSheet.getRange(targetRow, 7).setValue(colGString);  // Column G
  dataSheet.getRange(targetRow, 8).setValue(colHString);  // Column H
  // (No alert is shown; status is updated in B2 instead.)
}

/**
 * Gathers lines from a single column (rows 2 to 101) and keeps blank cells.
 * Trailing blank lines are removed to avoid excessive empties.
 * The lines are then joined using literal "\n" (i.e. "\\n").
 */
function gatherLinesKeepBlanks_(sheet, colIndex) {
  const range = sheet.getRange(2, colIndex, 100, 1);
  const values = range.getValues();
  const lines = [];
  for (let i = 0; i < values.length; i++) {
    let val = values[i][0] == null ? '' : values[i][0];
    lines.push(String(val));
  }
  while (lines.length > 0 && lines[lines.length - 1] === '') {
    lines.pop();
  }
  return lines.join('\\n');
}

/**
 * Builds the "totaldeposit" string from P2:P10 and Q2:Q10.
 * Each line is of the format "Label: Value" and joined with "\\n".
 */
function buildTotalDepositString_(sheet) {
  const rangeP = sheet.getRange('P2:P10');
  const rangeQ = sheet.getRange('Q2:Q10');
  const valuesP = rangeP.getValues();
  const valuesQ = rangeQ.getValues();
  let result = '';
  for (let i = 0; i < valuesP.length; i++) {
    const label = valuesP[i][0] || '';
    const val   = valuesQ[i][0] || '';
    result += `${label}: ${val}\\n`;
  }
  return result.trim();
}

/**
 * Builds the "coffeesales" string from cells F22:F28.
 * Each line is in the format "Key: Value" and joined with "\\n".
 */
function buildCoffeeSalesString_(sheet) {
  const keys = ['Sales','CupsEnd','Coffee','Choco','Caramel','Cappuccino','3in1'];
  const rowStart = 22;
  let lines = [];
  for (let i = 0; i < keys.length; i++) {
    const val = sheet.getRange(rowStart + i, 6).getValue() || '';
    lines.push(`${keys[i]}: ${val}`);
  }
  return lines.join('\\n');
}

/**
 * Deposit calculation used by deposit checkbox Q1.
 */
function calculateDeposit(sheet, valueCell) {
  let valueString = sheet.getRange(valueCell).getValue();
  valueString = String(valueString).replace(/\\n/g, '\n');
  const lines = valueString.trim().split(/\n/);
  let total = 0;
  lines.forEach(line => {
    const parts = line.split('=');
    if (parts.length === 2) {
      const value = parseFloat(parts[1].trim());
      if (!isNaN(value)) total += value;
    }
  });
  var income1range = sheet.getRange('Q2:Q5');
  var income1values = income1range.getValues();
  var income1 = 0;
  
  // Loop through the values and sum them
  for (var i = 0; i < income1values.length; i++) {
    income1 += parseFloat(income1values[i][0]) || 0;
  }
  const income2 = parseFloat(sheet.getRange('Q7').getValue()) || 0;
  const deposit = income1 + income2 - total;
  sheet.getRange('Q6').setValue(income1);
  sheet.getRange('Q9').setValue(deposit);
}

/**
 * Simple parser for "key=value;" pairs.
 */
function parseKeyValuePairs(str) {
  const result = {};
  if (!str || typeof str !== 'string') return result;
  const pairs = str.split(';');
  pairs.forEach(pair => {
    const trimmed = pair.trim();
    if (trimmed) {
      const eqIndex = trimmed.indexOf('=');
      if (eqIndex > 0) {
        const key = trimmed.substring(0, eqIndex).trim();
        const val = trimmed.substring(eqIndex + 1).trim();
        result[key] = val;
      }
    }
  });
  return result;
}

/**
 * Clears cells before retrieving or when the Clear checkbox is toggled.
 * (Note: Cells F9 and F15 are not cleared so that your SUM formulas remain.)
 */
function clearCellsForNewDate(spreadsheet) {
  const sheet = spreadsheet.getSheetByName('MoonlitReport');
  sheet.getRange('Q2:Q10').clearContent();
  sheet.getRange('S2:S').clearContent();
  sheet.getRange('T2:T').clearContent();
  sheet.getRange('B8:B23').clearContent();
  sheet.getRange('C8:C23').clearContent();
  sheet.getRange('F8').clearContent();
  sheet.getRange('F10:F14').clearContent();
  sheet.getRange('F16:F20').clearContent();
  sheet.getRange('F22:F28').clearContent();
  sheet.getRange('F5:F6').clearContent();
  sheet.getRange('J2:J').clearContent();
  sheet.getRange('K2:K').clearContent();
  sheet.getRange('L2:L').clearContent();
  sheet.getRange('M2:M').clearContent();
  sheet.getRange('N2:N').clearContent();
  sheet.getRange('O2:O').clearContent();
}

/**
 * Retrieval logic:
 * - Parses Column G into key-value pairs and writes them to F8:F20 (skipping F9 and F15).
 * - Parses Column D into multiple fields and writes them to columns J, K, L, M, N, O.
 * - Also parses Column H for totaldeposit, absences, accountsam, accountspm, coffeesales.
 * If the date is not found, a new row is automatically created.
 */
function retrieveDataFromMoonlitData(e) {
  const sheetName = 'MoonlitReport';
  const dataSheetName = 'MoonlitData';
  const dateCell = 'B1';
  
  const sheet = e.source.getSheetByName(sheetName);
  const dateValue = new Date(sheet.getRange(dateCell).getValue());
  const dataSheet = e.source.getSheetByName(dataSheetName);
  
  // Find or create the row in MoonlitData for the given date.
  let targetRow = findOrCreateDateRow(dataSheet, dateValue);
  
  // 1) Column B & C: populate B8:B20 and C8:C20
  const colBString = dataSheet.getRange(targetRow, 2).getValue();
  const colCString = dataSheet.getRange(targetRow, 3).getValue();
  const keysForBC = [
    'duty','sdsale','tstdstk','nssale','nsstk','sb','coins',
    'oil','plasticsb','plasticloaf','loaf','#3','#6','tiny','medium','large'
  ];
  const parsedB = parseKeyValuePairs(colBString);
  const parsedC = parseKeyValuePairs(colCString);
  const bValues = [];
  const cValues = [];
  for (let i = 0; i < keysForBC.length; i++) {
    const key = keysForBC[i];
    bValues.push([parsedB[key] || '']);
    cValues.push([parsedC[key] || '']);
  }
  sheet.getRange(8, 2, keysForBC.length, 1).setValues(bValues); // B8:B21
  sheet.getRange(8, 3, keysForBC.length, 1).setValues(cValues); // C8:C21
  
  // 2) Column E: "mineral=..." => F5
  const colEString = dataSheet.getRange(targetRow, 5).getValue();
  const parsedE = parseKeyValuePairs(colEString);
  const mineralVal = parsedE['mineral'] || '';
  sheet.getRange('F5').setValue(mineralVal);
  
  // 3) Column F: duties="..." => F6
  const colFString = dataSheet.getRange(targetRow, 6).getValue();
  const dutiesMatch = colFString.match(/duties="([^"]*)"/);
  if (dutiesMatch) {
    sheet.getRange('F6').setValue(dutiesMatch[1]);
  }
  
  // 4) Column G: timeSales => parse and write to F8:F20 (skip F9 and F15)
  const colGString = dataSheet.getRange(targetRow, 7).getValue();
  if (colGString && typeof colGString === 'string') {
    const pairs = colGString.split(';');
    const timeSalesMap = {};
    pairs.forEach(item => {
      const trimmed = item.trim();
      if (!trimmed) return;
      const eqIndex = trimmed.indexOf('=');
      if (eqIndex > 0) {
        const key = trimmed.substring(0, eqIndex).trim();
        const val = trimmed.substring(eqIndex + 1).trim();
        timeSalesMap[key] = val;
      }
    });
    const rowMap = {
      dutyam:  8,
      totalam: 9,   // skip writing
      '8pm':   10,
      '10pm':  11,
      '2am':   12,
      '6am':   13,
      dutypm:  14,
      totalpm: 15,  // skip writing
      '8am':   16,
      '10am':  17,
      '1pm':   18,
      '4pm':   19,
      '6pm':   20
    };
    for (const [key, row] of Object.entries(rowMap)) {
      if (key === 'totalam' || key === 'totalpm') continue;
      if (timeSalesMap[key] !== undefined) {
        sheet.getRange(row, 6).setValue(timeSalesMap[key]);
      }
    }
  }
  
  // 5) Column D: Parse multiline fields and write to columns J (expensesam), K (expensespm), L (inouts), M (bills), N (pulloutsam), O (pulloutspm)
  const colDString = dataSheet.getRange(targetRow, 4).getValue();
  function placeLinesInColumn(regexStr, columnIndex) {
    const match = colDString.match(new RegExp(regexStr));
    if (match) {
      const text = match[1];
      const lines = text.split('\\n');
      const arr = lines.map(line => [line]);
      if (arr.length > 0) {
        sheet.getRange(2, columnIndex, arr.length, 1).setValues(arr);
      }
    }
  }
  placeLinesInColumn('expensesam="([^"]*)"', 10);
  placeLinesInColumn('expensespm="([^"]*)"', 11);
  placeLinesInColumn('inouts="([^"]*)"',    12);
  placeLinesInColumn('bills="([^"]*)"',     13);
  placeLinesInColumn('pulloutsam="([^"]*)"',14);
  placeLinesInColumn('pulloutspm="([^"]*)"',15);
  
  // 6) Column H: Parse totaldeposit, absences, accountsam, accountspm, coffeesales
  let rightOtherCellValue = dataSheet.getRange(targetRow, 8).getValue();
  let rightOtherValue = '';
  const totalDepositMatch = rightOtherCellValue.match(/totaldeposit="([^"]*)"/);
  if (totalDepositMatch) {
    rightOtherValue = totalDepositMatch[1];
  }
  let linesArr = [];
  let expensesStarted = false;
  let expensesValue = '';
  rightOtherValue.split(/\\n/).forEach(line => {
    if (line.startsWith('Expenses')) {
      expensesStarted = true;
      expensesValue = line;
    } else if (expensesStarted && !line.startsWith('Grand Net Total')) {
      expensesValue += `\\n${line}`;
    } else {
      if (expensesStarted) {
        linesArr.push(expensesValue);
        expensesStarted = false;
      }
      linesArr.push(line);
    }
  });
  if (expensesStarted) {
    linesArr.push(expensesValue);
  }
  const pValues = [];
  const qValues = [];
  linesArr.forEach(line => {
    const colonIndex = line.indexOf(":");
    if (colonIndex !== -1) {
      const label = line.substring(0, colonIndex).trim();
      const value = line.substring(colonIndex + 1).trim();
      pValues.push([label]);
      qValues.push([value]);
    }
  });
  sheet.getRange('P2:P10').setValues(pValues);
  sheet.getRange('Q2:Q10').setValues(qValues);
  
  // absences => F7
  const absencesMatch = rightOtherCellValue.match(/absences="([^"]*)"/);
  if (absencesMatch) {
    sheet.getRange('F7').setValue(absencesMatch[1]);
  }
  
  // coffeesales => F22:F28
  const coffeesalesMatch = rightOtherCellValue.match(/coffeesales="([^"]*)"/);
  if (coffeesalesMatch) {
    const coffeeStr = coffeesalesMatch[1];
    const coffeeLines = coffeeStr.split('\\n');
    const coffeeObj = {};
    coffeeLines.forEach(line => {
      const idx = line.indexOf(':');
      if (idx !== -1) {
        const key = line.substring(0, idx).trim();
        const val = line.substring(idx + 1).trim();
        coffeeObj[key] = val;
      }
    });
    sheet.getRange('F22').setValue(coffeeObj['Sales'] || '');
    sheet.getRange('F23').setValue(coffeeObj['CupsEnd'] || '');
    sheet.getRange('F24').setValue(coffeeObj['Coffee'] || '');
    sheet.getRange('F25').setValue(coffeeObj['Choco'] || '');
    sheet.getRange('F26').setValue(coffeeObj['Caramel'] || '');
    sheet.getRange('F27').setValue(coffeeObj['Cappuccino'] || '');
    sheet.getRange('F28').setValue(coffeeObj['3in1'] || '');
  }
  
  // accountsam => S2:S, accountspm => T2:T
  const accountsamMatch = rightOtherCellValue.match(/accountsam="([^"]*)"/);
  if (accountsamMatch) {
    const accountsamString = accountsamMatch[1];
    const linesAm = accountsamString.split('\\n');
    const sValues = linesAm.map(line => [line]);
    sheet.getRange(2, 19, sValues.length, 1).setValues(sValues);
  }
  const accountspmMatch = rightOtherCellValue.match(/accountspm="([^"]*)"/);
  if (accountspmMatch) {
    const accountspmString = accountspmMatch[1];
    const linesPm = accountspmString.split('\\n');
    const tValues = linesPm.map(line => [line]);
    sheet.getRange(2, 20, tValues.length, 1).setValues(tValues);
  }
  // Retrieve status: gray text
}

/**
 * Automatically finds or creates a new row in MoonlitData for the given date.
 */
function findOrCreateDateRow(dataSheet, dateValue) {
  const dateColumn = dataSheet.getRange('A:A').getValues();
  for (let i = 0; i < dateColumn.length; i++) {
    const currentDate = new Date(dateColumn[i][0]);
    if (currentDate.getTime() === dateValue.getTime()) {
      //retrieved
      updateStatus("retrieved on " + getPhilippineTime(), "#C0C0C0");
      return i + 1;
    }
  }
  let newRow = dataSheet.getLastRow() + 1;
  dataSheet.getRange(newRow, 1).setValue(dateValue);
  const defaultDValue = 'expensesam=""; expensespm=""; inouts=""; bills=""; pulloutsam=""; pulloutspm="";';
  dataSheet.getRange(newRow, 4).setValue(defaultDValue);
  dataSheet.getRange(newRow, 6).setValue('duties=""');
  const defaultHValue =
    'totaldeposit="Sales AM: \\nSales PM: \\nCoffee Sales: \\nNS:\\n\\nTotal Sales:\\n\\nTotal Previous:\\nExpenses:\\nGrand Net Total:\\nDeposited (Y or N):"; absences=""; accountsam=""; accountspm=""; coffeesales="Sales: \\nCupsEnd: \\nCoffee: \\nChoco: \\nCaramel: \\nCappuccino: \\n3in1:"';
  dataSheet.getRange(newRow, 8).setValue(defaultHValue);
  // Update status: gray text
  updateStatus("created on " + getPhilippineTime(), "#C0C0C0");
  return newRow;
}

/**
 * Optional debugging
 */
function debugFunction() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const fakeEvent = { source: spreadsheet };
  retrieveDataFromMoonlitData(fakeEvent);
}

