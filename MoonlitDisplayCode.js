/***********************************************************
 *  MoonlitDisplayCode.gs
 ***********************************************************/

/**
 * Called when the checkbox (cell D2) on MoonlitDisplay is toggled.
 * It reads the date range from A2:B2, searches MoonlitData for rows within that range,
 * extracts the "accountsam" and "accountspm" data from Column H,
 * and writes the results to MoonlitDisplay starting at cell D3 (header) and D4 downward.
 * A final cell shows the total accounts per employee.
 */
function updateMoonlitAccounts(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var displaySheet = ss.getSheetByName("MoonlitDisplay");
  var dataSheet = ss.getSheetByName("MoonlitData");
  var timezone = ss.getSpreadsheetTimeZone();
  
  // Read start and end dates from MoonlitDisplay (cells A2 and B2)
  var startDateCell = displaySheet.getRange("A2").getValue();
  var endDateCell = displaySheet.getRange("B2").getValue();
  
  if (!(startDateCell instanceof Date) || !(endDateCell instanceof Date)) {
    SpreadsheetApp.getUi().alert("Please enter valid dates in cells A2 and B2.");
    return;
  }
  
  // Format the dates for header (MM/dd/yyyy)
  var startDateStr = Utilities.formatDate(startDateCell, timezone, "MM/dd/yyyy");
  var endDateStr = Utilities.formatDate(endDateCell, timezone, "MM/dd/yyyy");
  
  // Set header in cell D3
  displaySheet.getRange("D3").setValue("Moonlit Bakery Accounts from " + startDateStr + " to " + endDateStr + ":");
  
  // Clear previous results from D4 downward (adjust the range as needed)
  displaySheet.getRange("D4:D100").clearContent();
  
  // Read data from MoonlitData.
  // Column A contains dates and Column H (index 8) contains the composite key–value string.
  var lastRow = dataSheet.getLastRow();
  if (lastRow < 1) return;
  var dataValues = dataSheet.getRange(1, 1, lastRow, 8).getValues(); // columns A to H
  
  var output = []; // To store daily account outputs.
  var totals = {}; // To accumulate all individual account amounts.
  
  // Process each row in MoonlitData.
  for (var i = 0; i < dataValues.length; i++) {
    var row = dataValues[i];
    var dateVal = row[0]; // Column A
    if (!(dateVal instanceof Date)) continue;
    if (dateVal < startDateCell || dateVal > endDateCell) continue;
    
    // Format the date (using M/d/yyyy for accounts display)
    var dateStr = Utilities.formatDate(dateVal, timezone, "M/d/yyyy");
    
    // Column H (index 7) contains the composite string.
    var text = row[7];
    if (typeof text !== "string") text = "";
    
    // Extract accounts for AM and PM using our helper function.
    var accountsam = extractKeyFromCell(text, "accountsam");
    var accountspm = extractKeyFromCell(text, "accountspm");
    
    // Replace any escaped newline characters with actual newlines.
    if (accountsam) {
      accountsam = accountsam.replace(/\\n/g, "\n").trim();
    }
    if (accountspm) {
      accountspm = accountspm.replace(/\\n/g, "\n").trim();
    }
    
    // Split the account strings into an array of lines.
    var accountsamLines = accountsam ? accountsam.split("\n").filter(function(l){ return l.trim() != ""; }) : [];
    var accountspmLines = accountspm ? accountspm.split("\n").filter(function(l){ return l.trim() != ""; }) : [];
    
    // Update totals by processing each line in the form "Name=number"
    function processLine(line) {
      var parts = line.split("=");
      if (parts.length === 2) {
        var name = parts[0].trim();
        var value = parseFloat(parts[1].trim());
        if (!isNaN(value)) {
          if (!totals[name]) totals[name] = [];
          totals[name].push(value);
        }
      }
    }
    accountsamLines.forEach(processLine);
    accountspmLines.forEach(processLine);
    
    // Build the daily output text.
    var dayOutput = "Date: " + dateStr + ":\n";
    dayOutput += "  Account AM:\n";
    if (accountsamLines.length > 0) {
      accountsamLines.forEach(function(line) {
        dayOutput += "    " + line + "\n";
      });
    } else {
      dayOutput += "    None\n";
    }
    dayOutput += "  Account PM:\n";
    if (accountspmLines.length > 0) {
      accountspmLines.forEach(function(line) {
        dayOutput += "    " + line + "\n";
      });
    } else {
      dayOutput += "    None\n";
    }
    dayOutput = `${dayOutput}\n\n.`;
    //dayOutput = dayOutput.trim(); // Remove any trailing newline.
    output.push([dayOutput]);
  }
  
  // Write the daily outputs starting at cell D4.
  if (output.length > 0) {
    displaySheet.getRange(4, 4, output.length, 1).setValues(output);
  }
  
  // Now, build the totals summary string.
  // Format:
  // Total Accounts :
  //
  // Jeanny: 10+5+25+10+16+16+25+5 = 112
  // Nina: 80+100+10+10 = 200
  // Romelyn: 25+50 = 75
  // Anna: 10+5+25+5+15+6 = 66
  var totalsOutput = "Total Accounts :\n\n";
  for (var key in totals) {
    if (totals.hasOwnProperty(key)) {
      var amounts = totals[key];
      var sum = amounts.reduce(function(a, b) { return a + b; }, 0);
      var amountsStr = amounts.join("+");
      totalsOutput += key + ": " + amountsStr + " = " + sum + "\n";
    }
  }
  totalsOutput = totalsOutput.trim();
  
  // Write the totals summary in the next available row in column D.
  var totalsRow = 4 + output.length;
  displaySheet.getRange(totalsRow, 4).setValue(totalsOutput);
}

function updateBillsDisplay(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var displaySheet = ss.getSheetByName("MoonlitDisplay");
  var dataSheet    = ss.getSheetByName("MoonlitData");
  var timezone     = ss.getSpreadsheetTimeZone();
  
  // Read date range from A2 and B2
  var startDateCell = displaySheet.getRange("A2").getValue();
  var endDateCell   = displaySheet.getRange("B2").getValue();
  if (!(startDateCell instanceof Date) || !(endDateCell instanceof Date)) {
    SpreadsheetApp.getUi().alert("Please enter valid dates in cells A2 and B2.");
    return;
  }
  
  // Header in C3
  var startDateStr = Utilities.formatDate(startDateCell, timezone, "MM/dd/yyyy");
  var endDateStr   = Utilities.formatDate(endDateCell,   timezone, "MM/dd/yyyy");
  displaySheet.getRange("C3").setValue(
    "Check Bills from " + startDateStr + " to " + endDateStr
  );
  
  // Clear old results
  displaySheet.getRange("C4:C100").clearContent();
  
  // Fetch all rows from MoonlitData (cols A–D)
  var lastRow    = dataSheet.getLastRow();
  if (lastRow < 1) return;
  var dataValues = dataSheet.getRange(1, 1, lastRow, 4).getValues();
  
  var output = [];
  
  dataValues.forEach(function(row) {
    var dateVal = row[0];
    if (!(dateVal instanceof Date)) return;
    if (dateVal < startDateCell || dateVal > endDateCell) return;
    
    var dateStr = Utilities.formatDate(dateVal, timezone, "MM/dd/yyyy");
    var raw     = row[3] || "";
    var bills   = extractBills(raw).replace(/\\n/g, "\n").trim();
    
    // **ONLY** push if there's something in bills
    if (bills !== "") {
      var line = "Date " + dateStr + ":\n" + bills + "\n\n";
      output.push([line]);
    }
  });
  
  // Write filtered results starting at C4
  if (output.length) {
    displaySheet.getRange(4, 3, output.length, 1).setValues(output);
  }
}

/**
 * Extracts the bills value from a composite key="…" string.
 */
function extractBills(str) {
  var m = (str || "").match(/bills="([^"]*)"/);
  return (m && m[1]) || "";
}



/**
 * Helper function to extract the value for a given key from a composite string.
 * It expects a pattern like key="value" in the provided text.
 *
 * @param {string} text  The composite key–value string from a cell.
 * @param {string} key   The key to extract (e.g., "accountsam" or "accountspm").
 * @return {string}      The extracted value or an empty string if not found.
 */
function extractKeyFromCell(text, key) {
  var regex = new RegExp(key + '="([^"]*)"');
  var match = text.match(regex);
  if (match && match.length > 1) {
    return match[1];
  }
  return "";
}

/**
 * Helper function to extract the value of "bills" from a key-value string.
 * It expects the pattern bills="value" somewhere in the string.
 *
 * @param {string} str  The key–value string from MoonlitData's column D.
 * @return {string}     The extracted bills value or an empty string if not found.
 */
function extractBills(str) {
  if (typeof str !== "string") return "";
  var regex = /bills="([^"]*)"/;
  var match = str.match(regex);
  if (match && match.length > 1) {
    return match[1];
  }
  return "";
}


/**
 * Called when the "Display Salary" checkbox (cell E2) on MoonlitDisplay is checked.
 * It reads A2:B2 (date range), scans MoonlitData (cols A and F),
 * calculates each employee's total attendance (adjusting for HD, UT, OT),
 * then looks up that employee in MoonlitEmployeeInfo to pull rates/extras/deductions,
 * computes net salary, and writes a breakdown in column E starting at E4.
 */
function updateSalaryDisplay(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const displaySheet = ss.getSheetByName("MoonlitDisplay");
  const dataSheet    = ss.getSheetByName("MoonlitData");
  const empSheet     = ss.getSheetByName("MoonlitEmployeeInfo");
  const timezone     = ss.getSpreadsheetTimeZone();

  // 1) Read the date range from A2:B2
  const startDateCell = displaySheet.getRange("A2").getValue();
  const endDateCell   = displaySheet.getRange("B2").getValue();
  if (!(startDateCell instanceof Date) || !(endDateCell instanceof Date)) {
    SpreadsheetApp.getUi().alert("Please enter valid dates in cells A2 and B2.");
    return;
  }

  // 2) Format those dates for the header
  const startDateStr = Utilities.formatDate(startDateCell, timezone, "MM/dd/yyyy");
  const endDateStr   = Utilities.formatDate(endDateCell,   timezone, "MM/dd/yyyy");

  // 3) Write the header into E3, then clear any old results E4:E
  displaySheet.getRange("E3").setValue(
    "Attendance from " + startDateStr + " to " + endDateStr
  );
  displaySheet.getRange("E4:E100").clearContent();

  // 4) Build a map: employeeName -> attendanceAmount (a floating‐point sum),
  //    by scanning MoonlitData rows 2→lastRow for any date within [startDate, endDate].
  const lastRow = dataSheet.getLastRow();
  if (lastRow < 2) {
    // No data rows at all
    return;
  }
  // We only need columns A (Date) and F (duties="…")
  // dataSheet.getRange(row, 1, numRows, 6) gives us [Date, …, (col 6) duties]
  const dataValues = dataSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  // dataValues[i][0] = row’s Date; dataValues[i][5] = row’s “duties=…” string
  const attendanceCounts = {};

  for (let i = 0; i < dataValues.length; i++) {
    const row = dataValues[i];
    const dateVal = row[0];
    if (!(dateVal instanceof Date)) {
      continue;
    }
    if (dateVal < startDateCell || dateVal > endDateCell) {
      continue;
    }

    // Extract the inner duties string from something like:  duties="Romelyn, Nina(OT:3), Anna(UT:2)"
    const rawDuties = row[5] || "";
    const m = rawDuties.match(/duties="([^"]*)"/);
    const innerDuties = m && m[1] ? m[1].trim() : "";
    if (!innerDuties) continue;

    // Split on commas → ["Romelyn", "Nina(OT:3)", "Anna(UT:2)", …]
    const parts = innerDuties.split(",").map(s => s.trim()).filter(s => s.length);

    parts.forEach(entry => {
      let name = "";
      let incAttendance = 0;

      //  a) Half‐day?
      const hdMatch = entry.match(/^(.+)\(HD\)$/i);
      if (hdMatch) {
        name = hdMatch[1].trim();
        incAttendance = 0.5;
      }
      else {
        //  b) Overtime (OT:<hours>)
        const otMatch = entry.match(/^(.+)\(OT:(\d+(\.\d+)?)\)$/i);
        if (otMatch) {
          name = otMatch[1].trim();
          const hours = parseFloat(otMatch[2]);
          // Full day = 12 hours; OT adds on top. So total hours = 12 + hours.
          incAttendance = (12 + hours) / 12;
        }
        else {
          //  c) Undertime (UT:<hours>)
          const utMatch = entry.match(/^(.+)\(UT:(\d+(\.\d+)?)\)$/i);
          if (utMatch) {
            name = utMatch[1].trim();
            const hours = parseFloat(utMatch[2]);
            // Only worked “hours” out of 12, so attendance = hours/12
            incAttendance = hours / 12;
          }
          else {
            //  d) Plain full‐day name
            name = entry.trim();
            incAttendance = name ? 1 : 0;
          }
        }
      }

      if (name) {
        if (!attendanceCounts[name]) {
          attendanceCounts[name] = 0;
        }
        attendanceCounts[name] += incAttendance;
      }
    });
  }

  // 5) Load all of MoonlitEmployeeInfo (rows 2→end)
  const empLastRow = empSheet.getLastRow();
  if (empLastRow < 2) {
    // No employees at all → nothing to show
    return;
  }
  // We expect columns:
  //   A = Name
  //   B = Daily Rate 1
  //   C = Daily Rate 2
  //   D = Extra per day
  //   E = Solo for baker (days)
  //   F = Bakery Allowance
  //   G = Cash Advances
  //   H = Accounts
  //   I = Charges
  //   J = SSS
  //   K = Philhealth
  //   L = Previous Balance
  const empValues = empSheet.getRange(2, 1, empLastRow - 1, 12).getValues();
  // Build a quick lookup: name → { daily1, daily2, extraPerDay, soloDays, bakeryAllowance, ca, accounts, charges, sss, philhealth, previous }
  const empInfo = {};
  empValues.forEach(row => {
    const name = String(row[0] || "").trim();
    if (!name) return;
    empInfo[name] = {
      daily1:       parseFloat(row[1]) || 0,
      daily2:       parseFloat(row[2]) || 0,
      extraPerDay:  parseFloat(row[3]) || 0,
      soloDays:     parseFloat(row[4]) || 0,
      bakeryAllowance: parseFloat(row[5]) || 0,
      cashAdvances: parseFloat(row[6]) || 0,
      accounts:     parseFloat(row[7]) || 0,
      charges:      parseFloat(row[8]) || 0,
      sss:          parseFloat(row[9]) || 0,
      philhealth:   parseFloat(row[10]) || 0,
      previousBalance: parseFloat(row[11]) || 0
    };
  });

  // 6) Build each employee’s salary‐breakdown line
  const output = [];
  Object.keys(attendanceCounts).forEach(empName => {
    const att = attendanceCounts[empName];
    // Round attendance to, say, 6 decimals (optional). You can leave it un‐rounded if you want full precision.
    const attendance = parseFloat(att.toFixed(6));

    const info = empInfo[empName];
    // If no info found, “leave blank” (skip)
    if (!info) {
      return;
    }

    // Pull all relevant fields:
    const d1 = info.daily1;
    const d2 = info.daily2;
    const extra = info.extraPerDay;
    const solo = info.soloDays;
    const bakeryAllow = info.bakeryAllowance;
    const ca   = info.cashAdvances;
    const acct = info.accounts;
    const chg  = info.charges;
    const sss  = info.sss;
    const phi  = info.philhealth;
    const prevBal = info.previousBalance;

    //  a) Compute base daily pay
    const totalD1 = attendance * d1;
    const totalD2 = attendance * d2;
    const sumDaily = totalD1 + totalD2;

    //  b) Compute additional (extra per day)
    const totalAdditional = extra * attendance;

    //  c) Compute solo‐baker pay = soloDays * 135 (if >0)
    const totalSolo = solo > 0 ? solo * 135 : 0;

    //  d) Bakery Allowance (once, not multiplied by attendance)
    const totalBakeryAllow = bakeryAllow || 0;

    //  e) Collect any deductions > 0
    const deductionsArr = [];
    if (ca > 0)   deductionsArr.push({ label: "Cash Advances", value: ca });
    if (acct > 0) deductionsArr.push({ label: "Accounts", value: acct });
    if (chg > 0)  deductionsArr.push({ label: "Charges", value: chg });
    if (sss > 0)  deductionsArr.push({ label: "SSS", value: sss });
    if (phi > 0)  deductionsArr.push({ label: "Philhealth", value: phi });

    //  f) Net salary = (sumDaily + totalAdditional + totalSolo + totalBakeryAllow) - sum(deductions)
    let deductionSum = 0;
    deductionsArr.forEach(d => deductionSum += d.value);
    const gross = sumDaily + totalAdditional + totalSolo + totalBakeryAllow;
    const netSalary = gross - deductionSum;

    // 7) Build the output string exactly in the requested format
    let line = empName + ": " +
               attendance + " days: " +
               attendance + " x (daily rate 1) " + d1 +
               " + " + attendance + " x (daily rate 2) " + d2 +
               "= " + totalD1 + " + " + totalD2 + " = " + sumDaily;

    if (totalAdditional > 0) {
      line += " + (Additional) " + extra + " x " + attendance +
              "(=" + totalAdditional + ")";
    }
    if (totalSolo > 0) {
      line += " + (Solo Days) 135 x " + solo + "(=" + totalSolo + ")";
    }
    if (totalBakeryAllow > 0) {
      line += " + Bakery Allowance " + totalBakeryAllow;
    }
    // Append any deductions
    deductionsArr.forEach(d => {
      line += " - " + d.label + " " + d.value;
    });
    line += " = Net Salary " + netSalary;

    // 8) If there is a previous balance > 0, also show: “Previous Balance X - Cash Advances X = Current Balance Y”
    if (prevBal > 0) {
      // Current Balance only subtracts “Cash Advances” from previous balance, as requested
      const currentBalance = prevBal - (ca || 0);
      line += "\nPrevious Balance " + prevBal +
              " - Cash Advances " + (ca || 0) +
              " = Current Balance " + currentBalance;
    }

    output.push([line]);
  });

  // 9) Finally, write all salary‐breakdown lines into E4, E5, E6, …  
  if (output.length) {
    displaySheet.getRange(4, 5, output.length, 1).setValues(output);
  }
}


