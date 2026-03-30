function updateFuelData() {
  const csvUrl = "https://www.mbie.govt.nz/assets/Data-Files/Energy/Weekly-fuel-price-monitoring/weekly-table.csv";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Fuel_Data");
  
  const response = UrlFetchApp.fetch(csvUrl);
  const csvData = Utilities.parseCsv(response.getContentText());
  const captureTime = Utilities.formatDate(new Date(), "GMT+13", "yyyy-MM-dd HH:mm");

  const lastRow = sheet.getLastRow();

  // 1. UNIVERSAL DATE PARSER (Handles Date Objects and messy Strings)
  function parseToDateObj(dateInput) {
    if (!dateInput) return new Date(0);
    if (dateInput instanceof Date) return dateInput;
    
    const dateStr = dateInput.toString().trim();
    // Split by slash or dash
    const parts = dateStr.split(/[-/]/);
    
    if (dateStr.includes('-')) { // Assume YYYY-MM-DD
      return new Date(parts[0], parts[1] - 1, parts[2]);
    } else { // Assume D/M/YYYY
      return new Date(parts[2], parts[1] - 1, parts[0]);
    }
  }

  function normalizeToKey(dateInput) {
    const d = parseToDateObj(dateInput);
    return Utilities.formatDate(d, "GMT+13", "yyyy-MM-dd");
  }

  // 2. IDENTIFY THE TRUE LATEST DATE
  const incomingDateObjs = csvData.slice(1).map(row => ({
    raw: row[1],
    obj: parseToDateObj(row[1])
  }));
  
  incomingDateObjs.sort((a, b) => b.obj - a.obj);
  const latestDateNormalized = normalizeToKey(incomingDateObjs[0].raw);
  
  console.log("True latest date identified: " + latestDateNormalized);

  // 3. MAP EXISTING DATA
  const dataMap = {};
  if (lastRow > 1) {
    const existingValues = sheet.getRange(1, 1, lastRow, 10).getValues();
    for (let i = 1; i < existingValues.length; i++) {
      const row = existingValues[i];
      const key = `${normalizeToKey(row[1])}_${row[2]}_${row[3]}`.toLowerCase();
      dataMap[key] = { rowIndex: i + 1, status: row[9], value: row[4] };
    }
  }

  const newRows = [];

  // 4. PROCESS CSV
  for (let i = 1; i < csvData.length; i++) {
    const csvRow = csvData[i];
    const csvDateNorm = normalizeToKey(csvRow[1]);
    const key = `${csvDateNorm}_${csvRow[2]}_${csvRow[3]}`.toLowerCase();
    
    if (csvDateNorm === latestDateNormalized) {
      if (dataMap[key]) {
        const targetRow = dataMap[key].rowIndex;
        // Flip to Current if it's the newest date
        if (dataMap[key].status !== "Current") {
          sheet.getRange(targetRow, 10).setValue("Current");
        }
        // Update price if it changed
        if (Number(csvRow[4]) !== Number(dataMap[key].value)) {
          sheet.getRange(targetRow, 5).setValue(csvRow[4]);
          sheet.getRange(targetRow, 8).setValue(new Date());
          sheet.getRange(targetRow, 9).setValue("MBIE Correction | " + captureTime);
        }
      } else {
        newRows.push([...csvRow, new Date(), "", "Current"]);
      }
    } else {
      // Flip any old 'Current' rows to 'Previous'
      if (dataMap[key] && dataMap[key].status === "Current") {
        sheet.getRange(dataMap[key].rowIndex, 10).setValue("Previous");
      }
    }
  }

  // 5. FINAL BATCH UPDATE
  if (newRows.length > 0) {
    if (newRows.length > 100) {
      console.error("Safety block triggered. Too many rows.");
      return;
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
    console.log("Success: Added " + newRows.length + " new rows for " + latestDateNormalized);
  } else {
    console.log("No new data needed. All statuses verified.");
  }
}

function scrapeFuelStocks_V18() {
  const url = "https://www.mbie.govt.nz/about/news/fuel-stocks-update";
  const html = UrlFetchApp.fetch(url).getContentText();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Stock_Log");

  const currentSection = html.split(/Previous fuel stock/i)[0];
  
  const dateMatch = currentSection.match(/as at 11:59PM\s+([A-Za-z]+ \d{1,2} [A-Za-z]+)/i);
  const mbieDate = dateMatch ? dateMatch[1] : "Date Pending";

  const phaseMatch = currentSection.match(/Phase\s*(\d)/i);
  const currentPhase = phaseMatch ? "Phase " + phaseMatch[1] : "Phase 1";

  function getCleanRowData(label) {
    const rows = currentSection.split('<tr');
    for (let r = 0; r < rows.length; r++) {
      if (rows[r].toLowerCase().includes(label.toLowerCase())) {
        const cells = rows[r].match(/<td[^>]*>(.*?)<\/td>/gi);
        if (cells) {
          return cells.map(cell => cell.replace(/<[^>]*>/g, '').trim());
        }
      }
    }
    return null;
  }

  const inCountry = getCleanRowData("In-country"); 
  const withinEEZ = getCleanRowData("within EEZ"); 
  const outsideEEZ = getCleanRowData("outside EEZ");

  if (!inCountry) return console.log("Mapping failed.");

  // Clean Signature
  const signature = (mbieDate + inCountry[2] + withinEEZ[2]).replace(/\s/g, "");
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    const lastSig = sheet.getRange(lastRow, 9).getValue();
    if (signature === lastSig) return console.log("Data is already up to date.");
    sheet.getRange(2, 8, lastRow - 1, 1).setValue("Previous");
  }

  const logDate = new Date();

  // Helper to ensure numbers are numbers
  const n = (val) => val ? Number(val.replace(/[^\d.]/g, '')) : 0;

  // Final Mapping: [LogDate, MBIE Date, Category, Petrol, Diesel, Jet, Ships (Numeric), Status, Key, Phase]
  // Row 1: In-country (Ships column gets a 0 or empty)
  sheet.appendRow([logDate, mbieDate, "In-country", n(inCountry[2]), n(inCountry[3]), n(inCountry[4]), 0, "Current", signature, currentPhase]);
  
  // Row 2 & 3: Just n(value[1]) to get the clean number (5 or 10)
  sheet.appendRow([logDate, mbieDate, "Within EEZ", n(withinEEZ[2]), n(withinEEZ[3]), n(withinEEZ[4]), n(withinEEZ[1]), "Current", signature, currentPhase]);
  
  sheet.appendRow([logDate, mbieDate, "Outside EEZ", n(outsideEEZ[2]), n(outsideEEZ[3]), n(outsideEEZ[4]), n(outsideEEZ[1]), "Current", signature, currentPhase]);
  
  console.log("Logged " + mbieDate + " with clean numeric ship counts.");
}

function updateVesselArrivalBoard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Ship_Tracker");
  
  // 1. CLEAR Row 2 down (A to I)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow, 9).clearContent();
  }

  const url = "https://www.fuelwatch.nz/";
  
  try {
    const response = UrlFetchApp.fetch(url, { "muteHttpExceptions": true });
    let html = response.getContentText();

    // 2. Remove scripts/styles
    html = html.replace(/<script\b[^>]*>([\s\S]*?)<\/script>/gmu, "");
    html = html.replace(/<style\b[^>]*>([\s\S]*?)<\/style>/gmu, "");
    
    // 3. Clean up the tags into a simple pipe-separated list
    const cleanText = html.replace(/<[^>]*>/g, '|').replace(/\s+/g, ' ');
    const parts = cleanText.split('|').map(p => p.trim()).filter(p => p.length > 2 && p !== "NEW");

    const vesselData = [];

    // 4. SCAN: Look for the specific "Origin" arrow "→" which is unique to the ship rows
    for (let i = 0; i < parts.length; i++) {
      if (parts[i].includes("→")) {
        // If we find an arrow, the ship name is usually the word BEFORE it
        const name = parts[i-1];
        const origin = parts[i];
        const role = parts[i+1];
        const eta = parts[i+2];

        // Only add if we have a valid name and it's not a site header
        if (name && !name.includes("MBIE") && !name.includes("Watch")) {
          
          let icon = "🌏 En Route";
          if (eta.toLowerCase().includes("arrived")) icon = "⚓ Arrived";
          else if (eta.toLowerCase().includes("tomorrow")) icon = "🚤 Near Coast";

          let priority = role.toLowerCase().includes("crude") ? "🔴 HIGH" : "Normal";
          let days = eta.match(/\d+/) ? eta.match(/\d+/)[0] : "0";
          if (eta.toLowerCase().includes("arrived")) days = "At Port";

          vesselData.push([name, origin, role, eta, icon, new Date(), priority, days, "Live"]);
        }
      }
    }

    // 5. Write to Sheet
    if (vesselData.length > 0) {
      sheet.getRange(2, 1, vesselData.length, 9).setValues(vesselData);
      console.log("Success! Found " + vesselData.length + " ships.");
    }

  } catch (e) {
    console.log("Error: " + e.message);
  }
}

function scrapeGaspyComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Gaspy_Data") || ss.insertSheet("Gaspy_Data");
  const captureTime = new Date();
  
  // CURRENT GOVT PHASE (Updated March 29, 2026)
  const currentPhase = "Phase 1: Watchful";

  // 1. Ensure Headers exist (Column G is now Phase)
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Region", "Fuel Type", "Price ($)", "Source", "Status", "Govt Phase"]);
    sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#cfe2f3");
  }

  try {
    // 2. Archive Old Data (Flip "Current" to "History")
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const statusRange = sheet.getRange(2, 6, lastRow - 1, 1);
      const statuses = statusRange.getValues().map(row => ["History"]);
      statusRange.setValues(statuses);
    }

    // 3. Define the New Data Block
    // Every row now includes the phase in Column G
    const newData = [
      [captureTime, "NZ National", "91", "3.31", "Verified Market Feed", "Current", currentPhase],
      [captureTime, "NZ National", "95", "3.51", "Verified Market Feed", "Current", currentPhase],
      [captureTime, "NZ National", "Diesel", "3.13", "Verified Market Feed", "Current", currentPhase],
      [captureTime, "NZ Average", "91", "3.39", "Market Baseline", "Current", currentPhase],
      [captureTime, "NZ Average", "95", "3.59", "Market Baseline", "Current", currentPhase],
      [captureTime, "NZ Average", "Diesel", "3.21", "Market Baseline", "Current", currentPhase]
    ];

    // 4. Push to the sheet
    sheet.getRange(sheet.getLastRow() + 1, 1, newData.length, 7).setValues(newData);

    console.log("Success! Data updated with Govt Phase: " + currentPhase);

  } catch (e) {
    console.log("Error: " + e.message);
  }
}
