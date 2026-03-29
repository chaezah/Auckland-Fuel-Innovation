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

function scrapeFuelStocks_V3() {
  const url = "https://www.mbie.govt.nz/about/news/fuel-stocks-update";
  const html = UrlFetchApp.fetch(url).getContentText();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Stock_Log");

  const dateMatch = html.match(/As at 11:59PM on (.*?),/);
  const mbieDate = dateMatch ? dateMatch[1] : "Date Not Found";

  // --- NEW: PHASE CAPTURE LOGIC ---
  // This looks for "Phase 1", "Phase 2", etc. on the page.
  const phaseMatch = html.match(/Phase\s*(\d)/i);
  const currentPhase = phaseMatch ? "Phase " + phaseMatch[1] : "Phase 1";
  // --------------------------------

  const tables = html.split('<table');
  const inCountryData = tables[1].split('<tr')[2].match(/[\d.]+/g); 
  const onWaterData = tables[1].split('<tr')[3].match(/[\d.]+/g);
  
  // Update signature to include the Phase (so if Phase changes, it triggers!)
  const currentDataSignature = mbieDate + inCountryData.join("") + onWaterData.join("") + currentPhase;

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const lastSignature = sheet.getRange(lastRow, 9).getValue(); 
    if (currentDataSignature === lastSignature) {
      console.log("No change in data, date, or phase. Skipping.");
      return; 
    }
  }

  const logDate = new Date();
  if (lastRow > 1) {
    sheet.getRange(2, 8, lastRow - 1, 1).setValue("Previous");
  }

  // Column J (10th column) will now store the Phase
  sheet.appendRow([logDate, mbieDate, "In-country", inCountryData[0], inCountryData[1], inCountryData[2], "", "Current", currentDataSignature, currentPhase]);
  sheet.appendRow([logDate, mbieDate, "On-water", onWaterData[0], onWaterData[1], onWaterData[2], "Ships Here", "Current", currentDataSignature, currentPhase]);
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

