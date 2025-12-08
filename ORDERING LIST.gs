function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Refresh') 
    .addItem('Sync All Lists', 'runMasterSync') 
    .addToUi();
}

function runMasterSync() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // --- OLD LOGIC COMMENTED OUT FOR ISOLATED TESTING ---
    // 1. Sync PC & Config (Existing Logic)
    // updateOrderingList_PC_Config();
    
    // 2. Sync Layout Configuration (Insert Rows for Module/Vision)
    // updateOrderingList_Layout();

    // 3. Update Descriptions (Lookup from BOM for Module/Vision)
    // updateDescriptionsFromBOM();
    // ----------------------------------------------------

    // 4. NEW: Sync CORE Section from External Master BOM
    updateOrderingList_Core();

    // Notify User
    ui.alert("Sync Complete", "Core section has been updated successfully from the external Master BOM.", ui.ButtonSet.OK);

  } catch (e) {
    console.error(e);
    ui.alert("Error during Sync", e.message, ui.ButtonSet.OK);
  }
}

// =========================================
// SYNC 4: CORE SECTION (EXTERNAL SOURCE)
// =========================================
function updateOrderingList_Core() {
  // 1. SETUP SOURCE & DESTINATION
  var sourceSpreadsheetId = "1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA";
  var sourceTabName = "BOM Structure Tree Diagram";
  var destSheetName = "ORDERING LIST";
  var destSectionHeader = "CORE";

  // Open External Source
  try {
    var sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
  } catch (e) {
    throw new Error("Could not open Source Spreadsheet with ID: " + sourceSpreadsheetId + ". Check permissions.");
  }

  var sourceSheet = sourceSS.getSheetByName(sourceTabName);
  if (!sourceSheet) throw new Error("Source tab '" + sourceTabName + "' not found in external spreadsheet.");

  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var destSheet = destSS.getSheetByName(destSheetName);
  if (!destSheet) throw new Error("Destination sheet '" + destSheetName + "' not found.");

  // 2. READ SOURCE DATA
  // We look in Column C (Index 3) for the trigger "CORE :430000-A557"
  // We read cols C and D (Part ID and Description)
  
  var lastRow = sourceSheet.getLastRow();
  // Get all data from Col C and D (Rows 1 to LastRow)
  // getRange(row, col, numRows, numCols) -> Col C is 3. We want 2 columns (C & D).
  var sourceData = sourceSheet.getRange(1, 3, lastRow, 2).getValues();
  
  var startRowIndex = -1;
  var triggerPhrase = "CORE :430000-A557";

  // Find the anchor row
  for (var i = 0; i < sourceData.length; i++) {
    var cellValue = sourceData[i][0].toString().trim(); // Col C
    // We check if the cell *contains* the trigger, just in case of whitespace
    if (cellValue.indexOf(triggerPhrase) > -1) {
      startRowIndex = i + 1; // Data starts on the NEXT row
      break;
    }
  }

  if (startRowIndex === -1) {
    throw new Error("Could not find anchor '" + triggerPhrase + "' in Column C of the source file.");
  }

  // Extract Valid Items
  var coreItems = [];
  
  // Loop from below the anchor to the end of the sheet
  for (var k = startRowIndex; k < sourceData.length; k++) {
    var pID = sourceData[k][0].toString().trim(); // Col C
    var desc = sourceData[k][1].toString().trim(); // Col D
    
    // Logic: If Part ID is not empty, it's a valid item. 
    // We ignore empty rows or separators line "---" if they exist.
    if (pID !== "" && pID !== "---") {
      // Structure for sync: [Part ID, Description, Qty]
      // We hardcode Qty to "1" as it is not present in Source Col C/D
      coreItems.push([pID, desc, "1"]);
    }
  }

  console.log("Found " + coreItems.length + " Core items.");

  // 3. EXECUTE SYNC
  var sectionsToSync = [
    { destName: destSectionHeader, items: coreItems }
  ];

  performSurgicalSync(destSheet, sectionsToSync);
}


// =========================================
// SYNC 1: PC & CONFIG (EXISTING LOGIC)
// =========================================
function updateOrderingList_PC_Config() {
  var sourceSheetName = "OTHER MODULES/KITS MATRIX TO RELEASE";
  var destSheetName = "ORDERING LIST";
  
  var sectionMapping = {
    "PC": "PC",
    "MC Config": "CONFIG"
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(sourceSheetName);
  var destSheet = ss.getSheetByName(destSheetName);

  if (!sourceSheet || !destSheet) {
    console.warn("PC/Config Source or Dest sheet not found. Skipping Step 1.");
    return;
  }

  var lastRowSource = sourceSheet.getLastRow();
  if (lastRowSource < 1) return;
  
  var sourceData = sourceSheet.getRange(1, 1, lastRowSource, 7).getValues();

  var sections = [];
  var currentSectionName = "";
  var currentSectionRows = [];

  for (var i = 0; i < sourceData.length; i++) {
    var row = sourceData[i];
    var colA = row[0].toString().trim();

    if (colA !== "" && colA !== currentSectionName) {
      if (currentSectionRows.length > 0) {
        sections.push({name: currentSectionName, rows: currentSectionRows});
      }
      currentSectionName = colA;
      currentSectionRows = [];
    }
    currentSectionRows.push(row);
  }
  if (currentSectionRows.length > 0) {
    sections.push({name: currentSectionName, rows: currentSectionRows});
  }

  var processedSections = [];
  for (var j = 0; j < sections.length; j++) {
    var section = sections[j];
    if (!sectionMapping.hasOwnProperty(section.name)) continue;
    
    var validItems = []; 
    var dataStarted = false;

    for (var k = 0; k < section.rows.length; k++) {
      var r = section.rows[k];
      var valE = r[4].toString().trim().toUpperCase(); 
      var valF = r[5].toString().trim().toUpperCase(); 

      if (!dataStarted) {
        if (valE === "PART ID" || valF === "DESCRIPTION") dataStarted = true;
      } else {
        var pID = r[4].toString();
        var desc = r[5].toString();
        var qty = r[6].toString(); 
        
        if (pID !== "PART ID" && desc !== "DESCRIPTION" && (pID !== "" || desc !== "")) {
           validItems.push([pID, desc, qty]);
        }
      }
    }
    
    if (!dataStarted && validItems.length === 0 && section.rows.length > 0) {
      for (var k = 0; k < section.rows.length; k++) {
         var pID = section.rows[k][4].toString();
         var desc = section.rows[k][5].toString();
         var qty = section.rows[k][6].toString(); 
         if (pID !== "PART ID" && desc !== "DESCRIPTION" && (pID !== "" || desc !== "")) {
           validItems.push([pID, desc, qty]);
         }
      }
    }
    
    processedSections.push({destName: sectionMapping[section.name], items: validItems});
  }

  performSurgicalSync(destSheet, processedSections);
}

// =========================================
// SYNC 2: LAYOUT CONFIGURATION (INSERT IDs)
// =========================================
function updateOrderingList_Layout() {
  var layoutSheetName = "LAYOUT CONFIGURATION";
  var destSheetName = "ORDERING LIST";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(layoutSheetName);
  var destSheet = ss.getSheetByName(destSheetName);

  if (!sourceSheet) {
    console.warn("Layout Configuration sheet not found. Skipping Step 2.");
    return;
  }

  var lastRow = sourceSheet.getLastRow();
  var colA_Values = sourceSheet.getRange(1, 1, lastRow, 1).getValues(); 
  
  var startRow = -1;
  for (var i = 0; i < colA_Values.length; i++) {
    if (colA_Values[i][0].toString().trim().toUpperCase() === "CONFIGURATION") {
      startRow = i + 1;
      break;
    }
  }

  if (startRow === -1) return;

  var rowsToProcess = lastRow - startRow + 1;
  if (rowsToProcess < 1) return;

  // Grab data starting from the "CONFIGURATION" row (Col L=12, Col M=13)
  var dataRange = sourceSheet.getRange(startRow, 12, rowsToProcess, 2).getValues();

  var moduleItems = [];
  var visionItems = [];

  for (var r = 0; r < dataRange.length; r++) {
    var valL = dataRange[r][0].toString().trim(); // Module Part Number
    var valM = dataRange[r][1].toString().trim(); // Vision Part Number
    
    if (valL !== "" && valL !== "---") {
      // Description is blank ("") here. Will be filled by Step 3.
      moduleItems.push([valL, "", "1"]);
    }
    if (valM !== "" && valM !== "---") {
      visionItems.push([valM, "", "1"]);
    }
  }

  var sectionsToSync = [
    { destName: "MODULE", items: moduleItems },
    { destName: "VISION", items: visionItems }
  ];

  performSurgicalSync(destSheet, sectionsToSync);
}

// =========================================
// SYNC 3: BOM DESCRIPTION LOOKUP (NEW)
// =========================================
function updateDescriptionsFromBOM() {
  // *** UPDATE THIS IF YOUR BOM TAB NAME IS DIFFERENT ***
  var bomSheetName = "LATEST-BOM STRUCTURE TREE DIAGRAM";
  var destSheetName = "ORDERING LIST";
  
  // These are the ONLY sections we will touch
  var targetSections = ["MODULE", "VISION"];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bomSheet = ss.getSheetByName(bomSheetName);
  var destSheet = ss.getSheetByName(destSheetName);

  if (!bomSheet) {
    console.warn("BOM sheet not found: " + bomSheetName + ". Skipping Step 3.");
    return;
  }

  // 1. BUILD THE DICTIONARY (PART ID -> DESCRIPTION)
  // We scan the WHOLE sheet because data is scattered.
  var bomData = bomSheet.getDataRange().getValues();
  var bomMap = {};
  
  // Regex: Starts with 4 digits, contains a hyphen (e.g., 430001-A123, 2111-0054)
  var partIdPattern = /^\d{4,}-.*/;

  for (var r = 0; r < bomData.length; r++) {
    var row = bomData[r];
    // Loop through columns. Stop at length-1 because we look to the RIGHT.
    for (var c = 0; c < row.length - 1; c++) {
      var cellVal = row[c].toString().trim();
      
      if (partIdPattern.test(cellVal)) {
        var descriptionVal = row[c+1].toString().trim(); // Value to the RIGHT
        
        // Only save if description is not empty
        if (descriptionVal !== "" && descriptionVal !== "---") {
          // Store in map. If duplicate exists, later rows overwrite earlier ones.
          bomMap[cellVal] = descriptionVal;
        }
      }
    }
  }

  // 2. UPDATE DESTINATION SECTIONS
  for (var i = 0; i < targetSections.length; i++) {
    var sectionName = targetSections[i];
    console.log("Updating Descriptions for: " + sectionName);
    
    // A. FIND SECTION
    var textFinder = destSheet.getRange("A:A").createTextFinder(sectionName).matchEntireCell(true);
    var foundParams = textFinder.findAll();
    if (foundParams.length === 0) continue;
    
    var sectionStartRow = foundParams[0].getRow();
    
    // B. FIND ANCHOR ("DESCRIPTION" Header)
    var headerRow = -1;
    var checkRange = destSheet.getRange(sectionStartRow, 5, 20, 1).getValues();
    for (var r = 0; r < checkRange.length; r++) {
      if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
        headerRow = sectionStartRow + r;
        break;
      }
    }
    if (headerRow === -1) continue;

    var startDataRow = headerRow + 1;
    
    // C. ITERATE DATA ROWS AND UPDATE
    // We read Col D (ID) and Col E (Desc)
    // We process until we hit the next section or empty block
    var currentRow = startDataRow;
    var safetyLimit = 0;
    
    while (safetyLimit < 200) {
      var range = destSheet.getRange(currentRow, 1, 1, 5); // A to E
      var vals = range.getValues()[0];
      
      var colA = vals[0].toString();
      var partID = vals[3].toString().trim(); // Col D
      
      // Stop conditions
      if (colA !== "") break; // Hit next section title
      if (partID === "" && vals[4].toString() === "") break; // Empty row
      
      // D. PERFORM LOOKUP & OVERWRITE
      if (partID !== "" && bomMap.hasOwnProperty(partID)) {
        var newDesc = bomMap[partID];
        // Set value in Column E (Index 5)
        destSheet.getRange(currentRow, 5).setValue(newDesc);
      }
      
      currentRow++;
      safetyLimit++;
    }
  }
}


// =========================================
// SHARED HELPER: SURGICAL SYNC LOGIC
// =========================================
function performSurgicalSync(destSheet, sections) {
  
  for (var i = 0; i < sections.length; i++) {
    var procSec = sections[i];
    var sourceItems = procSec.items; // Expects [[ID, Desc, Qty], ...]
    var destSectionName = procSec.destName;
    
    console.log("Syncing Section: " + destSectionName);

    //A. FIND SECTION
    var textFinder = destSheet.getRange("A:A").createTextFinder(destSectionName).matchEntireCell(true);
    var foundParams = textFinder.findAll();
    
    if (foundParams.length === 0) {
      console.warn("Section " + destSectionName + " not found in Destination.");
      continue;
    }
    var sectionStartRow = foundParams[0].getRow();
    
    //B. FIND ANCHOR
    var headerRow = -1;
    // Look in Column 5 (E) for "DESCRIPTION"
    var checkRange = destSheet.getRange(sectionStartRow, 5, 20, 1).getValues(); 
    
    for (var r = 0; r < checkRange.length; r++) {
      if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
        headerRow = sectionStartRow + r;
        break;
      }
    }
    
    if (headerRow === -1) {
      console.warn("Header 'DESCRIPTION' not found for section " + destSectionName);
      continue;
    }

    //C. CALCULATE AVAILABLE SLOT SIZE
    var startWriteRow = headerRow + 1;
    var currentRow = startWriteRow;
    var existingDataCount = 0;
    
    while (existingDataCount < 200) {
      var rowVals = destSheet.getRange(currentRow, 1, 1, 5).getValues()[0];
      var colA_Val = rowVals[0].toString(); 
      var colD_Val = rowVals[3].toString().trim(); 
      var colE_Val = rowVals[4].toString().trim(); 
      
      if (colA_Val !== "") break; // Hit next section
      if (colD_Val === "" && colE_Val === "") break; // Hit empty row

      existingDataCount++;
      currentRow++;
    }
    
    //D. SYNC: INSERT OR DELETE
    var itemsNeeded = sourceItems.length;
    var itemsHave = existingDataCount;

    if (itemsNeeded > itemsHave) {
      destSheet.insertRowsAfter(startWriteRow + itemsHave - 1, itemsNeeded - itemsHave);
    } 
    else if (itemsNeeded < itemsHave) {
      destSheet.deleteRows(startWriteRow + itemsNeeded, itemsHave - itemsNeeded);
    }
    
    //E. WRITE DATA
    if (itemsNeeded > 0) {
      var outputBlock = [];
      for (var m = 0; m < itemsNeeded; m++) {
        outputBlock.push([
          m + 1,                // Item No (Col C)
          sourceItems[m][0],    // Part ID (Col D)
          sourceItems[m][1],    // Description (Col E)
          sourceItems[m][2]     // Qty (Col F)
        ]);
      }
      // Write to Col C, D, E, F (Index 3, 4 columns)
      destSheet.getRange(startWriteRow, 3, itemsNeeded, 4).setValues(outputBlock);

      //F. CHECKBOX ASSURANCE (Col G)
      var checkboxRange = destSheet.getRange(startWriteRow, 7, itemsNeeded, 1);
      checkboxRange.insertCheckboxes();
      
      var currentChecks = checkboxRange.getValues();
      var cleanChecks = [];
      var needsUpdate = false;

      for (var c = 0; c < currentChecks.length; c++) {
        var val = currentChecks[c][0];
        if (val === true || val === "TRUE") cleanChecks.push([true]);
        else if (val === false || val === "FALSE") cleanChecks.push([false]);
        else {
          cleanChecks.push([false]);
          needsUpdate = true;
        }
      }
      if (needsUpdate) checkboxRange.setValues(cleanChecks);

      //G. RELEASE TYPE VALIDATION (Col I)
      var releaseTypeRange = destSheet.getRange(startWriteRow, 9, itemsNeeded, 1);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['CHARGE OUT', 'MRP'], true)
        .setAllowInvalid(false)
        .build();
      releaseTypeRange.setDataValidation(rule);
    }
  }  
}
