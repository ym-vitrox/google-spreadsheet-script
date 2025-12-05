function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Refresh') 
    .addItem('Sync All Lists', 'runMasterSync') 
    .addToUi();
}

function runMasterSync() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // Run 1: PC & Config (Existing Logic)
    updateOrderingList_PC_Config();
    
    // Run 2: Layout Configuration (New Logic)
    updateOrderingList_Layout();
    
    // Notify User
    ui.alert("Sync Complete", "All sections have been updated successfully.", ui.ButtonSet.OK);
    
  } catch (e) {
    console.error(e);
    ui.alert("Error during Sync", e.message, ui.ButtonSet.OK);
  }
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
    console.warn("PC/Config Source or Dest sheet not found. Skipping this step.");
    return;
  }

  //1. READ SOURCE DATA 
  var lastRowSource = sourceSheet.getLastRow();
  //Safety check for empty sheet
  if (lastRowSource < 1) return;
  
  var sourceData = sourceSheet.getRange(1, 1, lastRowSource, 7).getValues(); // Cols A-G

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
    
    if (!sectionMapping.hasOwnProperty(section.name)) {
      continue;
    }

    var validItems = []; 
    var dataStarted = false;

    for (var k = 0; k < section.rows.length; k++) {
      var r = section.rows[k];
      var valE = r[4].toString().trim().toUpperCase(); 
      var valF = r[5].toString().trim().toUpperCase(); 

      if (!dataStarted) {
        if (valE === "PART ID" || valF === "DESCRIPTION") {
          dataStarted = true;
        }
      } else {
        var pID = r[4].toString();
        var desc = r[5].toString();
        var qty = r[6].toString(); 
        
        if (pID !== "PART ID" && desc !== "DESCRIPTION" && (pID !== "" || desc !== "")) {
           validItems.push([pID, desc, qty]);
        }
      }
    }
    
    //Fallback if headers weren't found but data exists
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
    
    processedSections.push({
      destName: sectionMapping[section.name], 
      items: validItems
    });
  }

  //EXECUTE SYNC
  performSurgicalSync(destSheet, processedSections);
}

// =========================================
// SYNC 2: LAYOUT CONFIGURATION (NEW LOGIC)
// =========================================
function updateOrderingList_Layout() {
  // *** PLEASE UPDATE THIS NAME TO MATCH YOUR ACTUAL SHEET NAME ***
  var layoutSheetName = "LAYOUT CONFIGURATION"; 
  var destSheetName = "ORDERING LIST";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(layoutSheetName);
  var destSheet = ss.getSheetByName(destSheetName);

  if (!sourceSheet) {
    console.warn("Layout Configuration sheet not found: " + layoutSheetName);
    return;
  }

  // 1. FIND THE TRIGGER ROW ("CONFIGURATION")
  var lastRow = sourceSheet.getLastRow();
  var colA_Values = sourceSheet.getRange(1, 1, lastRow, 1).getValues(); // Get Col A only
  
  var startRow = -1;
  for (var i = 0; i < colA_Values.length; i++) {
    if (colA_Values[i][0].toString().trim().toUpperCase() === "CONFIGURATION") {
      startRow = i + 1; // Row index is 1-based
      break;
    }
  }

  if (startRow === -1) {
    console.warn("Keyword 'CONFIGURATION' not found in Column A of Layout sheet.");
    return;
  }

  // 2. EXTRACT DATA FROM TRIGGER ROW DOWNWARDS
  // We need Column L (Index 11) and Column M (Index 12)
  // getRange(row, col, numRows, numCols)
  var rowsToProcess = lastRow - startRow + 1;
  if (rowsToProcess < 1) return;

  // Grab data starting from the "CONFIGURATION" row
  // Col L is the 12th column, Col M is the 13th column.
  var dataRange = sourceSheet.getRange(startRow, 12, rowsToProcess, 2).getValues(); 

  var moduleItems = [];
  var visionItems = [];

  for (var r = 0; r < dataRange.length; r++) {
    var valL = dataRange[r][0].toString().trim(); // Module Part Number
    var valM = dataRange[r][1].toString().trim(); // Vision Part Number
    
    // LOGIC FOR MODULE LIST (From Col L)
    // Ignore empty and "---"
    if (valL !== "" && valL !== "---") {
      // Structure: [PartID, Description, Qty]
      // Description is intentionally BLANK ("") for now
      moduleItems.push([valL, "", "1"]); 
    }

    // LOGIC FOR VISION LIST (From Col M)
    // Ignore empty and "---"
    if (valM !== "" && valM !== "---") {
      visionItems.push([valM, "", "1"]); 
    }
  }

  // 3. PREPARE SECTIONS FOR SYNC
  var sectionsToSync = [
    { destName: "MODULE", items: moduleItems },
    { destName: "VISION", items: visionItems }
  ];

  // 4. EXECUTE SYNC
  performSurgicalSync(destSheet, sectionsToSync);
}


// =========================================
// SHARED HELPER: SURGICAL SYNC LOGIC
// =========================================
// This performs the finding, calculating, inserting/deleting, and formatting
function performSurgicalSync(destSheet, sections) {
  
  for (var i = 0; i < sections.length; i++) {
    var procSec = sections[i];
    var sourceItems = procSec.items;
    var destSectionName = procSec.destName;
    
    console.log("Syncing Section: " + destSectionName);

    //A. FIND SECTION
    var textFinder = destSheet.getRange("A:A").createTextFinder(destSectionName).matchEntireCell(true);
    var foundParams = textFinder.findAll();
    
    if (foundParams.length === 0) {
      console.log("Section not found in destination: " + destSectionName);
      continue;
    }
    
    var sectionStartRow = foundParams[0].getRow();
    
    //B. FIND ANCHOR (Look for "DESCRIPTION" in headers below section title)
    var headerRow = -1;
    var searchLimit = 20; 
    var checkRange = destSheet.getRange(sectionStartRow, 5, searchLimit, 1).getValues(); //Check Col E
    
    for (var r = 0; r < checkRange.length; r++) {
      if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
        headerRow = sectionStartRow + r;
        break;
      }
    }
    
    if (headerRow === -1) {
      console.log("Anchor 'DESCRIPTION' not found for: " + destSectionName);
      continue;
    }

    //C. CALCULATE AVAILABLE SLOT SIZE
    var startWriteRow = headerRow + 1;
    var currentRow = startWriteRow;
    var existingDataCount = 0;
    
    //Safe loop to find end of section (Stop at empty A or empty D&E)
    while (existingDataCount < 200) {
      var rowVals = destSheet.getRange(currentRow, 1, 1, 5).getValues()[0];
      
      var colA_Val = rowVals[0].toString(); 
      var colD_Val = rowVals[3].toString().trim(); 
      var colE_Val = rowVals[4].toString().trim(); 
      
      // Stop if we hit a new section Title in Col A
      if (colA_Val !== "") break;
      // Stop if we hit an empty row (both ID and Desc empty)
      if (colD_Val === "" && colE_Val === "") break; 

      existingDataCount++;
      currentRow++;
    }
    
    //D. SYNC: INSERT OR DELETE ROWS
    var itemsNeeded = sourceItems.length;
    var itemsHave = existingDataCount;
    
    if (itemsNeeded > itemsHave) {
      var rowsToAdd = itemsNeeded - itemsHave;
      destSheet.insertRowsAfter(startWriteRow + itemsHave - 1, rowsToAdd);
    } 
    else if (itemsNeeded < itemsHave) {
      var rowsToDelete = itemsHave - itemsNeeded;
      destSheet.deleteRows(startWriteRow + itemsNeeded, rowsToDelete);
    }
    
    //E. WRITE DATA
    if (itemsNeeded > 0) {
      var outputBlock = [];
      for (var m = 0; m < itemsNeeded; m++) {
        outputBlock.push([
          m + 1,                //Col C: Item #
          sourceItems[m][0],    //Col D: Part ID
          sourceItems[m][1],    //Col E: Description (Might be empty)
          sourceItems[m][2]     //Col F: QTY
        ]);
      }
      destSheet.getRange(startWriteRow, 3, itemsNeeded, 4).setValues(outputBlock);

      //F. CHECKBOX ASSURANCE (Col G)
      var checkboxRange = destSheet.getRange(startWriteRow, 7, itemsNeeded, 1);
      checkboxRange.insertCheckboxes();
      
      //Normalize Checkboxes
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
