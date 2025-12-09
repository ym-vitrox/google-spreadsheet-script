function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Refresh') 
    .addItem('Sync All Lists', 'runMasterSync') 
    .addToUi();
}

function runMasterSync() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    // 1. SETUP SHARED RESOURCES
    var sourceSpreadsheetId = "1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA"; 
    var sourceTabName = "BOM Structure Tree Diagram";
    
    try {
      var sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
      var sourceSheet = sourceSS.getSheetByName(sourceTabName);
      if (!sourceSheet) throw new Error("Source tab '" + sourceTabName + "' not found.");
    } catch (e) {
      throw new Error("Could not open Source Spreadsheet. Check ID/Permissions. " + e.message);
    }
    
    var destSS = SpreadsheetApp.getActiveSpreadsheet();
    var destSheet = destSS.getSheetByName("ORDERING LIST");
    if (!destSheet) throw new Error("Destination sheet 'ORDERING LIST' not found.");

    // 2. REFRESH REFERENCE DATA (Hidden Sheet REF_DATA)
    updateReferenceData(sourceSS, sourceSheet);

    // 3. RUN SYNC OPERATIONS
    
    // A. CORE SECTION (Surgical Sync: Row insertion/deletion)
    updateSection_Core(sourceSheet, destSheet, "CORE", "CORE :430000-A557", 3, 4); 

    // B. CONFIG SECTION (Dropdown Setup)
    setupDropdownSection(destSheet, "CONFIG", "REF_DATA!A:A", "REF_DATA!A:B", null);

    // C. MODULE SECTION (Dropdown Setup)
    setupDropdownSection(destSheet, "MODULE", "REF_DATA!C:C", "REF_DATA!C:D", null);

    // D. VISION SECTION (Dropdown + Col B Category Display)
    setupDropdownSection(destSheet, "VISION", "REF_DATA!E:E", "REF_DATA!E:G", 3);
    
    ui.alert("Sync Complete", "All sections updated. Spacer rows preserved.", ui.ButtonSet.OK);

  } catch (e) {
    console.error(e);
    ui.alert("Error during Sync", e.message, ui.ButtonSet.OK);
  }
}

// =========================================
// HELPER: REFERENCE DATA MANAGER
// =========================================
function updateReferenceData(sourceSS, sourceSheet) {
  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var refSheetName = "REF_DATA";
  var refSheet = destSS.getSheetByName(refSheetName);
  
  if (!refSheet) {
    refSheet = destSS.insertSheet(refSheetName);
    refSheet.hideSheet();
  }
  
  refSheet.clear(); 
  
  // 1. FETCH CONFIG DATA (Cols A, B)
  var configItems = fetchRawItems(sourceSheet, "OPTIONAL MODULE: 430001-A712", 6, 7, ["CONFIGURABLE MODULE"]);
  
  // 2. FETCH MODULE DATA (Cols C, D)
  var moduleItems = fetchRawItems(sourceSheet, "CONFIGURABLE MODULE: 430001-A713", 6, 7, ["CONFIGURABLE VISION MODULE"]);

  // 3. FETCH VISION DATA (Cols E, F, G - includes Category)
  var visionItems = fetchVisionItems(sourceSheet, "CONFIGURABLE VISION MODULE", 6, ["CALIBRATION JIG"]);

  // 4. WRITE TO REFERENCE SHEET
  if (configItems.length > 0) {
    refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems); 
  }
  if (moduleItems.length > 0) {
    refSheet.getRange(1, 3, moduleItems.length, 2).setValues(moduleItems); 
  }
  if (visionItems.length > 0) {
    refSheet.getRange(1, 5, visionItems.length, 3).setValues(visionItems); 
  }
}

// =========================================
// FETCHERS (Scraping Source Data)
// =========================================

function fetchVisionItems(sourceSheet, triggerPhrase, colID_Index, stopPhrases) {
  var lastRow = sourceSheet.getLastRow();
  var rangeValues = sourceSheet.getRange(1, 5, lastRow, 3).getValues(); // Cols E, F, G
  
  var startRowIndex = -1;
  for (var i = 0; i < rangeValues.length; i++) {
    var val = rangeValues[i][1].toString().trim(); // Trigger in Col F
    if (val.indexOf(triggerPhrase) > -1) {
      startRowIndex = i + 1; 
      break;
    }
  }
  if (startRowIndex === -1) return [];

  var items = [];
  var currentCategory = ""; 

  for (var k = startRowIndex; k < rangeValues.length; k++) {
    var valE = rangeValues[k][0].toString().trim(); // Category (Col E)
    var valF = rangeValues[k][1].toString().trim(); // Part ID (Col F)
    var valG = rangeValues[k][2].toString().trim(); // Description (Col G)

    if (stopPhrases && stopPhrases.some(p => valF.indexOf(p) > -1 || valE.indexOf(p) > -1)) break;

    if (valE !== "" && valE !== "---") {
      currentCategory = valE;
    }

    if (valF !== "" && valF !== "---" && valF.indexOf(":") === -1) {
      items.push([valF, valG, currentCategory]);
    }
  }
  return items;
}

function fetchRawItems(sourceSheet, triggerPhrase, colID, colDesc, stopPhrases) {
  var lastRow = sourceSheet.getLastRow();
  var rangeValues = sourceSheet.getRange(1, colID, lastRow, 1).getValues();
  
  var startRowIndex = -1;
  for (var i = 0; i < rangeValues.length; i++) {
    var val = rangeValues[i][0].toString().trim();
    if (val.indexOf(triggerPhrase) > -1) {
      startRowIndex = i + 1; 
      break;
    }
  }
  if (startRowIndex === -1) return [];

  var rowsToGrab = lastRow - startRowIndex;
  var idData = sourceSheet.getRange(startRowIndex + 1, colID, rowsToGrab, 1).getValues();
  var descData = sourceSheet.getRange(startRowIndex + 1, colDesc, rowsToGrab, 1).getValues();

  var items = [];
  for (var k = 0; k < idData.length; k++) {
    var pID = idData[k][0].toString().trim();
    var desc = descData[k][0].toString().trim();

    if (stopPhrases && stopPhrases.some(s => pID.indexOf(s) > -1)) break;
    if (pID.indexOf(":") > -1) break; 
    
    if (pID !== "" && pID !== "---") {
      items.push([pID, desc]); 
    }
  }
  return items;
}

// =========================================
// SECTION LOGIC: CORE (Direct Sync)
// =========================================
function updateSection_Core(sourceSheet, destSheet, destHeaderName, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc) {
  var rawItems = fetchRawItems(sourceSheet, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc, []);
  var syncItems = rawItems.map(function(item) {
    return [item[0], item[1], "1"];
  });
  var sectionPayload = [{ destName: destHeaderName, items: syncItems }];
  performSurgicalSync(destSheet, sectionPayload);
}

// =========================================
// SECTION LOGIC: DROPDOWNS (Config, Module, Vision)
// =========================================
function setupDropdownSection(destSheet, sectionName, dropdownRangeString, vlookupRangeString, categoryColIndex) {
  var textFinder = destSheet.getRange("A:A").createTextFinder(sectionName).matchEntireCell(true);
  var foundParams = textFinder.findAll();
  if (foundParams.length === 0) return;
  var sectionStartRow = foundParams[0].getRow();
  
  var headerRow = -1;
  var checkRange = destSheet.getRange(sectionStartRow, 5, 20, 1).getValues(); 
  for (var r = 0; r < checkRange.length; r++) {
    if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
      headerRow = sectionStartRow + r;
      break;
    }
  }
  if (headerRow === -1) return;

  // HEADER RENAME (Vision Only)
  if (categoryColIndex != null) {
      destSheet.getRange(headerRow, 2).setValue("CATEGORY");
  }

  var startWriteRow = headerRow + 1;
  var targetRows = 10; 

  // Count Existing Rows to preserve exact count
  var currentRow = startWriteRow;
  var existingDataCount = 0;
  var safetyLimit = 0;
  
  while (safetyLimit < 500) { 
    var rowVals = destSheet.getRange(currentRow, 1, 1, 5).getValues()[0];
    var colA_Val = rowVals[0].toString();
    
    // Stop if we hit a new section header immediately
    if (colA_Val !== "") break; 

    var colD_Val = rowVals[3].toString().trim(); 
    var colE_Val = rowVals[4].toString().trim(); 

    // If current row appears empty, Look Ahead to see if it's just a gap or end of section
    if (colD_Val === "" && colE_Val === "") {
        // Look ahead 5 rows, checking Cols A (Header) and D (Data)
        // We get 5 rows, starting from current, spanning 4 columns (A, B, C, D)
        var lookAheadRange = destSheet.getRange(currentRow, 1, 5, 4).getValues();
        
        var foundValidData = false;
        for (var k = 0; k < lookAheadRange.length; k++) {
            var nextA = lookAheadRange[k][0].toString(); // Header Column
            var nextD = lookAheadRange[k][3].toString(); // Part ID Column
            
            // BOUNDARY CHECK: If we hit a row with a Header (Col A not empty), 
            // we have hit the next section. STOP LOOKING. The current section ends here.
            if (nextA !== "") {
                break; 
            }
            
            // If we find data in Col D (and no header in A), it is valid data for US.
            if (nextD !== "") {
                foundValidData = true;
                break;
            }
        }
        
        // If we didn't find any valid data in the look-ahead window, we are done.
        if (!foundValidData) break; 
    }
    
    existingDataCount++;
    currentRow++;
    safetyLimit++;
  }

  // Row Management
  if (existingDataCount > targetRows) {
    destSheet.deleteRows(startWriteRow + targetRows, existingDataCount - targetRows);
  } else if (existingDataCount < targetRows) {
    destSheet.insertRowsAfter(startWriteRow + existingDataCount - 1, targetRows - existingDataCount);
  }

  // Dropdown & Cleanup Rule
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRangeString), true)
    .setAllowInvalid(true).build();

  for (var i = 0; i < targetRows; i++) {
    var r = startWriteRow + i;
    var rangeD = destSheet.getRange(r, 4);
    var rangeE = destSheet.getRange(r, 5);
    
    if (rangeD.getDataValidation() == null) {
      rangeD.clearContent();
      rangeE.clearContent(); 
      if (categoryColIndex != null) destSheet.getRange(r, 2).clearContent(); 
    }
    
    rangeD.setDataValidation(dropdownRule);
    var cellD_Ref = "D" + r;
    
    // Description Formula
    if (rangeE.getFormula() === "") {
        rangeE.setFormula('=IFERROR(VLOOKUP(' + cellD_Ref + ', ' + vlookupRangeString + ', 2, FALSE), "")');
    }
    
    // Category Formula (Vision Only)
    if (categoryColIndex != null) {
      destSheet.getRange(r, 2).setFormula('=IFERROR(VLOOKUP(' + cellD_Ref + ', ' + vlookupRangeString + ', ' + categoryColIndex + ', FALSE), "")');
    }

    // Static Layout columns
    destSheet.getRange(r, 3).setValue(i + 1);
    if (destSheet.getRange(r, 7).getDataValidation() == null) destSheet.getRange(r, 7).insertCheckboxes();
    if (destSheet.getRange(r, 9).getDataValidation() == null) {
        destSheet.getRange(r, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
    }
  }
}

// =========================================
// SYNC ENGINE: SURGICAL INSERT/DELETE
// =========================================
function performSurgicalSync(destSheet, sections) {
  for (var i = 0; i < sections.length; i++) {
    var procSec = sections[i];
    var sourceItems = procSec.items; 
    var textFinder = destSheet.getRange("A:A").createTextFinder(procSec.destName).matchEntireCell(true);
    var foundParams = textFinder.findAll();
    if (foundParams.length === 0) continue;
    var startWriteRow = -1;
    var checkRange = destSheet.getRange(foundParams[0].getRow(), 5, 20, 1).getValues(); 
    for (var r = 0; r < checkRange.length; r++) {
      if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
        startWriteRow = foundParams[0].getRow() + r + 1;
        break;
      }
    }
    if (startWriteRow === -1) continue;

    var existingDataCount = 0;
    while (existingDataCount < 200) {
      var rowVals = destSheet.getRange(startWriteRow + existingDataCount, 1, 1, 5).getValues()[0];
      if (rowVals[0].toString() !== "" || (rowVals[3].toString().trim() === "" && rowVals[4].toString().trim() === "")) break;
      existingDataCount++;
    }
    
    if (sourceItems.length > existingDataCount) {
      destSheet.insertRowsAfter(startWriteRow + existingDataCount - 1, sourceItems.length - existingDataCount);
    } else if (sourceItems.length < existingDataCount) {
      destSheet.deleteRows(startWriteRow + sourceItems.length, existingDataCount - sourceItems.length);
    }
    
    if (sourceItems.length > 0) {
      var outputBlock = sourceItems.map((item, m) => [m + 1, item[0], item[1], item[2]]);
      destSheet.getRange(startWriteRow, 3, sourceItems.length, 4).setValues(outputBlock);
      destSheet.getRange(startWriteRow, 7, sourceItems.length, 1).insertCheckboxes();
      destSheet.getRange(startWriteRow, 9, sourceItems.length, 1).setDataValidation(
         SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
    }
  }  
}
