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

    // 2. REFRESH REFERENCE DATA (Hidden Sheet)
    updateReferenceData(sourceSS, sourceSheet);

    // 3. RUN SYNC OPERATIONS
    
    // A. CORE
    updateSection_Core(sourceSheet, destSheet, "CORE", "CORE :430000-A557", 3, 4); 

    // B. CONFIG
    setupDropdownSection(destSheet, "CONFIG", "REF_DATA!A:A", "REF_DATA!A:B", null);

    // C. MODULE
    setupDropdownSection(destSheet, "MODULE", "REF_DATA!C:C", "REF_DATA!C:D", null);

    // D. VISION (Dropdown + Col B Category)
    // Lookup Range is REF_DATA!E:G -> Col E=ID, Col F=Desc, Col G=Category
    // We pass '3' as the categoryColIndex. This tells the function to grab the 3rd column (Category) for Col B.
    setupDropdownSection(destSheet, "VISION", "REF_DATA!E:E", "REF_DATA!E:G", 3);
    
    ui.alert("Sync Complete", "VISION section updated successfully.", ui.ButtonSet.OK);

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
  
  // 1. FETCH CONFIG DATA
  var configItems = fetchRawItems(sourceSheet, "OPTIONAL MODULE: 430001-A712", 6, 7, ["CONFIGURABLE MODULE"]);
  
  // 2. FETCH MODULE DATA
  var moduleItems = fetchRawItems(sourceSheet, "CONFIGURABLE MODULE: 430001-A713", 6, 7, ["CONFIGURABLE VISION MODULE"]);

  // 3. FETCH VISION DATA (With Category Logic)
  // Trigger: "CONFIGURABLE VISION MODULE" (Col F), Stop: "CALIBRATION JIG"
  var visionItems = fetchVisionItems(sourceSheet, "CONFIGURABLE VISION MODULE", 6, ["CALIBRATION JIG"]);

  // 4. WRITE TO REFERENCE SHEET
  if (configItems.length > 0) {
    refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems); // Col A, B
  }
  if (moduleItems.length > 0) {
    refSheet.getRange(1, 3, moduleItems.length, 2).setValues(moduleItems); // Col C, D
  }
  if (visionItems.length > 0) {
    // Writes to Col E, F, G. 
    // E=PartID, F=Description, G=Category
    refSheet.getRange(1, 5, visionItems.length, 3).setValues(visionItems); 
  }
}

function fetchVisionItems(sourceSheet, triggerPhrase, colID_Index, stopPhrases) {
  var lastRow = sourceSheet.getLastRow();
  var rangeValues = sourceSheet.getRange(1, 5, lastRow, 3).getValues(); // Cols E, F, G from Source
  
  var startRowIndex = -1;
  for (var i = 0; i < rangeValues.length; i++) {
    var val = rangeValues[i][1].toString().trim(); // Col F (Part ID)
    if (val.indexOf(triggerPhrase) > -1) {
      startRowIndex = i + 1; 
      break;
    }
  }

  if (startRowIndex === -1) return [];

  var items = [];
  var currentCategory = ""; 

  for (var k = startRowIndex; k < rangeValues.length; k++) {
    var valE = rangeValues[k][0].toString().trim(); // Category (Source Col E)
    var valF = rangeValues[k][1].toString().trim(); // Part ID (Source Col F)
    var valG = rangeValues[k][2].toString().trim(); // Description (Source Col G)

    if (stopPhrases && stopPhrases.length > 0) {
      if (stopPhrases.some(p => valF.indexOf(p) > -1 || valE.indexOf(p) > -1)) break;
    }

    if (valE !== "" && valE !== "---") {
      currentCategory = valE;
    }

    if (valF !== "" && valF !== "---" && valF.indexOf(":") === -1) {
      // We push 3 items: [Part ID, Description, Category]
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
  if (rowsToGrab < 1) return [];

  var idData = sourceSheet.getRange(startRowIndex + 1, colID, rowsToGrab, 1).getValues();
  var descData = sourceSheet.getRange(startRowIndex + 1, colDesc, rowsToGrab, 1).getValues();

  var items = [];
  
  for (var k = 0; k < idData.length; k++) {
    var pID = idData[k][0].toString().trim();
    var desc = descData[k][0].toString().trim();

    if (stopPhrases && stopPhrases.length > 0) {
      var hitStop = false;
      for (var s = 0; s < stopPhrases.length; s++) {
        if (pID.indexOf(stopPhrases[s]) > -1) {
          hitStop = true;
          break;
        }
      }
      if (hitStop) break;
    }
    if (pID.indexOf(":") > -1) break; 
    
    if (pID !== "" && pID !== "---") {
      items.push([pID, desc]); 
    }
  }
  return items;
}

function updateSection_Core(sourceSheet, destSheet, destHeaderName, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc) {
  var rawItems = fetchRawItems(sourceSheet, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc, []);
  var syncItems = rawItems.map(function(item) {
    return [item[0], item[1], "1"];
  });
  var sectionPayload = [{ destName: destHeaderName, items: syncItems }];
  performSurgicalSync(destSheet, sectionPayload);
}

// =========================================
// SYNC LOGIC: DROPDOWN SECTIONS (Config, Module, Vision)
// =========================================
function setupDropdownSection(destSheet, sectionName, dropdownRangeString, vlookupRangeString, categoryColIndex) {
  console.log("Setting up Dropdowns for: " + sectionName);

  // 1. Find Section Start
  var textFinder = destSheet.getRange("A:A").createTextFinder(sectionName).matchEntireCell(true);
  var foundParams = textFinder.findAll();
  if (foundParams.length === 0) return;
  var sectionStartRow = foundParams[0].getRow();
  
  // 2. Find Anchor (DESCRIPTION Header)
  var headerRow = -1;
  var checkRange = destSheet.getRange(sectionStartRow, 5, 20, 1).getValues(); 
  for (var r = 0; r < checkRange.length; r++) {
    if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
      headerRow = sectionStartRow + r;
      break;
    }
  }
  if (headerRow === -1) return;

  // *** HEADER RENAMING ***
  // If this section involves categories (like Vision), we rename Column B header for clarity.
  if (categoryColIndex != null) {
      destSheet.getRange(headerRow, 2).setValue("CATEGORY"); // Column B is index 2
  }

  var startWriteRow = headerRow + 1;
  var targetRows = 10; 

  // 3. Count Existing Rows
  var currentRow = startWriteRow;
  var existingDataCount = 0;
  var safetyLimit = 0;
  
  while (safetyLimit < 500) { 
    var rowVals = destSheet.getRange(currentRow, 1, 1, 5).getValues()[0];
    var colA_Val = rowVals[0].toString(); 
    var colD_Val = rowVals[3].toString().trim(); 
    var colE_Val = rowVals[4].toString().trim(); 
    
    if (colA_Val !== "") break; 
    if (colD_Val === "" && colE_Val === "") {
        var lookAhead = destSheet.getRange(currentRow, 4, 5, 1).getValues().flat();
        var hasData = lookAhead.some(r => r !== "");
        if (!hasData) break; 
    }
    existingDataCount++;
    currentRow++;
    safetyLimit++;
  }

  // 4. Force Row Count to Exactly 10
  if (existingDataCount > targetRows) {
    var rowsToDelete = existingDataCount - targetRows;
    destSheet.deleteRows(startWriteRow + targetRows, rowsToDelete);
  } 
  else if (existingDataCount < targetRows) {
    var rowsToInsert = targetRows - existingDataCount;
    destSheet.insertRowsAfter(startWriteRow + existingDataCount - 1, rowsToInsert);
  }

  // 5. Setup Dropdowns & Cleanup
  var dropdownRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRangeString), true)
    .setAllowInvalid(true) 
    .build();

  for (var i = 0; i < targetRows; i++) {
    var r = startWriteRow + i;
    var rangeD = destSheet.getRange(r, 4); // Col D (Part ID)
    var rangeE = destSheet.getRange(r, 5); // Col E (Description)
    
    // SMART CHECK:
    if (rangeD.getDataValidation() == null) {
      // No validation = Old Static Text. Wipe it.
      rangeD.clearContent();
      rangeE.clearContent(); 
      
      // ONLY clear Col B if this section uses it (e.g. Vision)
      if (categoryColIndex != null) {
         destSheet.getRange(r, 2).clearContent(); 
      }
    }
    
    // Apply Validation
    rangeD.setDataValidation(dropdownRule);
    
    // Set Description Formula (Col E)
    var cellD_Ref = "D" + r;
    var formulaDesc = '=IFERROR(VLOOKUP(' + cellD_Ref + ', ' + vlookupRangeString + ', 2, FALSE), "")';
    
    if (rangeE.getFormula() === "") {
        rangeE.setFormula(formulaDesc);
    }
    
    // Set Category Formula (Col B) - ONLY IF categoryColIndex IS PROVIDED
    // This inserts the formula into Col B that looks up the Category (index 3) from REF_DATA
    if (categoryColIndex != null) {
      var rangeB = destSheet.getRange(r, 2); // Column B
      var formulaCat = '=IFERROR(VLOOKUP(' + cellD_Ref + ', ' + vlookupRangeString + ', ' + categoryColIndex + ', FALSE), "")';
      rangeB.setFormula(formulaCat);
    }

    // Ensure item number
    destSheet.getRange(r, 3).setValue(i + 1);

    // Ensure Checkboxes
    var checkRange = destSheet.getRange(r, 7);
    if (checkRange.getDataValidation() == null) {
        checkRange.insertCheckboxes();
    }
    
    // Ensure Release Type
    var releaseRange = destSheet.getRange(r, 9);
    if (releaseRange.getDataValidation() == null) {
        var releaseRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['CHARGE OUT', 'MRP'], true)
        .setAllowInvalid(false)
        .build();
        releaseRange.setDataValidation(releaseRule);
    }
  }
}

function performSurgicalSync(destSheet, sections) {
  for (var i = 0; i < sections.length; i++) {
    var procSec = sections[i];
    var sourceItems = procSec.items; 
    var destSectionName = procSec.destName;
    
    var textFinder = destSheet.getRange("A:A").createTextFinder(destSectionName).matchEntireCell(true);
    var foundParams = textFinder.findAll();
    if (foundParams.length === 0) continue;
    var sectionStartRow = foundParams[0].getRow();
    
    var headerRow = -1;
    var checkRange = destSheet.getRange(sectionStartRow, 5, 20, 1).getValues(); 
    for (var r = 0; r < checkRange.length; r++) {
      if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
        headerRow = sectionStartRow + r;
        break;
      }
    }
    if (headerRow === -1) continue;

    var startWriteRow = headerRow + 1;
    var currentRow = startWriteRow;
    var existingDataCount = 0;
    
    while (existingDataCount < 200) {
      var rowVals = destSheet.getRange(currentRow, 1, 1, 5).getValues()[0];
      var colA_Val = rowVals[0].toString(); 
      var colD_Val = rowVals[3].toString().trim(); 
      var colE_Val = rowVals[4].toString().trim(); 
      
      if (colA_Val !== "") break; 
      if (colD_Val === "" && colE_Val === "") break; 

      existingDataCount++;
      currentRow++;
    }
    
    var itemsNeeded = sourceItems.length;
    var itemsHave = existingDataCount;

    if (itemsNeeded > itemsHave) {
      destSheet.insertRowsAfter(startWriteRow + itemsHave - 1, itemsNeeded - itemsHave);
    } 
    else if (itemsNeeded < itemsHave) {
      destSheet.deleteRows(startWriteRow + itemsNeeded, itemsHave - itemsNeeded);
    }
    
    if (itemsNeeded > 0) {
      var outputBlock = [];
      for (var m = 0; m < itemsNeeded; m++) {
        outputBlock.push([
          m + 1,                
          sourceItems[m][0],    
          sourceItems[m][1],    
          sourceItems[m][2]     
        ]);
      }
      destSheet.getRange(startWriteRow, 3, itemsNeeded, 4).setValues(outputBlock);

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

      var releaseTypeRange = destSheet.getRange(startWriteRow, 9, itemsNeeded, 1);
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['CHARGE OUT', 'MRP'], true)
        .setAllowInvalid(false)
        .build();
      releaseTypeRange.setDataValidation(rule);
    }
  }  
}
