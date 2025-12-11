/**
 * ORDERING LIST SCRIPT
 * Features:
 * 1. Master Sync from Source BOM.
 * 2. Incremental Dependency Logic (Live Kit Insertion) via onEdit (Module Section).
 * 3. Shopping List Logic (Basic Tool & Pneumatic) via onEdit (Config Section).
 * 4. Renumbering Tool (Menu Option).
 */

// =========================================
// 1. LIVE TRIGGER (Handle Kit/Tool Insertion)
// =========================================
function onEdit(e) {
  if (!e) return;
  
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "ORDERING LIST") return;

  var range = e.range;
  var row = range.getRow();
  var col = range.getColumn();
  
  // We strictly look for edits in Column D (Part ID)
  if (col !== 4) return;

  var newVal = e.value; 
  var oldVal = e.oldValue; 
  
  // --- LOCATE SECTIONS ---
  var configFinder = sheet.getRange("A:A").createTextFinder("CONFIG").matchEntireCell(true).findNext();
  var moduleFinder = sheet.getRange("A:A").createTextFinder("MODULE").matchEntireCell(true).findNext();
  var visionFinder = sheet.getRange("A:A").createTextFinder("VISION").matchEntireCell(true).findNext();
  
  // Define Boundaries
  var configStart = configFinder ? configFinder.getRow() + 1 : 0;
  var configEnd = moduleFinder ? moduleFinder.getRow() - 1 : 0;
  
  var moduleStart = moduleFinder ? moduleFinder.getRow() + 1 : 0;
  var moduleEnd = visionFinder ? visionFinder.getRow() - 1 : 0;

  // =========================================================
  // LOGIC A: CONFIG SECTION (Shopping Lists)
  // =========================================================
  if (row >= configStart && row <= configEnd && configStart > 0) {
    
    var BASIC_TOOL_TRIGGER = "430001-A378";
    var PNEUMATIC_TRIGGER = "430001-A714";

    // --- 1. DELETE LOGIC (If Trigger is Removed/Switched) ---
    if (oldVal === BASIC_TOOL_TRIGGER) {
      if (row + 10 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 10);
    }
    if (oldVal === PNEUMATIC_TRIGGER) {
      if (row + 3 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 3);
    }

    // --- 2. INSERT LOGIC (If Trigger is Selected) ---
    
    // CASE A: BASIC TOOL (10 Rows)
    if (newVal === BASIC_TOOL_TRIGGER) {
      insertShoppingList(sheet, row, 10, "REF_DATA!I:I", "REF_DATA!I:J");
    }

    // CASE B: PNEUMATIC MODULE (3 Rows)
    if (newVal === PNEUMATIC_TRIGGER) {
      insertShoppingList(sheet, row, 3, "REF_DATA!K:K", "REF_DATA!K:L");
    }
  }

  // =========================================================
  // LOGIC B: MODULE SECTION (Incremental Dependencies)
  // =========================================================
  if (row >= moduleStart && row <= moduleEnd && moduleStart > 0) {
    
    var KIT_DEPENDENCIES = getKitDependencies(); 

    // 1. CLEANUP (Handle Removal or Swap)
    if (oldVal && KIT_DEPENDENCIES[oldVal]) {
      var checkRow = row + 1;
      if (checkRow <= sheet.getMaxRows()) {
        var childPartID = sheet.getRange(checkRow, 4).getValue();
        var possibleChildren = KIT_DEPENDENCIES[oldVal].map(function(k) { return k.id; });
        
        if (possibleChildren.includes(childPartID)) {
          sheet.deleteRow(checkRow);
        }
      }
    }

    // 2. INSERTION (Handle New Selection)
    if (newVal && KIT_DEPENDENCIES[newVal]) {
      var parentID = newVal;
      var childList = KIT_DEPENDENCIES[parentID];
      
      // --- GAP FILLING LOGIC (Prevent Duplicates) ---
      var scanHeight = moduleEnd - moduleStart + 5; 
      var sectionValues = sheet.getRange(moduleStart, 4, scanHeight, 1).getValues();
      var usedKitIds = [];

      for (var i = 0; i < sectionValues.length; i++) {
        var scanRowAbs = moduleStart + i;
        var scanVal = sectionValues[i][0];

        if (scanVal == parentID && scanRowAbs !== row) {
          if (i + 1 < sectionValues.length) {
            var potentialChildId = sectionValues[i + 1][0];
            var isKnownChild = childList.some(function(k) { return k.id === potentialChildId; });
            if (isKnownChild) usedKitIds.push(potentialChildId);
          }
        }
      }

      var targetKit = null;
      for (var k = 0; k < childList.length; k++) {
        if (usedKitIds.indexOf(childList[k].id) === -1) {
          targetKit = childList[k];
          break; 
        }
      }

      if (!targetKit) targetKit = childList[childList.length - 1];
      
      // Check row below
      var checkRow = row + 1;
      var rowBelowData = sheet.getRange(checkRow, 3, 1, 3).getValues()[0]; 
      var valBelowD = rowBelowData[1]; 

      if (valBelowD === targetKit.id) return;

      sheet.insertRowAfter(row);
      
      // Populate
      var kitPartIdCell = sheet.getRange(checkRow, 4);
      var kitDescCell = sheet.getRange(checkRow, 5);

      kitPartIdCell.setValue(targetKit.id);   
      kitDescCell.setValue(targetKit.desc);   
      
      kitPartIdCell.clearDataValidations(); 
      
      sheet.getRange(checkRow, 3).clearContent(); 
      sheet.getRange(checkRow, 7).insertCheckboxes(); 
      var releaseRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['CHARGE OUT', 'MRP'], true).build();
      sheet.getRange(checkRow, 9).setDataValidation(releaseRule);
    }
  }
}

// =========================================
// HELPER: SHOPPING LIST INSERTION
// =========================================
function insertShoppingList(sheet, row, count, dropdownRef, vlookupRef) {
  sheet.insertRowsAfter(row, count);
  var startInsertRow = row + 1;
  
  // A. Set Dropdowns (Column D)
  var dropDownRange = sheet.getRange(startInsertRow, 4, count, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRef), true)
    .setAllowInvalid(true).build();
  dropDownRange.setDataValidation(rule);
  
  // B. Set Descriptions (Column E) with VLOOKUP
  var descRange = sheet.getRange(startInsertRow, 5, count, 1);
  var formulas = [];
  for (var i = 0; i < count; i++) {
    var r = startInsertRow + i;
    formulas.push(['=IFERROR(VLOOKUP(D' + r + ', ' + vlookupRef + ', 2, FALSE), "")']);
  }
  descRange.setFormulas(formulas);
  
  // C. Set Checkboxes (Column G)
  sheet.getRange(startInsertRow, 7, count, 1).insertCheckboxes();
  
  // D. Set Release Type (Column I)
  var releaseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CHARGE OUT', 'MRP'], true).build();
  sheet.getRange(startInsertRow, 9, count, 1).setDataValidation(releaseRule);
  
  // E. Ensure Qty (Col F) is blank
  sheet.getRange(startInsertRow, 6, count, 1).clearContent();
}

// =========================================
// 2. STANDARD MENUS & SYNC
// =========================================
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Refresh') 
    .addItem('Sync All Lists', 'runMasterSync') 
    .addSeparator()
    .addItem('Renumber Kits (Tidy Up)', 'renumberKits') 
    .addToUi();
}

// =========================================
// 3. RENUMBERING FUNCTION (Module Section Only)
// =========================================
function renumberKits() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORDERING LIST");
  if (!sheet) return;

  var moduleFinder = sheet.getRange("A:A").createTextFinder("MODULE").matchEntireCell(true).findNext();
  var visionFinder = sheet.getRange("A:A").createTextFinder("VISION").matchEntireCell(true).findNext();
  
  if (!moduleFinder || !visionFinder) return;
  var startRow = moduleFinder.getRow() + 1; 
  var endRow = visionFinder.getRow() - 1;

  var range = sheet.getRange(startRow, 4, endRow - startRow + 1, 1);
  var values = range.getValues();
  var KIT_DEPENDENCIES = getKitDependencies();

  var parentCounts = {};
  for (var key in KIT_DEPENDENCIES) {
    parentCounts[key] = 0;
  }

  for (var i = 0; i < values.length; i++) {
    var parentID = values[i][0];
    
    if (KIT_DEPENDENCIES[parentID]) {
      parentCounts[parentID]++;
      var count = parentCounts[parentID];
      
      var childList = KIT_DEPENDENCIES[parentID];
      var kitIndex = count - 1;
      if (kitIndex >= childList.length) kitIndex = childList.length - 1; 
      var targetKit = childList[kitIndex];

      var currentRowAbs = startRow + i;
      var childRowAbs = currentRowAbs + 1;
      
      if (childRowAbs > endRow + 10) continue; 

      var actualChildID = sheet.getRange(childRowAbs, 4).getValue();
      var isRowHoldingChild = childList.some(function(k) { return k.id === actualChildID; });

      if (isRowHoldingChild) {
        if (actualChildID !== targetKit.id) {
          sheet.getRange(childRowAbs, 4).setValue(targetKit.id);
          sheet.getRange(childRowAbs, 5).setValue(targetKit.desc);
          sheet.getRange(childRowAbs, 4).clearDataValidations(); 
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert("Renumbering Complete", "Kits have been sorted sequentially (1, 2, 3...).", SpreadsheetApp.getUi().ButtonSet.OK);
}

// =========================================
// SHARED DATA
// =========================================
function getKitDependencies() {
  return {
    "430000-A973": [ 
      {id: "430001-A529", desc: "Kit-Misc. Ele. Reject Bin 1"},
      {id: "430001-A530", desc: "Kit-Misc. Ele. Reject Bin 2"},
      {id: "430001-A531", desc: "Kit-Misc. Ele. Reject Bin 3"},
      {id: "430001-A532", desc: "Kit-Misc. Ele. Reject Bin 4"}
    ],
    "430000-A959": [
      {id: "430000-A989", desc: "Kit-Misc. Ele. Dynamic Recentering V1-#1"},
      {id: "430001-A373", desc: "Kit-Misc. Ele. Dynamic Recentering V1-#2"}
    ],
    "430001-A229": [
      {id: "430001-A201", desc: "Kit-Misc. Ele. Direct Side Wall 1"},
      {id: "430001-A239", desc: "Kit-Misc. Ele. Direct Side Wall 2"}
    ],
    "430001-A276": [
      {id: "430001-A247", desc: "Kit-Misc. Ele. Rotary Module 1"},
      {id: "430001-A257", desc: "Kit-Misc. Ele. Rotary Module 2"}
    ],
    "430000-A974": [
      {id: "430001-A505", desc: "Kit-Misc. Ele. Dynamic Recentering V2-#1"},
      {id: "430001-A506", desc: "Kit-Misc. Ele. Dynamic Recentering V2-#2"}
    ]
  };
}

// =========================================
// 4. MASTER SYNC LOGIC
// =========================================
function runMasterSync() {
  var ui = SpreadsheetApp.getUi();
  
  try {
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

    // Update Main REF_DATA (Config/Module parents)
    updateReferenceData(sourceSS, sourceSheet);
    
    // Update REF_DATA for Shopping Lists (Basic Tool & Pneumatic)
    updateShoppingLists(sourceSheet);

    // Sync other sections
    updateSection_Core(sourceSheet, destSheet, "CORE", "CORE :430000-A557", 3, 4); 
    setupDropdownSection(destSheet, "CONFIG", "REF_DATA!A:A", "REF_DATA!A:B", null);
    setupDropdownSection(destSheet, "MODULE", "REF_DATA!C:C", "REF_DATA!C:D", null);
    setupDropdownSection(destSheet, "VISION", "REF_DATA!E:E", "REF_DATA!E:G", 3);
    
    ui.alert("Sync Complete", "Lists updated, including Basic Tool & Pneumatic Options.", ui.ButtonSet.OK);

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
  
  // Clear the Main Area (A-H), preserving I-L for the separate function
  refSheet.getRange("A:H").clear(); 
  
  var configItems = fetchRawItems(sourceSheet, "OPTIONAL MODULE: 430001-A712", 6, 7, ["CONFIGURABLE MODULE"]);
  var moduleItems = fetchRawItems(sourceSheet, "CONFIGURABLE MODULE: 430001-A713", 6, 7, ["CONFIGURABLE VISION MODULE"]);
  var visionItems = fetchVisionItems(sourceSheet, "CONFIGURABLE VISION MODULE", 6, ["CALIBRATION JIG"]);

  if (configItems.length > 0) refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems);
  if (moduleItems.length > 0) refSheet.getRange(1, 3, moduleItems.length, 2).setValues(moduleItems);
  if (visionItems.length > 0) refSheet.getRange(1, 5, visionItems.length, 3).setValues(visionItems);
}

function updateShoppingLists(sourceSheet) {
  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = destSS.getSheetByName("REF_DATA");
  
  // Clear Shopping List Area (Col I-L)
  refSheet.getRange("I:L").clear();

  // 1. Basic Tool Module (A378) -> Col I, J
  // Logic: Stop at "List-"
  var basicItems = fetchShoppingListItems(sourceSheet, "List-Optional Basic Tool Module: 430001-A378", 12, 13, "STRICT");
  if (basicItems.length > 0) {
    refSheet.getRange(1, 9, basicItems.length, 2).setValues(basicItems);
  }

  // 2. Pneumatic Module (A714) -> Col K, L
  // Logic: Skip Empty, Stop at "List-" or "Non-Numeric"
  var pneumaticItems = fetchShoppingListItems(sourceSheet, "List-Optional Pneumatic Module : 430001-A714", 12, 13, "SKIP_EMPTY");
  if (pneumaticItems.length > 0) {
    refSheet.getRange(1, 11, pneumaticItems.length, 2).setValues(pneumaticItems);
  }
}

// =========================================
// FETCHERS
// =========================================

// *** UNIFIED SHOPPING LIST FETCHER ***
function fetchShoppingListItems(sourceSheet, triggerPhrase, colID, colDesc, stopMode) {
  var lastRow = sourceSheet.getLastRow();
  var idColumnVals = sourceSheet.getRange(1, colID, lastRow, 1).getValues();
  
  // Find Start
  var startRowIndex = -1;
  for (var i = 0; i < idColumnVals.length; i++) {
    if (idColumnVals[i][0].toString().trim().indexOf(triggerPhrase) > -1) {
      startRowIndex = i + 1; 
      break;
    }
  }
  
  if (startRowIndex === -1) return [];

  var rowsRemaining = lastRow - startRowIndex;
  if (rowsRemaining < 1) return [];
  
  var idData = sourceSheet.getRange(startRowIndex + 1, colID, rowsRemaining, 1).getValues();
  var descData = sourceSheet.getRange(startRowIndex + 1, colDesc, rowsRemaining, 1).getValues();
  
  var items = [];
  
  for (var k = 0; k < idData.length; k++) {
    var pID = idData[k][0].toString().trim();
    var desc = descData[k][0].toString().trim();
    
    // --- STOP CHECK ---
    if (pID.toUpperCase().indexOf("LIST-") === 0) {
      break;
    }
    
    // --- MODE: STRICT (Basic Tool) ---
    if (stopMode === "STRICT") {
       if (pID === "" || pID === "---") break; // Stop at empty row in strict mode
    }

    // --- MODE: SKIP_EMPTY (Pneumatic) ---
    if (stopMode === "SKIP_EMPTY") {
       if (pID === "" || pID === "---") continue; // Skip empty row
       
       // Safety: Stop if ID is not numeric (prevents reading footer notes)
       // Checks if the first character is NOT a digit 0-9
       if (!/^\d/.test(pID)) {
          break;
       }
    }
    
    // Add valid item
    items.push([pID, desc]);
  }
  return items;
}

function fetchVisionItems(sourceSheet, triggerPhrase, colID_Index, stopPhrases) {
  var lastRow = sourceSheet.getLastRow();
  var rangeValues = sourceSheet.getRange(1, 5, lastRow, 3).getValues();
  var startRowIndex = -1;
  for (var i = 0; i < rangeValues.length; i++) {
    if (rangeValues[i][1].toString().trim().indexOf(triggerPhrase) > -1) {
      startRowIndex = i + 1; break;
    }
  }
  if (startRowIndex === -1) return [];

  var items = [];
  var currentCategory = ""; 
  for (var k = startRowIndex; k < rangeValues.length; k++) {
    var valE = rangeValues[k][0].toString().trim();
    var valF = rangeValues[k][1].toString().trim();
    var valG = rangeValues[k][2].toString().trim();
    if (stopPhrases && stopPhrases.some(function(p) { return valF.indexOf(p) > -1 || valE.indexOf(p) > -1; })) break;
    if (valE !== "" && valE !== "---") currentCategory = valE;
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
    if (rangeValues[i][0].toString().trim().indexOf(triggerPhrase) > -1) {
      startRowIndex = i + 1; break;
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
    if (stopPhrases && stopPhrases.some(function(s) { return pID.indexOf(s) > -1; })) break;
    if (pID.indexOf(":") > -1) break; 
    if (pID !== "" && pID !== "---") items.push([pID, desc]); 
  }
  return items;
}

function updateSection_Core(sourceSheet, destSheet, destHeaderName, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc) {
  var rawItems = fetchRawItems(sourceSheet, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc, []);
  var syncItems = rawItems.map(function(item) { return [item[0], item[1], "1"]; });
  performSurgicalSync(destSheet, [{ destName: destHeaderName, items: syncItems }]);
}

// =========================================
// DROPDOWN SECTION SETUP
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

  if (categoryColIndex != null) destSheet.getRange(headerRow, 2).setValue("CATEGORY");

  var startWriteRow = headerRow + 1;
  var targetItemCount = 10; 

  var currentRow = startWriteRow;
  var foundItemCount = 0;
  var safetyLimit = 0;
  
  while (safetyLimit < 500) { 
    var rowVals = destSheet.getRange(currentRow, 1, 1, 5).getValues()[0]; 
    
    if (rowVals[0].toString() !== "") break; 

    if (rowVals[2].toString() !== "") {
        foundItemCount++;
        
        if (foundItemCount > targetItemCount) {
            destSheet.deleteRows(currentRow, 1);
            foundItemCount--; 
            currentRow--; 
        } else {
            destSheet.getRange(currentRow, 3).setValue(foundItemCount);
            applyRowFormatting(destSheet, currentRow, dropdownRangeString, vlookupRangeString, categoryColIndex);
        }
    } 
    currentRow++;
    safetyLimit++;
  }

  if (foundItemCount < targetItemCount) {
    var itemsToAdd = targetItemCount - foundItemCount;
    destSheet.insertRowsAfter(currentRow - 1, itemsToAdd);
    
    for (var i = 0; i < itemsToAdd; i++) {
        var r = currentRow + i;
        var itemNum = foundItemCount + 1 + i;
        destSheet.getRange(r, 3).setValue(itemNum); 
        applyRowFormatting(destSheet, r, dropdownRangeString, vlookupRangeString, categoryColIndex);
        
        destSheet.getRange(r, 7).insertCheckboxes();
        destSheet.getRange(r, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
    }
  }
}

function applyRowFormatting(destSheet, r, dropdownRangeString, vlookupRangeString, categoryColIndex) {
    var rangeD = destSheet.getRange(r, 4);
    var rangeE = destSheet.getRange(r, 5);
    
    if (rangeD.getDataValidation() == null) {
      rangeD.clearContent();
      rangeE.clearContent(); 
      if (categoryColIndex != null) destSheet.getRange(r, 2).clearContent(); 
    }
    
    var dropdownRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRangeString), true)
        .setAllowInvalid(true).build();
    rangeD.setDataValidation(dropdownRule);

    var cellD_Ref = "D" + r;
    if (rangeE.getFormula() === "") {
        rangeE.setFormula('=IFERROR(VLOOKUP(' + cellD_Ref + ', ' + vlookupRangeString + ', 2, FALSE), "")');
    }
    if (categoryColIndex != null) {
      destSheet.getRange(r, 2).setFormula('=IFERROR(VLOOKUP(' + cellD_Ref + ', ' + vlookupRangeString + ', ' + categoryColIndex + ', FALSE), "")');
    }
}

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
    
    if (sourceItems.length > existingDataCount) destSheet.insertRowsAfter(startWriteRow + existingDataCount - 1, sourceItems.length - existingDataCount);
    else if (sourceItems.length < existingDataCount) destSheet.deleteRows(startWriteRow + sourceItems.length, existingDataCount - sourceItems.length);
    
    if (sourceItems.length > 0) {
      var outputBlock = sourceItems.map(function(item, m) { return [m + 1, item[0], item[1], item[2]]; });
      destSheet.getRange(startWriteRow, 3, sourceItems.length, 4).setValues(outputBlock);
      destSheet.getRange(startWriteRow, 7, sourceItems.length, 1).insertCheckboxes();
      destSheet.getRange(startWriteRow, 9, sourceItems.length, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
    }
  }  
}
