/**
 * ORDERING LIST SCRIPT
 * Features:
 * 1. Master Sync from Source BOM.
 * 2. Incremental Dependency Logic (Live Kit Insertion) via onEdit.
 * - Logic: "Gap Filling" (Finds first unused Kit ID to prevent duplicates).
 * - Supported Parents: 
 * A. Reject Bin (4 Kits)
 * B. Dynamic Recentering V1 (2 Kits)
 * 3. Smart Row Management (Preserves Kits/Spacers during Sync).
 */

// =========================================
// 1. LIVE TRIGGER (Handle Kit Insertion & Cleanup)
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

  // DEPENDENCY MAP: Parent ID -> List of Child Kits
  var KIT_DEPENDENCIES = {
    // Parent A: Reject Bin
    "430000-A973": [ 
      {id: "430001-A529", desc: "Kit-Misc. Ele. Reject Bin 1"},
      {id: "430001-A530", desc: "Kit-Misc. Ele. Reject Bin 2"},
      {id: "430001-A531", desc: "Kit-Misc. Ele. Reject Bin 3"},
      {id: "430001-A532", desc: "Kit-Misc. Ele. Reject Bin 4"}
    ],
    // Parent B: Dynamic Recentering V1
    "430000-A959": [
      {id: "430000-A989", desc: "Kit-Misc. Ele. Dynamic Recentering V1-#1"},
      {id: "430001-A373", desc: "Kit-Misc. Ele. Dynamic Recentering V1-#2"}
    ]
  };

  // 1. DETERMINE SECTION BOUNDARIES (MODULE ONLY)
  var moduleFinder = sheet.getRange("A:A").createTextFinder("MODULE").matchEntireCell(true).findNext();
  var visionFinder = sheet.getRange("A:A").createTextFinder("VISION").matchEntireCell(true).findNext();
  
  if (!moduleFinder || !visionFinder) return;
  var startRow = moduleFinder.getRow() + 1; // Approx start (header area)
  var endRow = visionFinder.getRow() - 1;   // End before VISION starts

  if (row < startRow || row > endRow) return;

  // Get Values safely
  var newVal = e.value; 
  var oldVal = e.oldValue; 

  // ---------------------------------------------------------
  // STEP A: CLEANUP (Handle Removal or Swap)
  // ---------------------------------------------------------
  // If the previous value was a known Parent, we must check if we need to remove its Child.
  if (oldVal && KIT_DEPENDENCIES[oldVal]) {
    var checkRow = row + 1;
    if (checkRow <= sheet.getMaxRows()) {
      var childPartID = sheet.getRange(checkRow, 4).getValue();
      var possibleChildren = KIT_DEPENDENCIES[oldVal].map(function(k) { return k.id; });
      
      // If the row immediately below contains one of the children of the OLD parent, DELETE IT.
      if (possibleChildren.includes(childPartID)) {
        sheet.deleteRow(checkRow);
        // Note: We don't return here. We proceed to Step B because this might be a Swap.
      }
    }
  }

  // ---------------------------------------------------------
  // STEP B: INSERTION (Handle New Selection)
  // ---------------------------------------------------------
  // If the NEW value is a known Parent, we calculate and insert the correct Child.
  if (newVal && KIT_DEPENDENCIES[newVal]) {
    var parentID = newVal;
    var childList = KIT_DEPENDENCIES[parentID];
    
    // --- GAP FILLING LOGIC (Prevent Duplicates) ---
    // 1. Scan the entire MODULE section to see which Kits are already taken.
    // We grab Column D (Part ID) for the whole section to check neighbors.
    // We grab extra rows to ensure we can check the "child" row of the last item.
    var scanHeight = endRow - startRow + 5; 
    var sectionValues = sheet.getRange(startRow, 4, scanHeight, 1).getValues();
    var usedKitIds = [];

    for (var i = 0; i < sectionValues.length; i++) {
      var scanRowAbs = startRow + i;
      var scanVal = sectionValues[i][0];

      // If we find an instance of the SAME Parent
      // AND it is NOT the row we are currently editing (don't count ourselves)
      if (scanVal == parentID && scanRowAbs !== row) {
        // Check the row immediately below it for a valid kit
        if (i + 1 < sectionValues.length) {
          var potentialChildId = sectionValues[i + 1][0];
          var isKnownChild = childList.some(function(k) { return k.id === potentialChildId; });
          
          if (isKnownChild) {
            usedKitIds.push(potentialChildId);
          }
        }
      }
    }

    // 2. Find the first Kit in the definition list that is NOT in 'usedKitIds'
    var targetKit = null;
    for (var k = 0; k < childList.length; k++) {
      if (usedKitIds.indexOf(childList[k].id) === -1) {
        targetKit = childList[k];
        break; // Found the first available one (e.g., Kit 1)
      }
    }

    // Fallback: If all are used (e.g. 5th Reject Bin added), default to the last one
    if (!targetKit) {
      targetKit = childList[childList.length - 1];
    }
    // ----------------------------------------------

    // Check the row immediately below (again, because Step A might have shifted rows)
    var checkRow = row + 1;
    var rowBelowData = sheet.getRange(checkRow, 3, 1, 3).getValues()[0]; // Col C, D, E
    var valBelowD = rowBelowData[1]; // Part ID

    // If the row below is ALREADY the correct kit, do nothing.
    if (valBelowD === targetKit.id) return;

    // Otherwise, INSERT a new row
    sheet.insertRowAfter(row);
    
    // Populate the new row
    var kitPartIdCell = sheet.getRange(checkRow, 4);
    var kitDescCell = sheet.getRange(checkRow, 5);

    // 1. Set Values
    kitPartIdCell.setValue(targetKit.id);   // Col D (Part ID)
    kitDescCell.setValue(targetKit.desc);   // Col E (Desc)
    
    // 2. CRITICAL FIX: Remove Dropdown from the Kit Row
    // insertRowAfter copies validation from the parent. We must remove it 
    // for the Kit to ensure it is treated as static text.
    kitPartIdCell.clearDataValidations(); 
    
    // 3. Clear Col C explicitly to ensure it's treated as a Kit row (protected from Sync)
    sheet.getRange(checkRow, 3).clearContent(); 
    
    // 4. Add Checkbox & formatting
    sheet.getRange(checkRow, 7).insertCheckboxes(); 
    var releaseRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['CHARGE OUT', 'MRP'], true).build();
    sheet.getRange(checkRow, 9).setDataValidation(releaseRule);
  }
}

// =========================================
// 2. STANDARD MENUS & SYNC
// =========================================
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
    
    // A. CORE (Surgical Sync)
    updateSection_Core(sourceSheet, destSheet, "CORE", "CORE :430000-A557", 3, 4); 

    // B. CONFIG (Dropdown)
    setupDropdownSection(destSheet, "CONFIG", "REF_DATA!A:A", "REF_DATA!A:B", null);

    // C. MODULE (Dropdown - Uses Smart Item Counting to preserve Kits)
    setupDropdownSection(destSheet, "MODULE", "REF_DATA!C:C", "REF_DATA!C:D", null);

    // D. VISION (Dropdown + Category)
    setupDropdownSection(destSheet, "VISION", "REF_DATA!E:E", "REF_DATA!E:G", 3);
    
    ui.alert("Sync Complete", "Lists updated. Kits and Spacers preserved.", ui.ButtonSet.OK);

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

  // 3. FETCH VISION DATA
  var visionItems = fetchVisionItems(sourceSheet, "CONFIGURABLE VISION MODULE", 6, ["CALIBRATION JIG"]);

  // 4. WRITE TO REFERENCE SHEET
  if (configItems.length > 0) refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems);
  if (moduleItems.length > 0) refSheet.getRange(1, 3, moduleItems.length, 2).setValues(moduleItems);
  if (visionItems.length > 0) refSheet.getRange(1, 5, visionItems.length, 3).setValues(visionItems);
}

// =========================================
// FETCHERS
// =========================================
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
// 3. DROPDOWN SECTION SETUP (Modified Logic)
// =========================================
function setupDropdownSection(destSheet, sectionName, dropdownRangeString, vlookupRangeString, categoryColIndex) {
  var textFinder = destSheet.getRange("A:A").createTextFinder(sectionName).matchEntireCell(true);
  var foundParams = textFinder.findAll();
  if (foundParams.length === 0) return;
  var sectionStartRow = foundParams[0].getRow();
  
  // Find Anchor (DESCRIPTION Header)
  var headerRow = -1;
  var checkRange = destSheet.getRange(sectionStartRow, 5, 20, 1).getValues(); 
  for (var r = 0; r < checkRange.length; r++) {
    if (checkRange[r][0].toString().toUpperCase() === "DESCRIPTION") {
      headerRow = sectionStartRow + r;
      break;
    }
  }
  if (headerRow === -1) return;

  // RENAME HEADER (VISION ONLY)
  if (categoryColIndex != null) destSheet.getRange(headerRow, 2).setValue("CATEGORY");

  var startWriteRow = headerRow + 1;
  var targetItemCount = 10; 

  // -----------------------------------------------------------------------
  // SMART ROW COUNTING (Count "Items" only, ignore Spacer/Kit rows)
  // -----------------------------------------------------------------------
  var currentRow = startWriteRow;
  var foundItemCount = 0;
  var safetyLimit = 0;
  
  // Loop until we hit the next Section or Safety Limit
  while (safetyLimit < 500) { 
    var rowVals = destSheet.getRange(currentRow, 1, 1, 5).getValues()[0]; // Cols A, B, C, D, E
    
    // 1. Boundary Check: If Col A has text, we hit the next section.
    if (rowVals[0].toString() !== "") break; 

    // 2. Item Check: If Col C (Index 2) has a number, it's a MAIN ITEM row.
    if (rowVals[2].toString() !== "") {
        foundItemCount++;
        
        // If we found more than 10 numbered items, DELETE this extra item row.
        if (foundItemCount > targetItemCount) {
            destSheet.deleteRows(currentRow, 1);
            foundItemCount--; // Decrement since we deleted
            currentRow--; // Move cursor back
        } else {
            // It's a valid item row. Renumber it to ensure sequence (1, 2, 3...)
            destSheet.getRange(currentRow, 3).setValue(foundItemCount);
            
            // APPLY FORMATTING TO MAIN ITEMS ONLY
            applyRowFormatting(destSheet, currentRow, dropdownRangeString, vlookupRangeString, categoryColIndex);
        }
    } 
    // 3. Kit/Spacer Check: Col C is empty. 
    // We Do NOTHING to these rows. We just skip them.
    // This preserves your Spacer rows and your Kit rows.
    
    currentRow++;
    safetyLimit++;
  }

  // -----------------------------------------------------------------------
  // FILL MISSING ITEMS
  // -----------------------------------------------------------------------
  // If we have fewer than 10 numbered items, insert new ones at the END of the section.
  if (foundItemCount < targetItemCount) {
    var itemsToAdd = targetItemCount - foundItemCount;
    // Insert after the last scanned row (which is effectively the end of the section)
    destSheet.insertRowsAfter(currentRow - 1, itemsToAdd);
    
    for (var i = 0; i < itemsToAdd; i++) {
        var r = currentRow + i;
        var itemNum = foundItemCount + 1 + i;
        destSheet.getRange(r, 3).setValue(itemNum); // Set Item Number
        applyRowFormatting(destSheet, r, dropdownRangeString, vlookupRangeString, categoryColIndex);
        
        // Add Checkboxes/Release defaults for new rows
        destSheet.getRange(r, 7).insertCheckboxes();
        destSheet.getRange(r, 9).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['CHARGE OUT', 'MRP'], true).build());
    }
  }
}

// Helper to apply dropdowns and formulas to a single row
function applyRowFormatting(destSheet, r, dropdownRangeString, vlookupRangeString, categoryColIndex) {
    var rangeD = destSheet.getRange(r, 4);
    var rangeE = destSheet.getRange(r, 5);
    
    // If no validation, clear content (assumed old text), but only if it's being managed as an item
    if (rangeD.getDataValidation() == null) {
      rangeD.clearContent();
      rangeE.clearContent(); 
      if (categoryColIndex != null) destSheet.getRange(r, 2).clearContent(); 
    }
    
    // Set Dropdown
    var dropdownRule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRangeString), true)
        .setAllowInvalid(true).build();
    rangeD.setDataValidation(dropdownRule);

    // Set Formulas
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
