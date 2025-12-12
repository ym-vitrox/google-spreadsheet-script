/**
 * ORDERING LIST SCRIPT
 * * Features:
 * 1. Master Sync from Source BOM (Preserves User Mappings in REF_DATA).
 * 2. Module Section: Spreadsheet-driven Dependency Logic (REF_DATA Cols C->E,F,G,H).
 * - Electrical (Col E/F): Rotational (Based on instance count).
 * - Tooling (Col G/H): Stacked (Fixed, multiple items allowed).
 * 3. Shopping List Logic (Basic Tool & Pneumatic) via onEdit (Config Section).
 * 4. Renumbering Tool.
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
    handleConfigSection(sheet, row, newVal, oldVal);
  }

  // =========================================================
  // LOGIC B: MODULE SECTION (Spreadsheet Dependencies)
  // =========================================================
  if (row >= moduleStart && row <= moduleEnd && moduleStart > 0) {
    handleModuleSection(sheet, row, newVal, oldVal, moduleStart, moduleEnd);
  }
}

// =========================================
// LOGIC HANDLERS
// =========================================

function handleConfigSection(sheet, row, newVal, oldVal) {
  var BASIC_TOOL_TRIGGER = "430001-A378";
  var PNEUMATIC_TRIGGER = "430001-A714";

  // --- 1. DELETE LOGIC ---
  if (oldVal === BASIC_TOOL_TRIGGER) {
    if (row + 10 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 10);
  }
  if (oldVal === PNEUMATIC_TRIGGER) {
    if (row + 3 <= sheet.getMaxRows()) sheet.deleteRows(row + 1, 3);
  }

  // --- 2. INSERT LOGIC ---
  if (newVal === BASIC_TOOL_TRIGGER) {
    insertShoppingList(sheet, row, 10, "REF_DATA!I:I", "REF_DATA!I:J");
  }
  if (newVal === PNEUMATIC_TRIGGER) {
    insertShoppingList(sheet, row, 3, "REF_DATA!K:K", "REF_DATA!K:L");
  }
}

function handleModuleSection(sheet, row, newVal, oldVal, startRow, endRow) {
  // 1. Fetch Mapping Data from REF_DATA (Cols C to H)
  // [ParentID, Desc, ElecIDs, ElecDescs, ToolIDs, ToolDescs]
  var refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  var refData = refSheet.getRange("C:H").getValues(); 

  // --- CLEANUP (Handle Removal or Swap) ---
  if (oldVal) {
    var oldParentConfig = findParentConfig(refData, oldVal);
    // If the old parent had dependencies, we try to find and remove them
    if (oldParentConfig) {
      var possibleChildren = [];
      
      // Gather all possible children (Elec + Tool) from the config
      if (oldParentConfig.elecIds) possibleChildren = possibleChildren.concat(oldParentConfig.elecIds.split(';').map(function(s){ return s.trim(); }));
      if (oldParentConfig.toolIds) possibleChildren = possibleChildren.concat(oldParentConfig.toolIds.split(';').map(function(s){ return s.trim(); }));
      
      // Look at the row immediately below
      var checkRow = row + 1;
      // Loop while the row below contains one of the possible children
      while (checkRow <= sheet.getMaxRows()) {
        var childPartID = sheet.getRange(checkRow, 4).getValue();
        if (possibleChildren.includes(childPartID)) {
          sheet.deleteRow(checkRow);
          // Don't increment checkRow, because the next row shifted up
        } else {
          break; // Stop if we hit a non-child
        }
      }
    }
  }

  // --- INSERTION (Handle New Selection) ---
  if (newVal) {
    var config = findParentConfig(refData, newVal);
    if (!config) return; // No mapping found for this parent

    // 1. Determine Electrical Kit (Rotation)
    var elecToInsert = null;
    if (config.elecIds) {
      var eIds = config.elecIds.split(';').map(function(s){ return s.trim(); });
      var eDescs = config.elecDesc.split(';').map(function(s){ return s.trim(); });
      
      // Calculate Instance Count
      var instanceCount = 0;
      var sectionIds = sheet.getRange(startRow, 4, endRow - startRow + 1, 1).getValues();
      for (var i = 0; i < sectionIds.length; i++) {
        if (sectionIds[i][0] == newVal) instanceCount++;
        if (startRow + i == row) break; // Stop counting at current row
      }
      
      // Rotation Logic: (Count - 1) % Length
      var index = (instanceCount - 1) % eIds.length;
      if (eIds[index]) {
        elecToInsert = { id: eIds[index], desc: (eDescs[index] || "") };
      }
    }

    // 2. Determine Tooling Kit (Stacking)
    var toolsToInsert = [];
    if (config.toolIds) {
      var tIds = config.toolIds.split(';').map(function(s){ return s.trim(); });
      var tDescs = config.toolDesc.split(';').map(function(s){ return s.trim(); });
      
      for (var t = 0; t < tIds.length; t++) {
        if (tIds[t]) {
          toolsToInsert.push({ id: tIds[t], desc: (tDescs[t] || "") });
        }
      }
    }

    // 3. Daisy Chain Insertion
    // Order: Electrical First, then Tooling
    var itemsToAdd = [];
    if (elecToInsert) itemsToAdd.push(elecToInsert);
    if (toolsToInsert.length > 0) itemsToAdd = itemsToAdd.concat(toolsToInsert);

    if (itemsToAdd.length === 0) return;

    // Insert block of rows
    sheet.insertRowsAfter(row, itemsToAdd.length);

    var startInsertRow = row + 1;
    for (var k = 0; k < itemsToAdd.length; k++) {
      var currentRow = startInsertRow + k;
      var item = itemsToAdd[k];

      // Set ID and Direct Description (No VLOOKUP)
      sheet.getRange(currentRow, 4).setValue(item.id);
      sheet.getRange(currentRow, 5).setValue(item.desc);
      
      // Formatting
      sheet.getRange(currentRow, 3).clearContent(); // Clear Item Num
      sheet.getRange(currentRow, 7).insertCheckboxes();
      sheet.getRange(currentRow, 4).clearDataValidations(); // Remove dropdown from child
      
      var releaseRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['CHARGE OUT', 'MRP'], true).build();
      sheet.getRange(currentRow, 9).setDataValidation(releaseRule);
    }
  }
}

// Helper to parse REF_DATA row
function findParentConfig(refData, parentID) {
  for (var i = 0; i < refData.length; i++) {
    // refData structure: [Col C, Col D, Col E, Col F, Col G, Col H]
    // Index 0 = ParentID
    if (refData[i][0] == parentID) {
      return {
        elecIds: refData[i][2],
        elecDesc: refData[i][3],
        toolIds: refData[i][4],
        toolDesc: refData[i][5]
      };
    }
  }
  return null;
}

// =========================================
// HELPER: SHOPPING LIST INSERTION
// =========================================
function insertShoppingList(sheet, row, count, dropdownRef, vlookupRef) {
  sheet.insertRowsAfter(row, count);
  var startInsertRow = row + 1;
  
  var dropDownRange = sheet.getRange(startInsertRow, 4, count, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(dropdownRef), true)
    .setAllowInvalid(true).build();
  dropDownRange.setDataValidation(rule);

  var descRange = sheet.getRange(startInsertRow, 5, count, 1);
  var formulas = [];
  for (var i = 0; i < count; i++) {
    var r = startInsertRow + i;
    formulas.push(['=IFERROR(VLOOKUP(D' + r + ', ' + vlookupRef + ', 2, FALSE), "")']);
  }
  descRange.setFormulas(formulas);
  
  sheet.getRange(startInsertRow, 7, count, 1).insertCheckboxes();
  var releaseRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['CHARGE OUT', 'MRP'], true).build();
  sheet.getRange(startInsertRow, 9, count, 1).setDataValidation(releaseRule);
  
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
// 3. RENUMBERING FUNCTION
// =========================================
function renumberKits() {
  // Note: Renumbering logic relies on REF_DATA now, similar to onEdit.
  // For simplicity, we keep the old structure but point it to REF_DATA.
  // This ensures checking consistency.
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ORDERING LIST");
  if (!sheet) return;
  
  var moduleFinder = sheet.getRange("A:A").createTextFinder("MODULE").matchEntireCell(true).findNext();
  var visionFinder = sheet.getRange("A:A").createTextFinder("VISION").matchEntireCell(true).findNext();
  
  if (!moduleFinder || !visionFinder) return;
  var startRow = moduleFinder.getRow() + 1;
  var endRow = visionFinder.getRow() - 1;

  var range = sheet.getRange(startRow, 4, endRow - startRow + 1, 1);
  var values = range.getValues();
  
  var refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  var refData = refSheet.getRange("C:H").getValues();
  
  var parentCounts = {};

  for (var i = 0; i < values.length; i++) {
    var parentID = values[i][0];
    var config = findParentConfig(refData, parentID);
    
    if (config && config.elecIds) {
      if (!parentCounts[parentID]) parentCounts[parentID] = 0;
      parentCounts[parentID]++;
      
      var eIds = config.elecIds.split(';').map(function(s){ return s.trim(); });
      var eDescs = config.elecDesc.split(';').map(function(s){ return s.trim(); });
      
      var count = parentCounts[parentID];
      var index = (count - 1) % eIds.length;
      var targetId = eIds[index];
      var targetDesc = eDescs[index] || "";

      // Check row below
      var childRowAbs = startRow + i + 1;
      if (childRowAbs > endRow + 10) continue;
      
      // Simple check: Is the row below holding *one* of the possible rotational items?
      var actualChildID = sheet.getRange(childRowAbs, 4).getValue();
      if (eIds.includes(actualChildID)) {
        if (actualChildID !== targetId) {
           sheet.getRange(childRowAbs, 4).setValue(targetId);
           sheet.getRange(childRowAbs, 5).setValue(targetDesc);
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert("Renumbering Complete", "Sequential rotation updated.", SpreadsheetApp.getUi().ButtonSet.OK);
}

// =========================================
// SHARED DATA (Deprecated - Now uses REF_DATA)
// =========================================
function getKitDependencies() {
  return {}; // Logic moved to REF_DATA
}

// =========================================
// 4. MASTER SYNC LOGIC (PRESERVES USER MAPPING)
// =========================================
function runMasterSync() {
  var ui = SpreadsheetApp.getUi();
  
  try {
    var sourceSpreadsheetId = "1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA"; 
    var sourceTabName = "BOM Structure Tree Diagram";
    
    var sourceSS;
    try {
      sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
    } catch (e) {
      throw new Error("Could not open Source Spreadsheet. Check ID. " + e.message);
    }
    var sourceSheet = sourceSS.getSheetByName(sourceTabName);
    if (!sourceSheet) throw new Error("Source tab '" + sourceTabName + "' not found.");
    
    var destSS = SpreadsheetApp.getActiveSpreadsheet();
    var destSheet = destSS.getSheetByName("ORDERING LIST");

    // A. Update Reference Data (Preserving User Mappings in E-H)
    updateReferenceData(sourceSS, sourceSheet);

    // B. Update Shopping Lists (I-L)
    updateShoppingLists(sourceSheet);

    // C. Sync Main Sheet Sections
    updateSection_Core(sourceSheet, destSheet, "CORE", "CORE :430000-A557", 3, 4); 
    setupDropdownSection(destSheet, "CONFIG", "REF_DATA!A:A", "REF_DATA!A:B", null);
    setupDropdownSection(destSheet, "MODULE", "REF_DATA!C:C", "REF_DATA!C:D", null);
    // Move Vision Sync to Q:R to avoid overwriting your new E-H section
    setupDropdownSection(destSheet, "VISION", "REF_DATA!Q:Q", "REF_DATA!Q:R", 3);

    ui.alert("Sync Complete", "REF_DATA updated. User mappings in Columns E-H were preserved.", ui.ButtonSet.OK);

  } catch (e) {
    console.error(e);
    ui.alert("Error during Sync", e.message, ui.ButtonSet.OK);
  }
}

// =========================================
// HELPER: REFERENCE DATA MANAGER (SMART SYNC)
// =========================================
function updateReferenceData(sourceSS, sourceSheet) {
  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var refSheetName = "REF_DATA";
  var refSheet = destSS.getSheetByName(refSheetName);
  
  if (!refSheet) {
    refSheet = destSS.insertSheet(refSheetName);
    refSheet.hideSheet();
  }

  // 1. BACKUP EXISTING MAPPINGS (Cols C -> E, F, G, H)
  // We read the existing mappings so we don't lose them when we refresh the Module List
  var lastRefRow = refSheet.getLastRow();
  var existingMappings = {}; // Map<ParentID, {eId, eDesc, tId, tDesc}>
  
  if (lastRefRow > 0) {
    var currentData = refSheet.getRange(1, 3, lastRefRow, 6).getValues(); // Cols C to H
    for (var i = 0; i < currentData.length; i++) {
      var pId = currentData[i][0].toString().trim();
      if (pId !== "") {
        existingMappings[pId] = {
          eId: currentData[i][2], // Col E
          eDesc: currentData[i][3], // Col F
          tId: currentData[i][4], // Col G
          tDesc: currentData[i][5]  // Col H
        };
      }
    }
  }
  
  // 2. CLEAR DATA (But we have the backup)
  // We clear A-D (Config & Module Masters) and Q-S (Vision). We rewrite E-H based on backup.
  refSheet.getRange("A:D").clear();
  refSheet.getRange("Q:S").clear();
  refSheet.getRange("E:H").clear(); // Clear mapping area to be rewritten cleanly

  // 3. FETCH NEW DATA FROM SOURCE
  var configItems = fetchRawItems(sourceSheet, "OPTIONAL MODULE: 430001-A712", 6, 7, ["CONFIGURABLE MODULE"]);
  var moduleItems = fetchRawItems(sourceSheet, "CONFIGURABLE MODULE: 430001-A713", 6, 7, ["CONFIGURABLE VISION MODULE"]);
  var visionItems = fetchVisionItems(sourceSheet, "CONFIGURABLE VISION MODULE", 6, ["CALIBRATION JIG"]);

  // 4. WRITE NEW DATA
  // Col A-B: Config Items
  if (configItems.length > 0) refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems);
  
  // Col Q-S: Vision Items (Moved from E)
  if (visionItems.length > 0) refSheet.getRange(1, 17, visionItems.length, 3).setValues(visionItems);

  // Col C-D: Module Items (The Phonebook)
  // AND RESTORE MAPPINGS to E-H
  if (moduleItems.length > 0) {
    var moduleOutput = [];
    
    for (var m = 0; m < moduleItems.length; m++) {
      var mId = moduleItems[m][0].toString().trim();
      var mDesc = moduleItems[m][1];
      
      // Retrieve backup mapping if it exists
      var eId = "", eDesc = "", tId = "", tDesc = "";
      if (existingMappings[mId]) {
        eId = existingMappings[mId].eId;
        eDesc = existingMappings[mId].eDesc;
        tId = existingMappings[mId].tId;
        tDesc = existingMappings[mId].tDesc;
      }
      
      // Build row: [ID, Desc, E_ID, E_Desc, T_ID, T_Desc]
      moduleOutput.push([mId, mDesc, eId, eDesc, tId, tDesc]);
    }
    
    // Write C-H in one go
    refSheet.getRange(1, 3, moduleOutput.length, 6).setValues(moduleOutput);
  }
}

function updateShoppingLists(sourceSheet) {
  var destSS = SpreadsheetApp.getActiveSpreadsheet();
  var refSheet = destSS.getSheetByName("REF_DATA");
  refSheet.getRange("I:L").clear(); // Clear Shopping List Area

  var basicItems = fetchShoppingListItems(sourceSheet, "List-Optional Basic Tool Module: 430001-A378", 12, 13, "STRICT");
  if (basicItems.length > 0) refSheet.getRange(1, 9, basicItems.length, 2).setValues(basicItems);

  var pneumaticItems = fetchShoppingListItems(sourceSheet, "List-Optional Pneumatic Module : 430001-A714", 12, 13, "SKIP_EMPTY");
  if (pneumaticItems.length > 0) refSheet.getRange(1, 11, pneumaticItems.length, 2).setValues(pneumaticItems);
}

// =========================================
// FETCHERS
// =========================================
function fetchShoppingListItems(sourceSheet, triggerPhrase, colID, colDesc, stopMode) {
  var lastRow = sourceSheet.getLastRow();
  var idColumnVals = sourceSheet.getRange(1, colID, lastRow, 1).getValues();
  
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
    
    if (pID.toUpperCase().indexOf("LIST-") === 0) break;
    
    if (stopMode === "STRICT") {
       if (pID === "" || pID === "---") break;
    }
    if (stopMode === "SKIP_EMPTY") {
       if (pID === "" || pID === "---") continue;
       if (!/^\d/.test(pID)) break;
    }
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
      startRowIndex = i + 1;
      break;
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
// SECTION SETUP UTILS
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
