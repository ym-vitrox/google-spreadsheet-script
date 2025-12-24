/**
 * ORDERING LIST SCRIPT
 * * Features:
 * 1. Master Sync from Source BOM (Preserves User Mappings in REF_DATA Cols W-AD).
 * 2. Module Section: Spreadsheet-driven Dependency Logic (REF_DATA Cols C & W,X,Y,Z, AA,AB, AC,AD).
 * - Electrical (Col W/X): Rotational (Based on instance count).
 * - Tooling (Col Y/Z): Stacked (Fixed, multiple items allowed).
 * - Tooling Options (REF_DATA P->S): Single Row Dropdown with Dynamic Formulas.
 * - Jigs (Col AA/AB): Stacked (Fixed, manual mapping, bottom of list).
 * - Vision (Col AC/AD): Triggered by manual mapping. Single (Fixed) or Multiple (Dropdown). New Bottom Layer.
 * 3. Shopping List Logic (Basic Tool & Pneumatic) via onEdit (Config Section).
 * 4. Vision Section: Categorized Dropdowns (REF_DATA Cols AF-AH).
 * 5. Renumbering Tool.
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
  var refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REF_DATA");
  // --- CONSTANTS: RUBBER TIP DEPENDENCY ---
  var RUBBER_TIP_PARENTS = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
  var RUBBER_TIP_SOURCE_ID = "430001-A380";
  
  // 1. Fetch Module Mapping Data (Cols C to AD)
  // Extended range C:AD to capture Jigs (AA, AB) and Vision (AC, AD)
  var refData = refSheet.getRange("C:AD").getValues();
  
  // 2. Fetch Tooling Option Data for Deletion Checking (Cols P to Q)
  var optionData = refSheet.getRange("P:Q").getValues();

  // 3. Fetch Shadow Menu Data for Insertion (Cols U to V)
  var menuData = refSheet.getRange("U:V").getValues();
  
  // --- CLEANUP (Handle Removal or Swap) ---
  if (oldVal) {
    var oldParentConfig = findParentConfig(refData, oldVal);
    if (oldParentConfig) {
      var possibleChildren = [];
      // Gather direct children (Elec + Tool)
      var oldToolIds = [];
      if (oldParentConfig.elecIds) possibleChildren = possibleChildren.concat(oldParentConfig.elecIds.split(';').map(function(s){ return s.trim(); }));
      if (oldParentConfig.toolIds) {
        var tIds = oldParentConfig.toolIds.split(';').map(function(s){ return s.trim(); });
        possibleChildren = possibleChildren.concat(tIds);
        oldToolIds = tIds;
      }
      // Gather Jigs
      if (oldParentConfig.jigIds) {
         possibleChildren = possibleChildren.concat(oldParentConfig.jigIds.split(';').map(function(s){ return s.trim(); }));
      }
      // Gather Vision (For deletion, we add all IDs in the map string)
      if (oldParentConfig.visionIds) {
         possibleChildren = possibleChildren.concat(oldParentConfig.visionIds.split(';').map(function(s){ return s.trim(); }));
      }
      
      // Gather Grandchildren (Standard Tooling Options AND Rubber Tips)
      for (var t = 0; t < oldToolIds.length; t++) {
        var currentToolId = oldToolIds[t];
        // A. Standard Options
        var grandChildren = getToolingOptionIDs(optionData, currentToolId);
        if (grandChildren.length > 0) {
          possibleChildren = possibleChildren.concat(grandChildren);
        }

        // B. Special Rubber Tip Options
        if (RUBBER_TIP_PARENTS.includes(currentToolId)) {
             var rtChildren = getToolingOptionIDs(optionData, RUBBER_TIP_SOURCE_ID);
             if (rtChildren.length > 0) {
                 possibleChildren = possibleChildren.concat(rtChildren);
             }
        }
      }

      // Execute Deletion
      var checkRow = row + 1;
      while (checkRow <= sheet.getMaxRows()) {
        var childPartID = sheet.getRange(checkRow, 4).getValue();
        // We delete if it matches a known child, OR if it's blank but seemingly part of the block (safety for empty dropdowns)
        if (possibleChildren.includes(childPartID) || (childPartID === "" && sheet.getRange(checkRow, 3).getValue() === "")) {
          sheet.deleteRow(checkRow);
        } else {
          break; // Stop at first unrelated row
        }
      }
    }
  }

  // --- INSERTION (Handle New Selection) ---
  if (newVal) {
    var config = findParentConfig(refData, newVal);
    if (!config) return; 

    // 1. Determine Electrical Kit (Rotation)
    var elecToInsert = null;
    if (config.elecIds) {
      var eIds = config.elecIds.split(';').map(function(s){ return s.trim(); });
      var eDescs = config.elecDesc.split(';').map(function(s){ return s.trim(); });
      
      var instanceCount = 0;
      var sectionIds = sheet.getRange(startRow, 4, endRow - startRow + 1, 1).getValues();
      for (var i = 0; i < sectionIds.length; i++) {
        if (sectionIds[i][0] == newVal) instanceCount++;
        if (startRow + i == row) break; 
      }
      
      var index = (instanceCount - 1) % eIds.length;
      if (eIds[index]) {
        elecToInsert = { id: eIds[index], desc: (eDescs[index] || ""), type: 'child' };
      }
    }

    var itemsToAdd = [];
    if (elecToInsert) itemsToAdd.push(elecToInsert);

    // 2. Determine Tooling Kit (Stacking) AND Grandchildren
    if (config.toolIds) {
      var tIds = config.toolIds.split(';').map(function(s){ return s.trim(); });
      var tDescs = config.toolDesc.split(';').map(function(s){ return s.trim(); });
      
      for (var t = 0; t < tIds.length; t++) {
        if (tIds[t]) {
          // Push the Tooling Kit itself
          itemsToAdd.push({ id: tIds[t], desc: (tDescs[t] || ""), type: 'child' });
          // CHECK FOR OPTIONS
          var optionRange = getToolingOptionRange(menuData, tIds[t]);
          if (optionRange) {
             itemsToAdd.push({
               type: 'grandchild',
               parentToolId: tIds[t],
               refDataStart: optionRange.startRow, 
               refDataEnd: optionRange.endRow    
            });
          }

          // CHECK FOR RUBBER TIP
          if (RUBBER_TIP_PARENTS.includes(tIds[t])) {
              var rtRange = getToolingOptionRange(menuData, RUBBER_TIP_SOURCE_ID);
              if (rtRange) {
                  itemsToAdd.push({
                      type: 'rubber_tip',
                      parentToolId: RUBBER_TIP_SOURCE_ID,
                      refDataStart: rtRange.startRow,
                      refDataEnd: rtRange.endRow
                  });
              }
          }
        }
      }
    }
    
    // 3. Determine Jig Items (Stacking, Bottom of List)
    if (config.jigIds) {
      var jIds = config.jigIds.split(';').map(function(s){ return s.trim(); });
      var jDescs = config.jigDesc.split(';').map(function(s){ return s.trim(); });
      
      for (var j = 0; j < jIds.length; j++) {
        if (jIds[j]) {
          itemsToAdd.push({ id: jIds[j], desc: (jDescs[j] || ""), type: 'jig' });
        }
      }
    }

    // 4. Determine Vision Items (New Last Layer)
    if (config.visionIds) {
      var vIds = config.visionIds.split(';').map(function(s){ return s.trim(); });
      // Clean empty strings if any
      vIds = vIds.filter(function(id) { return id.length > 0; });
      
      if (vIds.length === 1) {
        // Option A: Single ID (Fixed)
        itemsToAdd.push({ id: vIds[0], type: 'vision_fixed' });
      } else if (vIds.length > 1) {
        // Option B: Multiple IDs (Dropdown - Strict Subset)
        itemsToAdd.push({ ids: vIds, type: 'vision_select' });
      }
    }

    if (itemsToAdd.length === 0) return;
    
    // Insert block of rows
    sheet.insertRowsAfter(row, itemsToAdd.length);

    var startInsertRow = row + 1;
    for (var k = 0; k < itemsToAdd.length; k++) {
      var currentRow = startInsertRow + k;
      var item = itemsToAdd[k];

      if (item.type === 'child') {
        // Standard Child (Elec or Tool)
        sheet.getRange(currentRow, 4).setValue(item.id);
        sheet.getRange(currentRow, 5).setValue(item.desc);
        sheet.getRange(currentRow, 4).clearDataValidations(); 
      } 
      else if (item.type === 'jig') {
        // Jig Item (Static)
        sheet.getRange(currentRow, 4).setValue(item.id);
        sheet.getRange(currentRow, 5).setValue(item.desc);
        sheet.getRange(currentRow, 4).clearDataValidations();
      }
      else if (item.type === 'grandchild') {
        // Option Row (Dropdown + Formulas)
        var rangeNotation = "REF_DATA!V" + item.refDataStart + ":V" + item.refDataEnd;
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(rangeNotation), true)
          .setAllowInvalid(true).build();
        sheet.getRange(currentRow, 4).setDataValidation(rule);
        
        var formulaB = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!Q:S, 2, FALSE), "")';
        sheet.getRange(currentRow, 2).setFormula(formulaB);

        var formulaE = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!Q:S, 3, FALSE), "")';
        sheet.getRange(currentRow, 5).setFormula(formulaE);
      }
      else if (item.type === 'rubber_tip') {
        // RUBBER TIP ROW
        var rangeNotation = "REF_DATA!V" + item.refDataStart + ":V" + item.refDataEnd;
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getRange(rangeNotation), true)
          .setAllowInvalid(true).build();
        sheet.getRange(currentRow, 4).setDataValidation(rule);
        
        sheet.getRange(currentRow, 2).clearContent(); 
        var formulaE = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!Q:S, 3, FALSE), "")';
        sheet.getRange(currentRow, 5).setFormula(formulaE);
      }
      else if (item.type === 'vision_fixed') {
        // Vision Fixed (Single ID, but use DB for Category/Desc to be safe)
        sheet.getRange(currentRow, 4).setValue(item.id);
        
        // Category Formula (Col B) -> REF_DATA!AH (Index 3 of AF:AH)
        var formulaB = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 3, FALSE), "")';
        sheet.getRange(currentRow, 2).setFormula(formulaB);
        
        // Description Formula (Col E) -> REF_DATA!AG (Index 2 of AF:AH)
        var formulaE = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 2, FALSE), "")';
        sheet.getRange(currentRow, 5).setFormula(formulaE);
        
        sheet.getRange(currentRow, 4).clearDataValidations();
      }
      else if (item.type === 'vision_select') {
        // Vision Select (Dropdown of strict subset)
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(item.ids, true)
          .setAllowInvalid(true).build();
        sheet.getRange(currentRow, 4).setDataValidation(rule);
        
        // Category Formula
        var formulaB = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 3, FALSE), "")';
        sheet.getRange(currentRow, 2).setFormula(formulaB);
        
        // Description Formula
        var formulaE = '=IFERROR(VLOOKUP(D' + currentRow + ', REF_DATA!AF:AH, 2, FALSE), "")';
        sheet.getRange(currentRow, 5).setFormula(formulaE);
      }

      // Formatting for ALL inserted rows
      if (item.type !== 'vision_fixed' && item.type !== 'vision_select') {
        // Standard cleanup for others (Vision sets its own formulas)
        sheet.getRange(currentRow, 3).clearContent();
      }
      
      sheet.getRange(currentRow, 7).insertCheckboxes();
      var releaseRule = SpreadsheetApp.newDataValidation()
          .requireValueInList(['CHARGE OUT', 'MRP'], true).build();
      sheet.getRange(currentRow, 9).setDataValidation(releaseRule);
    }
  }
}

// --- HELPER: Find Parent Config ---
function findParentConfig(refData, parentID) {
  // refData is now C:AD
  // Col C (0) = ID
  // Col W (20) = Elec IDs
  // Col X (21) = Elec Desc
  // Col Y (22) = Tool IDs
  // Col Z (23) = Tool Desc
  // Col AA (24) = Jig IDs
  // Col AB (25) = Jig Desc
  // Col AC (26) = Vision IDs
  // Col AD (27) = Vision Desc
  for (var i = 0; i < refData.length; i++) {
    if (refData[i][0] == parentID) {
      return {
        elecIds: refData[i][20],
        elecDesc: refData[i][21],
        toolIds: refData[i][22],
        toolDesc: refData[i][23],
        jigIds: refData[i][24],
        jigDesc: refData[i][25],
        visionIds: refData[i][26],
        visionDesc: refData[i][27]
      };
    }
  }
  return null;
}

// --- HELPER: Get Option IDs for Deletion (From Database P:Q) ---
function getToolingOptionIDs(optionData, parentToolID) {
  var ids = [];
  for (var i = 0; i < optionData.length; i++) {
    if (optionData[i][0] == parentToolID) {
      if(optionData[i][1]) ids.push(optionData[i][1]);
    }
  }
  return ids;
}

// --- HELPER: Get Option Range for Insertion (From Shadow Table U:V) ---
function getToolingOptionRange(menuData, parentToolID) {
  var startRow = -1;
  var endRow = -1;
  for (var i = 0; i < menuData.length; i++) {
    if (menuData[i][0] == parentToolID) {
      if (startRow === -1) startRow = i + 1;
      endRow = i + 1;
    }
  }
  if (startRow !== -1) return { startRow: startRow, endRow: endRow };
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
  // Updated Range: Fetch C:AD
  var refData = refSheet.getRange("C:AD").getValues();
  
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
      var childRowAbs = startRow + i + 1;
      if (childRowAbs > endRow + 10) continue;
      
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
// 4. MASTER SYNC LOGIC
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
    var refSheet = destSS.getSheetByName("REF_DATA");
    
    // A. Update Reference Data (Preserving User Mappings in W-AD)
    updateReferenceData(sourceSS, sourceSheet);
    // B. Update Shopping Lists (I-L)
    updateShoppingLists(sourceSheet);
    // C. PROCESS TOOLING OPTIONS
    processToolingOptions(sourceSS, refSheet);
    // D. PROCESS VISION DATA (New Logic AF-AH, AJ-AK)
    processVisionData(sourceSheet, refSheet);
    
    // E. Sync Main Sheet Sections
    updateSection_Core(sourceSheet, destSheet, "CORE", "CORE :430000-A557", 3, 4);
    setupDropdownSection(destSheet, "CONFIG", "REF_DATA!A:A", "REF_DATA!A:B", null);
    setupDropdownSection(destSheet, "MODULE", "REF_DATA!C:C", "REF_DATA!C:D", null);
    // F. Setup Vision Section (Categorized)
    setupVisionSection(destSheet);
    
    ui.alert("Sync Complete", "REF_DATA updated.\n- Manual Mappings preserved in W:AD\n- Tooling Options\n- Vision Data (AF:AH)\n- Module & Config Masters updated.", ui.ButtonSet.OK);
  } catch (e) {
    console.error(e);
    ui.alert("Error during Sync", e.message, ui.ButtonSet.OK);
  }
}

// =========================================
// HELPER: TOOLING ILLUSTRATION PARSER
// =========================================
function processToolingOptions(sourceSS, refSheet) {
  var toolingSheet = sourceSS.getSheetByName("Tooling Illustration");
  if (!toolingSheet) {
    console.warn("Tooling Illustration sheet not found.");
    return;
  }
  
  var lastRow = toolingSheet.getLastRow();
  var rawData = toolingSheet.getRange(1, 1, lastRow, 8).getValues();
  var databaseOutput = [];
  var currentParentID = null;
  var currentCategory = null;

  for (var i = 0; i < rawData.length; i++) {
    var colA = String(rawData[i][0]).trim();
    var colB = String(rawData[i][1]).trim();
    var colF = String(rawData[i][5]).trim();
    var colH = String(rawData[i][7]).trim();

    var match = colA.match(/\[(.*?)\]/);
    if (match && match[1]) {
      currentParentID = match[1];
      currentCategory = null;
    }
    if (colB !== "") {
      currentCategory = colB;
    }
    if (currentParentID && colF !== "" && colF !== "Part ID") {
      databaseOutput.push([currentParentID, colF, (currentCategory || ""), colH]);
    }
  }

  refSheet.getRange("P:S").clearContent();
  if (databaseOutput.length > 0) {
    refSheet.getRange(1, 16, databaseOutput.length, 4).setValues(databaseOutput);
  }

  var menuOutput = [];
  if (databaseOutput.length > 0) {
    var grouped = {};
    var orderParents = [];
    
    for (var k = 0; k < databaseOutput.length; k++) {
      var pID = databaseOutput[k][0];
      if (!grouped[pID]) {
        grouped[pID] = [];
        orderParents.push(pID);
      }
      grouped[pID].push({
        partId: databaseOutput[k][1],
        cat: databaseOutput[k][2]
      });
    }

    for (var p = 0; p < orderParents.length; p++) {
      var parent = orderParents[p];
      var items = grouped[parent];
      var lastCat = null;
      
      for (var m = 0; m < items.length; m++) {
        var itm = items[m];
        var thisCat = itm.cat;

        if (thisCat !== "" && thisCat !== lastCat) {
          menuOutput.push([parent, "--- " + thisCat + " (Do Not Click) ---"]);
          lastCat = thisCat;
        }
        menuOutput.push([parent, itm.partId]);
      }
    }
  }

  refSheet.getRange("U:V").clearContent();
  if (menuOutput.length > 0) {
    refSheet.getRange(1, 21, menuOutput.length, 2).setValues(menuOutput);
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

  // 1. BACKUP EXISTING MAPPINGS (Cols W, X, Y, Z, AA, AB, AC, AD)
  // We scan range C:AD. Key is C (Index 0). Data is W:AD (Indices 20-27).
  var lastRefRow = refSheet.getLastRow();
  var existingMappings = {}; 
  
  if (lastRefRow > 0) {
    // Fetch columns 3 (C) through 30 (AD). Total 28 columns.
    var currentData = refSheet.getRange(1, 3, lastRefRow, 28).getValues();
    for (var i = 0; i < currentData.length; i++) {
      var pId = currentData[i][0].toString().trim(); // Column C
      if (pId !== "") {
        existingMappings[pId] = {
          eId: currentData[i][20],   // Col W
          eDesc: currentData[i][21], // Col X
          tId: currentData[i][22],   // Col Y
          tDesc: currentData[i][23], // Col Z
          jId: currentData[i][24],   // Col AA 
          jDesc: currentData[i][25], // Col AB 
          vId: currentData[i][26],   // Col AC (Vision ID)
          vDesc: currentData[i][27]  // Col AD (Vision Desc)
        };
      }
    }
  }
  
  // 2. CLEAR DATA (Split Clearing: A:D and W:AD)
  refSheet.getRange("A:D").clear();
  refSheet.getRange("W:AD").clear(); 

  // 3. FETCH NEW DATA FROM SOURCE
  var configItems = fetchRawItems(sourceSheet, "OPTIONAL MODULE: 430001-A712", 6, 7, ["CONFIGURABLE MODULE"]);
  var moduleItems = fetchRawItems(sourceSheet, "CONFIGURABLE MODULE: 430001-A713", 6, 7, ["CONFIGURABLE VISION MODULE"]);

  // 4. WRITE NEW DATA
  if (configItems.length > 0) refSheet.getRange(1, 1, configItems.length, 2).setValues(configItems);
  
  if (moduleItems.length > 0) {
    var moduleOutput = []; // For C:D
    var mappingOutput = []; // For W:AD

    for (var m = 0; m < moduleItems.length; m++) {
      var mId = moduleItems[m][0].toString().trim();
      var mDesc = moduleItems[m][1];
      
      var eId = "", eDesc = "", tId = "", tDesc = "", jId = "", jDesc = "", vId = "", vDesc = "";
      if (existingMappings[mId]) {
        eId = existingMappings[mId].eId;
        eDesc = existingMappings[mId].eDesc;
        tId = existingMappings[mId].tId;
        tDesc = existingMappings[mId].tDesc;
        jId = existingMappings[mId].jId;
        jDesc = existingMappings[mId].jDesc;
        vId = existingMappings[mId].vId;
        vDesc = existingMappings[mId].vDesc;
      }
      
      moduleOutput.push([mId, mDesc]);
      mappingOutput.push([eId, eDesc, tId, tDesc, jId, jDesc, vId, vDesc]);
    }
    // Write Module Data to C:D
    refSheet.getRange(1, 3, moduleOutput.length, 2).setValues(moduleOutput);
    // Write Mapping Data to W:AD (Col 23 is W)
    refSheet.getRange(1, 23, mappingOutput.length, 8).setValues(mappingOutput);
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

// =========================================
// VISION LOGIC (UPDATED for AF-AH, AJ-AK)
// =========================================
function processVisionData(sourceSheet, refSheet) {
  // 1. Locate "CONFIGURABLE VISION MODULE" Header
  var textFinder = sourceSheet.createTextFinder("CONFIGURABLE VISION MODULE").matchEntireCell(false);
  var found = textFinder.findNext();
  
  if (!found) {
    console.warn("Header 'CONFIGURABLE VISION MODULE' not found in Source.");
    return;
  }
  
  var startRow = found.getRow() + 1;
  var lastRow = sourceSheet.getLastRow();
  if (lastRow < startRow) return;
  
  // Read Columns E (Category), F (ID), G (Desc)
  // E=5, F=6, G=7.
  var numRows = lastRow - startRow + 1;
  var rawData = sourceSheet.getRange(startRow, 5, numRows, 3).getValues();
  
  var databaseOutput = []; // For AF:AH (ID, Desc, Cat)
  var menuOutput = [];     // For AJ:AK (Header/ID) - Optional visual list
  
  var currentCategory = "Uncategorized";
  var groupedData = {};
  var orderCategories = [];
  
  for (var i = 0; i < rawData.length; i++) {
    var cat = rawData[i][0].toString().trim();
    var id = rawData[i][1].toString().trim();
    var desc = rawData[i][2].toString().trim();
    
    // Stop if we hit end of block (e.g., next header or empty block)
    // Adjust logic based on file structure. 
    // Assuming empty ID means end or just blank row.
    if (id === "" && cat === "") continue; 
    
    // Update Category if present (Group Header logic)
    if (cat !== "") {
      currentCategory = cat;
      // If this row also has an ID, use it. If not, it's just a header row.
    }
    
    if (id !== "" && id !== "Part Number") { // Skip sub-header if exists
       databaseOutput.push([id, desc, currentCategory]);
       
       if (!groupedData[currentCategory]) {
         groupedData[currentCategory] = [];
         orderCategories.push(currentCategory);
       }
       groupedData[currentCategory].push(id);
    }
  }

  // 2. Clear Old Data (AF:AH and AJ:AK)
  // AF=32, AG=33, AH=34. (3 columns)
  // AJ=36, AK=37. (2 columns)
  refSheet.getRange(1, 32, refSheet.getMaxRows(), 3).clearContent();
  refSheet.getRange(1, 36, refSheet.getMaxRows(), 2).clearContent();

  // 3. Write Database (AF:AH)
  if (databaseOutput.length > 0) {
    refSheet.getRange(1, 32, databaseOutput.length, 3).setValues(databaseOutput);
  }

  // 4. Write Shadow Menu (AJ:AK) - Even if not used by Option X, we keep it for reference
  if (orderCategories.length > 0) {
    for (var c = 0; c < orderCategories.length; c++) {
      var cName = orderCategories[c];
      menuOutput.push(["--- " + cName + " ---", ""]);
      var ids = groupedData[cName];
      for (var k = 0; k < ids.length; k++) {
        menuOutput.push(["", ids[k]]);
      }
    }
    if (menuOutput.length > 0) {
      refSheet.getRange(1, 36, menuOutput.length, 2).setValues(menuOutput);
    }
  }
}

function setupVisionSection(sheet) {
  var visionFinder = sheet.getRange("A:A").createTextFinder("VISION").matchEntireCell(true).findNext();
  if (!visionFinder) return;
  var visionRow = visionFinder.getRow();
  
  var headerRow = -1;
  var checkRange = sheet.getRange(visionRow, 5, 20, 1).getValues(); 
  for (var r = 0; r < checkRange.length; r++) {
    if (String(checkRange[r][0]).toUpperCase() === "DESCRIPTION") {
      headerRow = visionRow + r;
      break;
    }
  }
  if (headerRow === -1) return;
  
  sheet.getRange(headerRow, 2).setValue("CATEGORY");
  var startRow = headerRow + 1;
  var searchRange = sheet.getRange(startRow, 1, sheet.getMaxRows() - startRow + 1, 1);
  var toolingFinder = searchRange.createTextFinder("TOOLING").matchEntireCell(true).findNext();
  if (!toolingFinder) {
    // If not found, maybe at bottom.
    // Proceed cautiously.
  }
  
  // This function sets up the "VISION SECTION" (Logic 4 from header).
  // It uses REF_DATA!AF:AH (Database) for VLOOKUPs now, not M:O.
  // Previous code used M:O. We must update to AF:AH.
  
  var endRow = toolingFinder ? toolingFinder.getRow() - 1 : sheet.getLastRow();
  if (endRow < startRow) return;

  // Dropdown Rule (From AF column - ID list)
  // Actually, for the Vision Section (Logic 4), it's a Categorized Dropdown.
  // Previous code referenced M:M. Now we should reference AF:AF (IDs).
  // But wait, the Vision Section (Logic 4) is different from Module-Dependent Vision (Logic 2).
  // Logic 4 is the "VISION" section at the bottom of the sheet.
  // The user prompt discussed "Vision Layer" inside Module Block.
  // This function setupVisionSection handles the separate Vision section.
  // I will update it to use the NEW database (AF:AH) for consistency.
  
  const validationRange = SpreadsheetApp.getActive().getSheetByName("REF_DATA").getRange("AF:AF");
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(validationRange)
    .setAllowInvalid(true) 
    .build();
  sheet.getRange(startRow, 4, endRow - startRow + 1).setDataValidation(rule);

  for (let r = startRow; r <= endRow; r++) {
    // Col B: Category (Index 3 of AF:AH)
    let formulaB = `=IFERROR(VLOOKUP(D${r}, REF_DATA!$AF:$AH, 3, FALSE), "")`;
    sheet.getRange(r, 2).setFormula(formulaB);

    // Col E: Description (Index 2 of AF:AH)
    let formulaE = `=IFERROR(VLOOKUP(D${r}, REF_DATA!$AF:$AH, 2, FALSE), "")`;
    sheet.getRange(r, 5).setFormula(formulaE);
  }
}

// =========================================
// SECTION SETUP UTILS
// =========================================
function updateSection_Core(sourceSheet, destSheet, destHeaderName, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc) {
  var rawItems = fetchRawItems(sourceSheet, sourceTriggerPhrase, sourceColIndex_ID, sourceColIndex_Desc, []);
  var syncItems = rawItems.map(function(item) { return [item[0], item[1], "1"]; });
  performSurgicalSync(destSheet, [{ destName: destHeaderName, items: syncItems }]);
}

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
