/**
 * ViTrox BOM Configurator - API Layer
 * Provides frontend-facing API functions for the web app
 * 
 * PHASE 1: Foundation - Read Operations
 */

// =========================================
// CONFIGURATION CONSTANTS
// =========================================

const CONFIG = {
  MAIN_SPREADSHEET_ID: '1a4fDQd1U7E650gdxzEIIoQIDv2n38gP735BEbB86mT8',
  SOURCE_BOM_ID: '1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA',
  SHEETS: {
    ORDERING_LIST: 'ORDERING LIST',
    REF_DATA: 'REF_DATA'
  },
  SPECIAL_TRIGGERS: {
    '430001-A378': {
      name: 'Basic Tool Kit',
      shoppingListSize: 10,
      shoppingListColumns: { ids: 9, descriptions: 10 } // Columns I, J (1-indexed)
    },
    '430001-A714': {
      name: 'Pneumatic Kit',
      shoppingListSize: 3,
      shoppingListColumns: { ids: 11, descriptions: 12 } // Columns K, L (1-indexed)
    }
  },
  RUBBER_TIP_PARENTS: ['430001-A689', '430001-A690', '430001-A691', '430001-A692'],
  RUBBER_TIP_SOURCE_ID: '430001-A380'
};

// =========================================
// WEB APP ENTRY POINT
// =========================================

/**
 * Serves the web app HTML
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ViTrox BOM Configurator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// =========================================
// PHASE 1: READ OPERATIONS
// =========================================

/**
 * Get complete application state on startup
 * Loads reference data from REF_DATA and current state from ORDERING LIST
 * 
 * @returns {Object} { success, data: { refData, coreItems, configItems, modules, visionItems }, error? }
 */
function getFullState() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!orderSheet || !refSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Load reference data
    const refData = loadRefData(refSheet);
    
    // Load current state from ORDERING LIST
    const currentState = loadCurrentState(orderSheet, refData);
    
    return {
      success: true,
      data: {
        refData: refData,
        coreItems: currentState.coreItems,
        configItems: currentState.configItems,
        modules: currentState.modules,
        visionItems: currentState.visionItems
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('getFullState error:', error);
    return {
      success: false,
      error: {
        code: 'LOAD_ERROR',
        message: error.message,
        details: error.stack
      }
    };
  }
}

/**
 * Get reference data only (for refresh without reloading current state)
 * 
 * @returns {Object} { success, data: { configItems, moduleItems, visionItems, toolingOptions, shoppingLists }, error? }
 */
function getRefData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!refSheet) {
      throw new Error('REF_DATA sheet not found');
    }
    
    const refData = loadRefData(refSheet);
    
    return {
      success: true,
      data: refData,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('getRefData error:', error);
    return {
      success: false,
      error: {
        code: 'LOAD_ERROR',
        message: error.message
      }
    };
  }
}

// =========================================
// DATA LOADING HELPERS
// =========================================

/**
 * Load all reference data from REF_DATA sheet
 */
function loadRefData(refSheet) {
  const lastRow = refSheet.getLastRow();
  if (lastRow < 1) {
    return {
      configItems: [],
      moduleItems: [],
      visionItems: [],
      toolingOptions: {},
      shoppingLists: { basicTool: [], pneumatic: [] }
    };
  }
  
  // Load Config Items (A:B)
  const configItems = loadConfigItems(refSheet);
  
  // Load Module Items with mappings (C:D + W:AD)
  const moduleItems = loadModuleItems(refSheet);
  
  // Load Vision Items (AF:AH)
  const visionItems = loadVisionItems(refSheet);
  
  // Load Tooling Options (P:S + U:V)
  const toolingOptions = loadToolingOptions(refSheet);
  
  // Load Shopping Lists (I:J, K:L)
  const shoppingLists = loadShoppingLists(refSheet);
  
  return {
    configItems,
    moduleItems,
    visionItems,
    toolingOptions,
    shoppingLists
  };
}

/**
 * Load config items from REF_DATA!A:B
 */
function loadConfigItems(refSheet) {
  const lastRow = refSheet.getLastRow();
  if (lastRow < 1) return [];
  
  const data = refSheet.getRange(1, 1, lastRow, 2).getValues();
  const items = [];
  
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    const desc = String(data[i][1]).trim();
    
    if (id && id !== '' && !id.startsWith('---')) {
      items.push({
        id: id,
        description: desc
      });
    }
  }
  
  return items;
}

/**
 * Load module items with mappings from REF_DATA!C:D and W:AD
 */
function loadModuleItems(refSheet) {
  const lastRow = refSheet.getLastRow();
  if (lastRow < 1) return [];
  
  // Load C:D (ID, Description)
  const basicData = refSheet.getRange(1, 3, lastRow, 2).getValues();
  
  // Load W:AD (Electrical, Tooling, Jig, Vision mappings)
  // W=23, X=24, Y=25, Z=26, AA=27, AB=28, AC=29, AD=30
  const mappingData = refSheet.getRange(1, 23, lastRow, 8).getValues();
  
  const items = [];
  
  for (let i = 0; i < basicData.length; i++) {
    const id = String(basicData[i][0]).trim();
    const desc = String(basicData[i][1]).trim();
    
    if (id && id !== '' && !id.startsWith('---')) {
      items.push({
        id: id,
        description: desc,
        mappings: {
          elecIds: String(mappingData[i][0] || '').trim(),
          elecDesc: String(mappingData[i][1] || '').trim(),
          toolIds: String(mappingData[i][2] || '').trim(),
          toolDesc: String(mappingData[i][3] || '').trim(),
          jigIds: String(mappingData[i][4] || '').trim(),
          jigDesc: String(mappingData[i][5] || '').trim(),
          visionIds: String(mappingData[i][6] || '').trim(),
          visionDesc: String(mappingData[i][7] || '').trim()
        }
      });
    }
  }
  
  return items;
}

/**
 * Load vision items from REF_DATA!AF:AH
 * AF=32, AG=33, AH=34
 */
function loadVisionItems(refSheet) {
  const lastRow = refSheet.getLastRow();
  if (lastRow < 1) return [];
  
  const data = refSheet.getRange(1, 32, lastRow, 3).getValues();
  const items = [];
  
  for (let i = 0; i < data.length; i++) {
    const id = String(data[i][0]).trim();
    const desc = String(data[i][1]).trim();
    const category = String(data[i][2]).trim();
    
    if (id && id !== '' && !id.startsWith('---') && id !== 'Part ID') {
      items.push({
        id: id,
        description: desc,
        category: category || 'Uncategorized'
      });
    }
  }
  
  return items;
}

/**
 * Load tooling options from REF_DATA!P:S and shadow menu U:V
 * P=16 (Parent ID), Q=17 (Child ID), R=18 (Category), S=19 (Description)
 */
function loadToolingOptions(refSheet) {
  const lastRow = refSheet.getLastRow();
  if (lastRow < 1) return {};
  
  const data = refSheet.getRange(1, 16, lastRow, 4).getValues();
  const options = {};
  
  for (let i = 0; i < data.length; i++) {
    const parentId = String(data[i][0]).trim();
    const childId = String(data[i][1]).trim();
    const category = String(data[i][2]).trim();
    const desc = String(data[i][3]).trim();
    
    if (parentId && childId && !childId.startsWith('---')) {
      if (!options[parentId]) {
        options[parentId] = [];
      }
      options[parentId].push({
        id: childId,
        description: desc,
        category: category
      });
    }
  }
  
  return options;
}

/**
 * Load shopping lists from REF_DATA!I:J and K:L
 * I=9, J=10 (Basic Tool), K=11, L=12 (Pneumatic)
 */
function loadShoppingLists(refSheet) {
  const lastRow = refSheet.getLastRow();
  if (lastRow < 1) {
    return { basicTool: [], pneumatic: [] };
  }
  
  // Load Basic Tool (I:J)
  const basicData = refSheet.getRange(1, 9, lastRow, 2).getValues();
  const basicTool = [];
  for (let i = 0; i < basicData.length; i++) {
    const id = String(basicData[i][0]).trim();
    const desc = String(basicData[i][1]).trim();
    if (id && id !== '' && !id.startsWith('---')) {
      basicTool.push({ id, description: desc });
    }
  }
  
  // Load Pneumatic (K:L)
  const pneumaticData = refSheet.getRange(1, 11, lastRow, 2).getValues();
  const pneumatic = [];
  for (let i = 0; i < pneumaticData.length; i++) {
    const id = String(pneumaticData[i][0]).trim();
    const desc = String(pneumaticData[i][1]).trim();
    if (id && id !== '' && !id.startsWith('---')) {
      pneumatic.push({ id, description: desc });
    }
  }
  
  return { basicTool, pneumatic };
}

// =========================================
// CURRENT STATE LOADING
// =========================================

/**
 * Load current state from ORDERING LIST sheet
 */
function loadCurrentState(orderSheet, refData) {
  const lastRow = orderSheet.getLastRow();
  if (lastRow < 1) {
    return {
      coreItems: [],
      configItems: [],
      modules: [],
      visionItems: []
    };
  }
  
  // Get all data at once for performance
  const allData = orderSheet.getRange(1, 1, lastRow, 10).getValues();
  
  // Find section boundaries
  const sections = findSectionBoundaries(allData);
  
  // Parse each section
  const coreItems = parseCoreSection(allData, sections.core, refData);
  const configItems = parseConfigSection(allData, sections.config, refData);
  const modules = parseModuleSection(allData, sections.module, refData);
  const visionItems = parseVisionSection(allData, sections.vision, refData);
  
  return { coreItems, configItems, modules, visionItems };
}

/**
 * Find section boundaries in the sheet
 */
function findSectionBoundaries(allData) {
  const sections = {
    core: { start: -1, end: -1 },
    config: { start: -1, end: -1 },
    module: { start: -1, end: -1 },
    vision: { start: -1, end: -1 }
  };
  
  for (let i = 0; i < allData.length; i++) {
    const sectionMarker = String(allData[i][0]).trim().toUpperCase();
    
    if (sectionMarker === 'CORE') {
      sections.core.start = i;
    } else if (sectionMarker === 'CONFIG') {
      if (sections.core.start >= 0 && sections.core.end < 0) {
        sections.core.end = i - 1;
      }
      sections.config.start = i;
    } else if (sectionMarker === 'MODULE') {
      if (sections.config.start >= 0 && sections.config.end < 0) {
        sections.config.end = i - 1;
      }
      sections.module.start = i;
    } else if (sectionMarker === 'VISION') {
      if (sections.module.start >= 0 && sections.module.end < 0) {
        sections.module.end = i - 1;
      }
      sections.vision.start = i;
    } else if (sectionMarker === 'TOOLING') {
      if (sections.vision.start >= 0 && sections.vision.end < 0) {
        sections.vision.end = i - 1;
      }
    }
  }
  
  // Set end to last row if not found
  if (sections.vision.start >= 0 && sections.vision.end < 0) {
    sections.vision.end = allData.length - 1;
  }
  
  return sections;
}

/**
 * Parse CORE section (read-only items)
 */
function parseCoreSection(allData, bounds, refData) {
  if (bounds.start < 0) return [];
  
  const items = [];
  const startRow = bounds.start + 1; // Skip header row
  const endRow = bounds.end >= 0 ? bounds.end : allData.length - 1;
  
  for (let i = startRow; i <= endRow; i++) {
    const row = allData[i];
    const itemNum = String(row[2]).trim(); // Column C
    const partId = String(row[3]).trim();  // Column D
    const desc = String(row[4]).trim();    // Column E
    const qty = row[5] || 1;               // Column F
    const isChecked = row[6] === true;     // Column G
    const checkDate = row[7] ? String(row[7]) : null; // Column H
    const releaseType = String(row[8]).trim() || null; // Column I
    
    if (itemNum && partId) {
      items.push({
        lineNumber: parseInt(itemNum, 10),
        sheetRow: i + 1, // Actual sheet row (1-indexed)
        id: partId,
        description: desc,
        quantity: qty,
        isChecked: isChecked,
        checkDate: checkDate,
        releaseType: releaseType
      });
    }
  }
  
  return items;
}

/**
 * Parse CONFIG section
 */
function parseConfigSection(allData, bounds, refData) {
  if (bounds.start < 0) return [];
  
  const items = [];
  const startRow = bounds.start + 1;
  const endRow = bounds.end >= 0 ? bounds.end : allData.length - 1;
  
  let currentSlot = -1;
  let currentConfig = null;
  
  for (let i = startRow; i <= endRow; i++) {
    const row = allData[i];
    const itemNum = String(row[2]).trim();
    const partId = String(row[3]).trim();
    const desc = String(row[4]).trim();
    
    if (itemNum && itemNum !== '') {
      // This is a parent config item
      if (currentConfig) {
        items.push(currentConfig);
      }
      
      currentSlot++;
      const isSpecial = CONFIG.SPECIAL_TRIGGERS.hasOwnProperty(partId);
      
      currentConfig = {
        slotIndex: currentSlot,
        selectedId: partId || null,
        description: desc,
        isSpecialTrigger: isSpecial,
        shoppingListSelections: isSpecial ? [] : []
      };
    } else if (currentConfig && currentConfig.isSpecialTrigger && partId) {
      // This is a shopping list item under a special trigger
      currentConfig.shoppingListSelections.push(partId);
    }
  }
  
  // Don't forget the last config item
  if (currentConfig) {
    items.push(currentConfig);
  }
  
  return items;
}

/**
 * Parse MODULE section with children
 */
function parseModuleSection(allData, bounds, refData) {
  if (bounds.start < 0) return [];
  
  const modules = [];
  const startRow = bounds.start + 1;
  const endRow = bounds.end >= 0 ? bounds.end : allData.length - 1;
  
  let currentSlot = -1;
  let currentModule = null;
  
  // Track parent counts for instance numbering
  const parentCounts = {};
  
  for (let i = startRow; i <= endRow; i++) {
    const row = allData[i];
    const itemNum = String(row[2]).trim();
    const partId = String(row[3]).trim();
    const desc = String(row[4]).trim();
    
    if (itemNum && itemNum !== '') {
      // This is a parent module
      if (currentModule) {
        modules.push(currentModule);
      }
      
      currentSlot++;
      
      // Find module config in refData
      const moduleConfig = refData.moduleItems.find(m => m.id === partId);
      
      // Track instance count
      if (!parentCounts[partId]) {
        parentCounts[partId] = 0;
      }
      parentCounts[partId]++;
      
      currentModule = {
        slotIndex: currentSlot,
        parentId: partId || null,
        parentDescription: desc || (moduleConfig ? moduleConfig.description : ''),
        instanceNumber: parentCounts[partId],
        instanceTotal: 0, // Will be calculated later
        children: {
          electrical: null,
          tooling: [],
          jigs: [],
          vision: null
        }
      };
      
      // Pre-populate children structure from mappings
      if (moduleConfig && moduleConfig.mappings) {
        currentModule.children = parseModuleChildren(partId, moduleConfig, currentModule.instanceNumber, refData);
      }
    } else if (currentModule && partId) {
      // This is a child row - identify its type
      identifyAndUpdateChild(currentModule, partId, desc, refData);
    }
  }
  
  // Don't forget the last module
  if (currentModule) {
    modules.push(currentModule);
  }
  
  // Calculate total instance counts
  modules.forEach(mod => {
    if (mod.parentId && parentCounts[mod.parentId]) {
      mod.instanceTotal = parentCounts[mod.parentId];
    }
  });
  
  return modules;
}

/**
 * Parse module children from mappings
 */
function parseModuleChildren(parentId, moduleConfig, instanceNumber, refData) {
  const mappings = moduleConfig.mappings;
  const children = {
    electrical: null,
    tooling: [],
    jigs: [],
    vision: null
  };
  
  // Electrical (rotational)
  if (mappings.elecIds) {
    const elecIds = mappings.elecIds.split(';').map(s => s.trim()).filter(s => s);
    const elecDescs = (mappings.elecDesc || '').split(';').map(s => s.trim());
    
    if (elecIds.length > 0) {
      const rotationIndex = (instanceNumber - 1) % elecIds.length;
      children.electrical = {
        autoSelectedId: elecIds[rotationIndex],
        currentId: elecIds[rotationIndex],
        description: elecDescs[rotationIndex] || '',
        isOverridden: false,
        options: elecIds.map((id, i) => ({ id, description: elecDescs[i] || '' })),
        rotationIndex: rotationIndex,
        rotationTotal: elecIds.length
      };
    }
  }
  
  // Tooling (stacked)
  if (mappings.toolIds) {
    const toolIds = mappings.toolIds.split(';').map(s => s.trim()).filter(s => s);
    const toolDescs = (mappings.toolDesc || '').split(';').map(s => s.trim());
    
    children.tooling = toolIds.map((id, i) => ({
      id,
      description: toolDescs[i] || '',
      selectedOption: null,
      optionChoices: refData.toolingOptions[id] || []
    }));
  }
  
  // Jigs (stacked)
  if (mappings.jigIds) {
    const jigIds = mappings.jigIds.split(';').map(s => s.trim()).filter(s => s);
    const jigDescs = (mappings.jigDesc || '').split(';').map(s => s.trim());
    
    children.jigs = jigIds.map((id, i) => ({
      id,
      description: jigDescs[i] || ''
    }));
  }
  
  // Vision
  if (mappings.visionIds) {
    const visionIds = mappings.visionIds.split(';').map(s => s.trim()).filter(s => s);
    
    if (visionIds.length === 1) {
      const visionItem = refData.visionItems.find(v => v.id === visionIds[0]);
      children.vision = {
        type: 'fixed',
        selectedId: visionIds[0],
        description: visionItem?.description || '',
        category: visionItem?.category || '',
        options: []
      };
    } else if (visionIds.length > 1) {
      children.vision = {
        type: 'select',
        selectedId: null,
        description: '',
        category: '',
        options: visionIds.map(id => {
          const visionItem = refData.visionItems.find(v => v.id === id);
          return { id, description: visionItem?.description || '', category: visionItem?.category || '' };
        })
      };
    }
  }
  
  return children;
}

/**
 * Identify child type and update module children
 */
function identifyAndUpdateChild(module, partId, desc, refData) {
  // Skip placeholder rows
  if (partId === '-' || partId === '-,-') return;
  
  // Check if it's an electrical kit (update current if overridden)
  if (module.children.electrical) {
    const elecOption = module.children.electrical.options.find(o => o.id === partId);
    if (elecOption) {
      module.children.electrical.currentId = partId;
      module.children.electrical.description = elecOption.description;
      module.children.electrical.isOverridden = (partId !== module.children.electrical.autoSelectedId);
      return;
    }
  }
  
  // Check if it's a tooling option
  for (const tool of module.children.tooling) {
    const option = tool.optionChoices.find(o => o.id === partId);
    if (option) {
      tool.selectedOption = partId;
      return;
    }
  }
  
  // Check if it's a vision selection
  if (module.children.vision && module.children.vision.type === 'select') {
    const visionOption = module.children.vision.options.find(o => o.id === partId);
    if (visionOption) {
      module.children.vision.selectedId = partId;
      module.children.vision.description = visionOption.description;
      module.children.vision.category = visionOption.category;
      return;
    }
  }
  
  // Otherwise, it's an unknown child (could be from external modification)
  console.log('Unknown child in module:', partId);
}

/**
 * Parse VISION section (standalone vision items)
 */
function parseVisionSection(allData, bounds, refData) {
  if (bounds.start < 0) return [];
  
  const items = [];
  const startRow = bounds.start + 1;
  const endRow = bounds.end >= 0 ? bounds.end : allData.length - 1;
  
  let slotIndex = 0;
  
  for (let i = startRow; i <= endRow; i++) {
    const row = allData[i];
    const itemNum = String(row[2]).trim();
    const partId = String(row[3]).trim();
    const desc = String(row[4]).trim();
    const category = String(row[1]).trim(); // Column B has category
    
    if (itemNum && partId) {
      // Find vision item in refData for validation
      const visionItem = refData.visionItems.find(v => v.id === partId);
      
      items.push({
        slotIndex: slotIndex++,
        selectedId: partId,
        description: visionItem ? visionItem.description : desc,
        category: visionItem ? visionItem.category : category
      });
    }
  }
  
  return items;
}

// =========================================
// UTILITY FUNCTIONS
// =========================================

/**
 * Get the password for uncheck operation from Script Properties
 */
function getUncheckPassword() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty('UNCHECK_PASSWORD') || '123'; // Default fallback
}

/**
 * Set the password for uncheck operation
 */
function setUncheckPassword(newPassword) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('UNCHECK_PASSWORD', newPassword);
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

/**
 * Validate password for uncheck operation
 */
function validatePassword(password) {
  const correctPassword = getUncheckPassword();
  return { valid: password === correctPassword };
}

// =========================================
// PHASE 2: CONFIG OPERATIONS
// =========================================

/**
 * Add a config item to a slot
 * Handles special triggers (430001-A378, 430001-A714) that expand shopping lists
 * 
 * @param {number} slotIndex - 0-based slot index (0-9)
 * @param {string} partId - Part ID to add
 * @returns {Object} { success, data: { configItem, shoppingList? }, error? }
 */
function addConfigItem(slotIndex, partId) {
  try {
    // Validate inputs
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    if (!partId || partId.trim() === '') {
      throw new Error('Part ID is required.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!orderSheet || !refSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Find CONFIG section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const configBounds = findConfigSectionBounds(allData);
    
    if (configBounds.start < 0) {
      throw new Error('CONFIG section not found in sheet');
    }
    
    // Calculate target row (CONFIG header + 1 + slotIndex)
    // But need to account for existing shopping list rows
    const targetRow = findConfigSlotRow(orderSheet, configBounds, slotIndex);
    
    // Get current value to check if we need to clean up old shopping list
    const currentPartId = orderSheet.getRange(targetRow, 4).getValue();
    
    // If current slot has a special trigger, we need to delete its shopping list first
    if (CONFIG.SPECIAL_TRIGGERS.hasOwnProperty(currentPartId)) {
      deleteShoppingListRows(orderSheet, targetRow, CONFIG.SPECIAL_TRIGGERS[currentPartId].shoppingListSize);
    }
    
    // Validate partId exists in REF_DATA
    const configItems = loadConfigItems(refSheet);
    const validItem = configItems.find(item => item.id === partId);
    
    if (!validItem) {
      throw new Error(`Part ID "${partId}" not found in reference data.`);
    }
    
    // Update the config slot
    orderSheet.getRange(targetRow, 4).setValue(partId); // Column D - Part ID
    
    // Check if this is a special trigger
    const isSpecialTrigger = CONFIG.SPECIAL_TRIGGERS.hasOwnProperty(partId);
    let shoppingList = null;
    
    if (isSpecialTrigger) {
      const triggerConfig = CONFIG.SPECIAL_TRIGGERS[partId];
      
      // Insert shopping list rows after the config item
      shoppingList = insertShoppingListRows(
        orderSheet, 
        refSheet, 
        targetRow, 
        triggerConfig
      );
    }
    
    // Build response
    const configItem = {
      slotIndex: slotIndex,
      selectedId: partId,
      description: validItem.description,
      isSpecialTrigger: isSpecialTrigger,
      shoppingListSelections: isSpecialTrigger ? new Array(CONFIG.SPECIAL_TRIGGERS[partId].shoppingListSize).fill(null) : []
    };
    
    return {
      success: true,
      data: {
        configItem: configItem,
        shoppingList: shoppingList
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('addConfigItem error:', error);
    return {
      success: false,
      error: {
        code: 'CONFIG_ADD_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Remove a config item from a slot
 * Also removes shopping list rows if it was a special trigger
 * 
 * @param {number} slotIndex - 0-based slot index (0-9)
 * @returns {Object} { success, error? }
 */
function removeConfigItem(slotIndex) {
  try {
    // Validate input
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!orderSheet) {
      throw new Error('ORDERING LIST sheet not found');
    }
    
    // Find CONFIG section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const configBounds = findConfigSectionBounds(allData);
    
    if (configBounds.start < 0) {
      throw new Error('CONFIG section not found in sheet');
    }
    
    // Calculate target row
    const targetRow = findConfigSlotRow(orderSheet, configBounds, slotIndex);
    
    // Get current value
    const currentPartId = orderSheet.getRange(targetRow, 4).getValue();
    
    if (!currentPartId || currentPartId === '') {
      throw new Error('Slot is already empty.');
    }
    
    // If current slot has a special trigger, delete its shopping list first
    if (CONFIG.SPECIAL_TRIGGERS.hasOwnProperty(currentPartId)) {
      deleteShoppingListRows(orderSheet, targetRow, CONFIG.SPECIAL_TRIGGERS[currentPartId].shoppingListSize);
    }
    
    // Clear the config slot (columns D, E for Part ID and Description)
    orderSheet.getRange(targetRow, 4, 1, 2).clearContent();
    
    return {
      success: true,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('removeConfigItem error:', error);
    return {
      success: false,
      error: {
        code: 'CONFIG_REMOVE_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Update a shopping list selection for a special config trigger
 * 
 * @param {number} slotIndex - Config slot index (0-9)
 * @param {number} shoppingIndex - Shopping list item index (0-based within the shopping list)
 * @param {string} partId - Selected part ID from shopping list options
 * @returns {Object} { success, data: { selectedId, description }, error? }
 */
function updateConfigShoppingList(slotIndex, shoppingIndex, partId) {
  try {
    // Validate inputs
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!orderSheet || !refSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Find CONFIG section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const configBounds = findConfigSectionBounds(allData);
    
    if (configBounds.start < 0) {
      throw new Error('CONFIG section not found in sheet');
    }
    
    // Find the config slot row
    const configSlotRow = findConfigSlotRow(orderSheet, configBounds, slotIndex);
    
    // Get the parent config item to verify it's a special trigger
    const parentPartId = orderSheet.getRange(configSlotRow, 4).getValue();
    
    if (!CONFIG.SPECIAL_TRIGGERS.hasOwnProperty(parentPartId)) {
      throw new Error('This config item does not have a shopping list.');
    }
    
    const triggerConfig = CONFIG.SPECIAL_TRIGGERS[parentPartId];
    
    // Validate shopping index
    if (shoppingIndex < 0 || shoppingIndex >= triggerConfig.shoppingListSize) {
      throw new Error(`Invalid shopping index. Must be 0-${triggerConfig.shoppingListSize - 1}.`);
    }
    
    // Calculate the shopping list row (parent row + 1 + shoppingIndex)
    const shoppingRow = configSlotRow + 1 + shoppingIndex;
    
    // Update the part ID in the shopping list row
    orderSheet.getRange(shoppingRow, 4).setValue(partId || ''); // Column D
    
    // Get description from REF_DATA if partId provided
    let description = '';
    if (partId) {
      const shoppingLists = loadShoppingLists(refSheet);
      const listKey = parentPartId === '430001-A378' ? 'basicTool' : 'pneumatic';
      const item = shoppingLists[listKey].find(i => i.id === partId);
      description = item ? item.description : '';
    }
    
    return {
      success: true,
      data: {
        selectedId: partId || null,
        description: description
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('updateConfigShoppingList error:', error);
    return {
      success: false,
      error: {
        code: 'SHOPPING_UPDATE_ERROR',
        message: error.message
      }
    };
  }
}

// =========================================
// CONFIG HELPER FUNCTIONS
// =========================================

/**
 * Find CONFIG section boundaries in the sheet
 */
function findConfigSectionBounds(allData) {
  let start = -1;
  let end = -1;
  
  for (let i = 0; i < allData.length; i++) {
    const sectionMarker = String(allData[i][0]).trim().toUpperCase();
    
    if (sectionMarker === 'CONFIG') {
      start = i;
    } else if (sectionMarker === 'MODULE' && start >= 0) {
      end = i - 1;
      break;
    }
  }
  
  return { start, end };
}

/**
 * Find the actual row number for a config slot, accounting for shopping list rows
 */
function findConfigSlotRow(orderSheet, configBounds, targetSlot) {
  const startRow = configBounds.start + 2; // +1 for 0-index to 1-index, +1 to skip header
  
  // We need to traverse the CONFIG section to find the correct slot
  // Each slot with a special trigger will have additional shopping list rows
  
  let currentSlot = 0;
  let currentRow = startRow;
  const lastRow = orderSheet.getLastRow();
  
  while (currentSlot <= targetSlot && currentRow <= lastRow) {
    if (currentSlot === targetSlot) {
      return currentRow;
    }
    
    // Check if this row has an ITEM number (indicating it's a config slot, not a shopping list item)
    const itemNum = orderSheet.getRange(currentRow, 3).getValue();
    const partId = orderSheet.getRange(currentRow, 4).getValue();
    
    if (itemNum && itemNum !== '') {
      // This is a config slot row
      
      // Check if it's a special trigger with shopping list
      if (CONFIG.SPECIAL_TRIGGERS.hasOwnProperty(partId)) {
        // Skip over the shopping list rows
        currentRow += CONFIG.SPECIAL_TRIGGERS[partId].shoppingListSize + 1;
        currentSlot++;
      } else {
        currentRow++;
        currentSlot++;
      }
    } else {
      // Skip non-slot rows (shouldn't happen in normal flow)
      currentRow++;
    }
  }
  
  // If we couldn't find it, return the next available row
  return currentRow;
}

/**
 * Insert shopping list rows after a config item
 */
function insertShoppingListRows(orderSheet, refSheet, parentRow, triggerConfig) {
  const shoppingListSize = triggerConfig.shoppingListSize;
  
  // Insert blank rows after the parent
  orderSheet.insertRowsAfter(parentRow, shoppingListSize);
  
  // Load shopping list options from REF_DATA
  const shoppingLists = loadShoppingLists(refSheet);
  const listKey = triggerConfig.shoppingListColumns.ids === 9 ? 'basicTool' : 'pneumatic';
  const options = shoppingLists[listKey];
  
  // Set up each shopping list row with data validation (dropdown)
  for (let i = 0; i < shoppingListSize; i++) {
    const row = parentRow + 1 + i;
    
    // Set up data validation for the Part ID column (D)
    const validationRange = refSheet.getRange(
      1, 
      triggerConfig.shoppingListColumns.ids, 
      refSheet.getLastRow(), 
      1
    );
    
    const validation = SpreadsheetApp.newDataValidation()
      .requireValueInRange(validationRange, true)
      .setAllowInvalid(false)
      .build();
    
    orderSheet.getRange(row, 4).setDataValidation(validation);
    
    // Set a VLOOKUP formula for the description column (E)
    const formula = `=IF(D${row}="","",VLOOKUP(D${row},REF_DATA!$${columnLetter(triggerConfig.shoppingListColumns.ids)}:$${columnLetter(triggerConfig.shoppingListColumns.descriptions)},2,FALSE))`;
    orderSheet.getRange(row, 5).setFormula(formula);
    
    // Set default quantity to 1
    orderSheet.getRange(row, 6).setValue(1);
  }
  
  return {
    size: shoppingListSize,
    options: options
  };
}

/**
 * Delete shopping list rows after a config item
 */
function deleteShoppingListRows(orderSheet, parentRow, count) {
  // Delete rows from parentRow + 1 to parentRow + count
  if (count > 0) {
    orderSheet.deleteRows(parentRow + 1, count);
  }
}

/**
 * Convert column number to letter (1 = A, 2 = B, etc.)
 */
function columnLetter(columnNumber) {
  let letter = '';
  let temp = columnNumber;
  while (temp > 0) {
    let remainder = (temp - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    temp = Math.floor((temp - 1) / 26);
  }
  return letter;
}

// =========================================
// PHASE 3: MODULE OPERATIONS
// =========================================

/**
 * Add a module to a slot (triggers child insertion)
 * Implements electrical rotation, tooling options, jigs, and vision
 * 
 * @param {number} slotIndex - 0-based slot index (0-9)
 * @param {string} parentId - Parent module Part ID
 * @returns {Object} { success, data: { module }, error? }
 */
function addModule(slotIndex, parentId) {
  try {
    // Validate inputs
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    if (!parentId || parentId.trim() === '') {
      throw new Error('Part ID is required.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!orderSheet || !refSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Find MODULE section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const moduleBounds = findModuleSectionBounds(allData);
    
    if (moduleBounds.start < 0) {
      throw new Error('MODULE section not found in sheet');
    }
    
    // Calculate target row for this module slot
    const targetRow = findModuleSlotRow(orderSheet, moduleBounds, slotIndex);
    
    // Get current value to check if we need to clean up old children
    const currentPartId = orderSheet.getRange(targetRow, 4).getValue();
    
    // If current slot has a module, delete its children first
    if (currentPartId && currentPartId !== '') {
      deleteModuleChildren(orderSheet, refSheet, targetRow, currentPartId);
    }
    
    // Validate parentId exists in REF_DATA
    const moduleConfig = findModuleConfig(refSheet, parentId);
    
    if (!moduleConfig) {
      throw new Error(`Module "${parentId}" not found in reference data.`);
    }
    
    // Update the module slot with new parent ID
    orderSheet.getRange(targetRow, 4).setValue(parentId);
    
    // Calculate instance count for electrical rotation
    const instanceCount = calculateInstanceCount(orderSheet, moduleBounds, targetRow, parentId);
    
    // Build children to insert
    const childrenToInsert = buildModuleChildren(refSheet, parentId, moduleConfig, instanceCount);
    
    // Insert all children rows
    insertModuleChildren(orderSheet, refSheet, targetRow, childrenToInsert);
    
    // Build response
    const moduleData = {
      slotIndex: slotIndex,
      instanceNumber: instanceCount,
      parentId: parentId,
      parentDescription: moduleConfig.description,
      children: childrenToInsert
    };
    
    return {
      success: true,
      data: {
        module: moduleData
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('addModule error:', error);
    return {
      success: false,
      error: {
        code: 'MODULE_ADD_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Remove a module from a slot (removes all children)
 * 
 * @param {number} slotIndex - 0-based slot index (0-9)
 * @returns {Object} { success, error? }
 */
function removeModule(slotIndex) {
  try {
    // Validate input
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!orderSheet || !refSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Find MODULE section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const moduleBounds = findModuleSectionBounds(allData);
    
    if (moduleBounds.start < 0) {
      throw new Error('MODULE section not found in sheet');
    }
    
    // Calculate target row
    const targetRow = findModuleSlotRow(orderSheet, moduleBounds, slotIndex);
    
    // Get current value
    const currentPartId = orderSheet.getRange(targetRow, 4).getValue();
    
    if (!currentPartId || currentPartId === '') {
      throw new Error('Slot is already empty.');
    }
    
    // Delete all children
    deleteModuleChildren(orderSheet, refSheet, targetRow, currentPartId);
    
    // Clear the module slot
    orderSheet.getRange(targetRow, 4, 1, 2).clearContent();
    
    return {
      success: true,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('removeModule error:', error);
    return {
      success: false,
      error: {
        code: 'MODULE_REMOVE_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Update a module child selection (tooling option, vision selection, etc.)
 * 
 * @param {number} slotIndex - Module slot index
 * @param {string} childType - 'toolingOption' | 'rubberTip' | 'vision'
 * @param {number} childIndex - Index within child type
 * @param {string} partId - Selected part ID
 * @returns {Object} { success, error? }
 */
function updateModuleChild(slotIndex, childType, childIndex, partId) {
  try {
    // Validate inputs
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!orderSheet) {
      throw new Error('ORDERING LIST sheet not found');
    }
    
    // Find MODULE section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const moduleBounds = findModuleSectionBounds(allData);
    
    if (moduleBounds.start < 0) {
      throw new Error('MODULE section not found in sheet');
    }
    
    // Find the module slot row
    const moduleRow = findModuleSlotRow(orderSheet, moduleBounds, slotIndex);
    
    // Find the specific child row based on type and index
    const childRow = findModuleChildRow(orderSheet, moduleRow, childType, childIndex);
    
    if (childRow < 0) {
      throw new Error(`Child of type "${childType}" at index ${childIndex} not found.`);
    }
    
    // Update the part ID in the child row
    orderSheet.getRange(childRow, 4).setValue(partId || '');
    
    return {
      success: true,
      data: {
        childType: childType,
        childIndex: childIndex,
        selectedId: partId || null
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('updateModuleChild error:', error);
    return {
      success: false,
      error: {
        code: 'MODULE_CHILD_UPDATE_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Override electrical kit (user manual override)
 * 
 * @param {number} slotIndex - Module slot index
 * @param {string} electricalId - New electrical kit Part ID
 * @returns {Object} { success, error? }
 */
function overrideElectricalKit(slotIndex, electricalId) {
  try {
    // This is essentially updating the electrical child
    // The electrical child is always the first child after the parent
    return updateModuleChild(slotIndex, 'electrical', 0, electricalId);
    
  } catch (error) {
    console.error('overrideElectricalKit error:', error);
    return {
      success: false,
      error: {
        code: 'ELECTRICAL_OVERRIDE_ERROR',
        message: error.message
      }
    };
  }
}

// =========================================
// MODULE HELPER FUNCTIONS
// =========================================

/**
 * Find MODULE section boundaries in the sheet
 */
function findModuleSectionBounds(allData) {
  let start = -1;
  let end = -1;
  
  for (let i = 0; i < allData.length; i++) {
    const sectionMarker = String(allData[i][0]).trim().toUpperCase();
    
    if (sectionMarker === 'MODULE') {
      start = i;
    } else if (sectionMarker === 'VISION' && start >= 0) {
      end = i - 1;
      break;
    }
  }
  
  return { start, end };
}

/**
 * Find the actual row number for a module slot
 */
function findModuleSlotRow(orderSheet, moduleBounds, targetSlot) {
  const startRow = moduleBounds.start + 2; // +1 for 0-index to 1-index, +1 to skip header
  
  let currentSlot = 0;
  let currentRow = startRow;
  const lastRow = orderSheet.getLastRow();
  
  while (currentSlot <= targetSlot && currentRow <= lastRow) {
    if (currentSlot === targetSlot) {
      return currentRow;
    }
    
    // Check if this row has an ITEM number (indicating it's a module slot, not a child)
    const itemNum = orderSheet.getRange(currentRow, 3).getValue();
    
    if (itemNum && itemNum !== '') {
      // This is a module slot row
      // Count how many children this module has
      const childCount = countModuleChildren(orderSheet, currentRow);
      currentRow += childCount + 1; // Skip children + move to next slot
      currentSlot++;
    } else {
      // This shouldn't happen in normal flow, but skip just in case
      currentRow++;
    }
  }
  
  // If we couldn't find it, return the next available row
  return currentRow;
}

/**
 * Count how many children a module has
 */
function countModuleChildren(orderSheet, parentRow) {
  let count = 0;
  let checkRow = parentRow + 1;
  const lastRow = orderSheet.getLastRow();
  
  while (checkRow <= lastRow) {
    const itemNum = orderSheet.getRange(checkRow, 3).getValue();
    
    // If we hit a row with an ITEM number, it's the next parent
    if (itemNum && itemNum !== '') {
      break;
    }
    
    // Check if this row has content (Part ID or is part of the module block)
    const partId = orderSheet.getRange(checkRow, 4).getValue();
    if (partId !== '' || count > 0) {
      count++;
      checkRow++;
    } else {
      break;
    }
  }
  
  return count;
}

/**
 * Find module configuration in REF_DATA
 */
function findModuleConfig(refSheet, parentId) {
  // REF_DATA columns: C (Part ID), D (Description), W-AD (Mappings)
  const moduleData = refSheet.getRange('C:AD').getValues();
  
  for (let i = 1; i < moduleData.length; i++) { // Skip header row
    if (moduleData[i][0] === parentId) {
      return {
        id: moduleData[i][0],
        description: moduleData[i][1],
        elecIds: moduleData[i][20] || '', // Column W (index 20)
        elecDesc: moduleData[i][21] || '', // Column X (index 21)
        toolIds: moduleData[i][22] || '', // Column Y (index 22)
        toolDesc: moduleData[i][23] || '', // Column Z (index 23)
        jigIds: moduleData[i][24] || '', // Column AA (index 24)
        jigDesc: moduleData[i][25] || '', // Column AB (index 25)
        visionIds: moduleData[i][26] || '', // Column AC (index 26)
        visionDesc: moduleData[i][27] || ''  // Column AD (index 27)
      };
    }
  }
  
  return null;
}

/**
 * Calculate instance count for electrical rotation
 */
function calculateInstanceCount(orderSheet, moduleBounds, currentRow, parentId) {
  let count = 0;
  const startRow = moduleBounds.start + 2;
  
  for (let row = startRow; row <= currentRow; row++) {
    const itemNum = orderSheet.getRange(row, 3).getValue();
    const partId = orderSheet.getRange(row, 4).getValue();
    
    // Only count rows with ITEM numbers (parent modules)
    if (itemNum && itemNum !== '' && partId === parentId) {
      count++;
    }
  }
  
  return count;
}

/**
 * Build list of children to insert for a module
 */
function buildModuleChildren(refSheet, parentId, moduleConfig, instanceCount) {
  const children = [];
  
  // 1. ELECTRICAL KIT (Rotational)
  if (moduleConfig.elecIds) {
    const elecIds = moduleConfig.elecIds.split(';').map(s => s.trim()).filter(s => s);
    const elecDescs = moduleConfig.elecDesc.split(';').map(s => s.trim());
    
    if (elecIds.length > 0) {
      const index = (instanceCount - 1) % elecIds.length;
      children.push({
        type: 'electrical',
        id: elecIds[index],
        description: elecDescs[index] || ''
      });
    }
  }
  
  // 2. TOOLING KITS (with options and rubber tips)
  if (moduleConfig.toolIds) {
    const toolIds = moduleConfig.toolIds.split(';').map(s => s.trim()).filter(s => s);
    const toolDescs = moduleConfig.toolDesc.split(';').map(s => s.trim());
    
    for (let i = 0; i < toolIds.length; i++) {
      // Add the tooling kit itself
      children.push({
        type: 'tooling',
        id: toolIds[i],
        description: toolDescs[i] || ''
      });
      
      // Check for tooling options
      const optionRange = getToolingOptionRange(refSheet, toolIds[i]);
      if (optionRange) {
        children.push({
          type: 'toolingOption',
          parentToolId: toolIds[i],
          refDataStart: optionRange.startRow,
          refDataEnd: optionRange.endRow
        });
      }
      
      // Check for rubber tip
      if (CONFIG.RUBBER_TIP_PARENTS.includes(toolIds[i])) {
        const rtRange = getToolingOptionRange(refSheet, CONFIG.RUBBER_TIP_SOURCE_ID);
        if (rtRange) {
          children.push({
            type: 'rubberTip',
            refDataStart: rtRange.startRow,
            refDataEnd: rtRange.endRow
          });
        }
      }
    }
  }
  
  // 3. JIGS
  if (moduleConfig.jigIds) {
    const jigIds = moduleConfig.jigIds.split(';').map(s => s.trim()).filter(s => s);
    const jigDescs = moduleConfig.jigDesc.split(';').map(s => s.trim());
    
    for (let i = 0; i < jigIds.length; i++) {
      children.push({
        type: 'jig',
        id: jigIds[i],
        description: jigDescs[i] || ''
      });
    }
  }
  
  // 4. VISION
  if (moduleConfig.visionIds) {
    const visionIds = moduleConfig.visionIds.split(';').map(s => s.trim()).filter(s => s);
    
    if (visionIds.length === 1) {
      children.push({
        type: 'visionFixed',
        id: visionIds[0]
      });
    } else if (visionIds.length > 1) {
      children.push({
        type: 'visionSelect',
        ids: visionIds
      });
    }
  }
  
  return children;
}

/**
 * Get tooling option range from shadow menu (REF_DATA U:V)
 */
function getToolingOptionRange(refSheet, toolId) {
  const menuData = refSheet.getRange('U:V').getValues();
  
  let startRow = -1;
  let endRow = -1;
  let inRange = false;
  
  for (let i = 1; i < menuData.length; i++) { // Skip header
    const parentId = menuData[i][0]; // Column U
    const displayValue = menuData[i][1]; // Column V
    
    if (parentId === toolId && !inRange) {
      startRow = i + 1; // Convert to 1-indexed row number
      inRange = true;
    } else if (inRange && parentId && parentId !== toolId) {
      // Hit a new parent, end the range
      endRow = i; // Previous row was the last
      break;
    } else if (inRange && !parentId && !displayValue) {
      // Empty row, end the range
      endRow = i;
      break;
    }
  }
  
  // If we reached the end while in range
  if (inRange && endRow < 0) {
    endRow = menuData.length;
  }
  
  return (startRow > 0 && endRow > 0) ? { startRow, endRow } : null;
}

/**
 * Insert module children rows
 */
function insertModuleChildren(orderSheet, refSheet, parentRow, children) {
  if (children.length === 0) return;
  
  // Insert rows after parent
  orderSheet.insertRowsAfter(parentRow, children.length);
  
  for (let i = 0; i < children.length; i++) {
    const row = parentRow + 1 + i;
    const child = children[i];
    
    // Clear ITEM column (Col C) for all children
    orderSheet.getRange(row, 3).clearContent();
    
    if (child.type === 'electrical' || child.type === 'tooling' || child.type === 'jig') {
      // Static child - set values directly
      orderSheet.getRange(row, 4).setValue(child.id);
      orderSheet.getRange(row, 5).setValue(child.description);
      orderSheet.getRange(row, 4).clearDataValidations();
      
    } else if (child.type === 'toolingOption') {
      // Dropdown for tooling options
      const rangeNotation = `REF_DATA!V${child.refDataStart}:V${child.refDataEnd}`;
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(refSheet.getRange(rangeNotation), true)
        .setAllowInvalid(true)
        .build();
      orderSheet.getRange(row, 4).setDataValidation(rule);
      
      // Formula for Category (Col B)
      orderSheet.getRange(row, 2).setFormula(`=IFERROR(VLOOKUP(D${row}, REF_DATA!Q:S, 2, FALSE), "")`);
      
      // Formula for Description (Col E)
      orderSheet.getRange(row, 5).setFormula(`=IFERROR(VLOOKUP(D${row}, REF_DATA!Q:S, 3, FALSE), "")`);
      
    } else if (child.type === 'rubberTip') {
      // Dropdown for rubber tip options
      const rangeNotation = `REF_DATA!V${child.refDataStart}:V${child.refDataEnd}`;
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInRange(refSheet.getRange(rangeNotation), true)
        .setAllowInvalid(true)
        .build();
      orderSheet.getRange(row, 4).setDataValidation(rule);
      
      // Clear Category (Col B) for rubber tips
      orderSheet.getRange(row, 2).clearContent();
      
      // Formula for Description (Col E)
      orderSheet.getRange(row, 5).setFormula(`=IFERROR(VLOOKUP(D${row}, REF_DATA!Q:S, 3, FALSE), "")`);
      
    } else if (child.type === 'visionFixed') {
      // Fixed vision component
      orderSheet.getRange(row, 4).setValue(child.id);
      orderSheet.getRange(row, 4).clearDataValidations();
      
      // Formula for Category (Col B)
      orderSheet.getRange(row, 2).setFormula(`=IFERROR(VLOOKUP(D${row}, REF_DATA!AF:AH, 3, FALSE), "")`);
      
      // Formula for Description (Col E)
      orderSheet.getRange(row, 5).setFormula(`=IFERROR(VLOOKUP(D${row}, REF_DATA!AF:AH, 2, FALSE), "")`);
      
    } else if (child.type === 'visionSelect') {
      // Dropdown for vision selection
      const rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(child.ids, true)
        .setAllowInvalid(true)
        .build();
      orderSheet.getRange(row, 4).setDataValidation(rule);
      
      // Formula for Category (Col B)
      orderSheet.getRange(row, 2).setFormula(`=IFERROR(VLOOKUP(D${row}, REF_DATA!AF:AH, 3, FALSE), "")`);
      
      // Formula for Description (Col E)
      orderSheet.getRange(row, 5).setFormula(`=IFERROR(VLOOKUP(D${row}, REF_DATA!AF:AH, 2, FALSE), "")`);
    }
    
    // Add checkbox to Col G (RELEASED)
    orderSheet.getRange(row, 7).insertCheckboxes();
    
    // Add release type dropdown to Col I
    const releaseRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['CHARGE OUT', 'MRP'], true)
      .build();
    orderSheet.getRange(row, 9).setDataValidation(releaseRule);
  }
}

/**
 * Delete module children
 */
function deleteModuleChildren(orderSheet, refSheet, parentRow, parentId) {
  // Get module config to know what children to look for
  const moduleConfig = findModuleConfig(refSheet, parentId);
  if (!moduleConfig) return;
  
  // Build list of all possible child IDs
  const possibleChildren = [];
  
  // Electrical IDs
  if (moduleConfig.elecIds) {
    possibleChildren.push(...moduleConfig.elecIds.split(';').map(s => s.trim()).filter(s => s));
  }
  
  // Tooling IDs
  const toolIds = [];
  if (moduleConfig.toolIds) {
    const tIds = moduleConfig.toolIds.split(';').map(s => s.trim()).filter(s => s);
    possibleChildren.push(...tIds);
    toolIds.push(...tIds);
  }
  
  // Jig IDs
  if (moduleConfig.jigIds) {
    possibleChildren.push(...moduleConfig.jigIds.split(';').map(s => s.trim()).filter(s => s));
  }
  
  // Vision IDs
  if (moduleConfig.visionIds) {
    possibleChildren.push(...moduleConfig.visionIds.split(';').map(s => s.trim()).filter(s => s));
  }
  
  // Tooling option IDs and rubber tip IDs
  const optionData = refSheet.getRange('P:Q').getValues();
  for (const toolId of toolIds) {
    // Standard options
    const optionIds = getToolingOptionIDs(optionData, toolId);
    possibleChildren.push(...optionIds);
    
    // Rubber tip options
    if (CONFIG.RUBBER_TIP_PARENTS.includes(toolId)) {
      const rtIds = getToolingOptionIDs(optionData, CONFIG.RUBBER_TIP_SOURCE_ID);
      possibleChildren.push(...rtIds);
    }
  }
  
  // Delete rows
  let checkRow = parentRow + 1;
  while (checkRow <= orderSheet.getLastRow()) {
    const itemNum = orderSheet.getRange(checkRow, 3).getValue();
    const childPartId = orderSheet.getRange(checkRow, 4).getValue();
    
    // Stop if we hit the next parent (row with ITEM number)
    if (itemNum && itemNum !== '') {
      break;
    }
    
    // Delete if it's a known child OR if it's part of the module block (empty)
    if (possibleChildren.includes(childPartId) || (childPartId === '' && itemNum === '')) {
      orderSheet.deleteRow(checkRow);
      // Don't increment checkRow since rows shift up
    } else {
      break;
    }
  }
}

/**
 * Get tooling option IDs from option data (REF_DATA P:Q)
 */
function getToolingOptionIDs(optionData, parentId) {
  const ids = [];
  
  for (let i = 1; i < optionData.length; i++) { // Skip header
    if (optionData[i][0] === parentId) {
      const childId = optionData[i][1];
      if (childId) {
        ids.push(childId);
      }
    }
  }
  
  return ids;
}

/**
 * Find a specific module child row
 */
function findModuleChildRow(orderSheet, parentRow, childType, childIndex) {
  let currentIndex = 0;
  let checkRow = parentRow + 1;
  const lastRow = orderSheet.getLastRow();
  
  while (checkRow <= lastRow) {
    const itemNum = orderSheet.getRange(checkRow, 3).getValue();
    
    // Stop if we hit the next parent
    if (itemNum && itemNum !== '') {
      break;
    }
    
    // Try to identify the child type based on data validation and formulas
    const hasDataValidation = orderSheet.getRange(checkRow, 4).getDataValidation() !== null;
    const formulaB = orderSheet.getRange(checkRow, 2).getFormula();
    
    let rowType = null;
    if (!hasDataValidation) {
      // Check formula to determine if electrical, tooling, or jig
      if (childIndex === 0 && checkRow === parentRow + 1) {
        rowType = 'electrical';
      } else {
        rowType = 'tooling'; // or jig, hard to distinguish
      }
    } else if (formulaB.includes('REF_DATA!Q:S')) {
      rowType = 'toolingOption';
    } else if (formulaB === '' && hasDataValidation) {
      rowType = 'rubberTip';
    } else if (formulaB.includes('REF_DATA!AF:AH')) {
      rowType = 'vision';
    }
    
    if (rowType === childType) {
      if (currentIndex === childIndex) {
        return checkRow;
      }
      currentIndex++;
    }
    
    checkRow++;
  }
  
  return -1; // Not found
}

// =========================================
// PHASE 4: VISION OPERATIONS
// =========================================

/**
 * Add a standalone vision item to a slot
 * 
 * @param {number} slotIndex - 0-based slot index (0-9)
 * @param {string} visionId - Vision Part ID
 * @returns {Object} { success, data: { visionItem }, error? }
 */
function addVisionItem(slotIndex, visionId) {
  try {
    // Validate inputs
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    if (!visionId || visionId.trim() === '') {
      throw new Error('Vision Part ID is required.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!orderSheet || !refSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Find VISION section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const visionBounds = findVisionSectionBounds(allData);
    
    if (visionBounds.start < 0) {
      throw new Error('VISION section not found in sheet');
    }
    
    // Calculate target row for this vision slot
    const targetRow = findVisionSlotRow(orderSheet, visionBounds, slotIndex);
    
    // Get current value to check if we need to clear
    const currentPartId = orderSheet.getRange(targetRow, 4).getValue();
    
    // Validate visionId exists in REF_DATA
    const visionData = findVisionInRefData(refSheet, visionId);
    
    if (!visionData) {
      throw new Error(`Vision ID "${visionId}" not found in reference data.`);
    }
    
    // Update the vision slot with new ID
    orderSheet.getRange(targetRow, 4).setValue(visionId);
    
    // Set up formulas for category and description
    // Category Formula (Col B) -> REF_DATA!AH (Index 3 of AF:AH)
    orderSheet.getRange(targetRow, 2).setFormula(`=IFERROR(VLOOKUP(D${targetRow}, REF_DATA!AF:AH, 3, FALSE), "")`);
    
    // Description Formula (Col E) -> REF_DATA!AG (Index 2 of AF:AH)
    orderSheet.getRange(targetRow, 5).setFormula(`=IFERROR(VLOOKUP(D${targetRow}, REF_DATA!AF:AH, 2, FALSE), "")`);
    
    // Build response
    const visionItem = {
      slotIndex: slotIndex,
      id: visionId,
      description: visionData.description,
      category: visionData.category
    };
    
    return {
      success: true,
      data: {
        visionItem: visionItem
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('addVisionItem error:', error);
    return {
      success: false,
      error: {
        code: 'VISION_ADD_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Remove a standalone vision item from a slot
 * 
 * @param {number} slotIndex - 0-based slot index (0-9)
 * @returns {Object} { success, error? }
 */
function removeVisionItem(slotIndex) {
  try {
    // Validate input
    if (slotIndex < 0 || slotIndex > 9) {
      throw new Error('Invalid slot index. Must be 0-9.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!orderSheet) {
      throw new Error('ORDERING LIST sheet not found');
    }
    
    // Find VISION section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const visionBounds = findVisionSectionBounds(allData);
    
    if (visionBounds.start < 0) {
      throw new Error('VISION section not found in sheet');
    }
    
    // Calculate target row
    const targetRow = findVisionSlotRow(orderSheet, visionBounds, slotIndex);
    
    // Get current value
    const currentPartId = orderSheet.getRange(targetRow, 4).getValue();
    
    if (!currentPartId || currentPartId === '') {
      throw new Error('Slot is already empty.');
    }
    
    // Clear the vision slot (columns B, D, E for Category, Part ID, Description)
    orderSheet.getRange(targetRow, 2).clearContent(); // Category
    orderSheet.getRange(targetRow, 4).clearContent(); // Part ID
    orderSheet.getRange(targetRow, 5).clearContent(); // Description
    
    return {
      success: true,
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('removeVisionItem error:', error);
    return {
      success: false,
      error: {
        code: 'VISION_REMOVE_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Update a standalone vision item (change selection)
 * 
 * @param {number} slotIndex - 0-based slot index (0-9)
 * @param {string} visionId - New Vision Part ID
 * @returns {Object} { success, data: { visionItem }, error? }
 */
function updateVisionItem(slotIndex, visionId) {
  try {
    // This is essentially the same as addVisionItem
    // It will overwrite the existing value
    return addVisionItem(slotIndex, visionId);
    
  } catch (error) {
    console.error('updateVisionItem error:', error);
    return {
      success: false,
      error: {
        code: 'VISION_UPDATE_ERROR',
        message: error.message
      }
    };
  }
}

// =========================================
// VISION HELPER FUNCTIONS
// =========================================

/**
 * Find VISION section boundaries in the sheet
 */
function findVisionSectionBounds(allData) {
  let start = -1;
  let end = -1;
  
  for (let i = 0; i < allData.length; i++) {
    const sectionMarker = String(allData[i][0]).trim().toUpperCase();
    
    if (sectionMarker === 'VISION') {
      start = i;
    } else if ((sectionMarker === 'TOOLING' || sectionMarker === 'VCM' || sectionMarker === 'OTHERS') && start >= 0) {
      // VISION section ends before TOOLING or other sections
      end = i - 1;
      break;
    }
  }
  
  // If we didn't find an end, assume it goes to a reasonable distance
  if (start >= 0 && end < 0) {
    end = start + 15; // Max 10 items + some buffer
  }
  
  return { start, end };
}

/**
 * Find the actual row number for a vision slot
 */
function findVisionSlotRow(orderSheet, visionBounds, targetSlot) {
  const startRow = visionBounds.start + 2; // +1 for 0-index to 1-index, +1 to skip header
  
  let currentSlot = 0;
  let currentRow = startRow;
  
  while (currentSlot <= targetSlot && currentRow <= orderSheet.getLastRow()) {
    if (currentSlot === targetSlot) {
      return currentRow;
    }
    
    // Check if this row has an ITEM number (indicating it's a vision slot)
    const itemNum = orderSheet.getRange(currentRow, 3).getValue();
    
    if (itemNum && itemNum !== '') {
      // This is a vision slot row
      currentRow++;
      currentSlot++;
    } else {
      // Skip non-slot rows
      currentRow++;
    }
  }
  
  // If we couldn't find it, return the next available row
  return currentRow;
}

/**
 * Find vision item in REF_DATA (AF:AH)
 */
function findVisionInRefData(refSheet, visionId) {
  const visionData = refSheet.getRange('AF:AH').getValues();
  
  for (let i = 1; i < visionData.length; i++) { // Skip header row
    if (visionData[i][0] === visionId) {
      return {
        id: visionData[i][0],
        description: visionData[i][1],
        category: visionData[i][2]
      };
    }
  }
  
  return null;
}

// =========================================
// PHASE 5: ADMIN OPERATIONS
// =========================================

/**
 * Renumber electrical kits based on rotation
 * Recalculates electrical kit assignments for all modules based on instance counts
 * 
 * @returns {Object} { success, data: { updatedCount, modules }, error? }
 */
function renumberKits() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!orderSheet || !refSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Find MODULE section boundaries
    const allData = orderSheet.getDataRange().getValues();
    const moduleBounds = findModuleSectionBounds(allData);
    
    if (moduleBounds.start < 0) {
      throw new Error('MODULE section not found');
    }
    
    const startRow = moduleBounds.start + 2; // Skip marker and header
    const endRow = moduleBounds.end + 1; // Convert to 1-indexed
    
    // Get all Part IDs in MODULE section
    const range = orderSheet.getRange(startRow, 4, endRow - startRow + 1, 1);
    const values = range.getValues();
    
    // Get REF_DATA for module configurations
    const refData = refSheet.getRange('C:AD').getValues();
    
    // Track parent counts for rotation calculation
    const parentCounts = {};
    let updatedCount = 0;
    const updatedModules = [];
    
    // Iterate through MODULE section
    for (let i = 0; i < values.length; i++) {
      const parentID = values[i][0];
      
      if (!parentID || parentID === '') continue;
      
      const config = findModuleConfig(refSheet, parentID);
      
      if (config && config.elecIds) {
        // Increment count for this parent
        if (!parentCounts[parentID]) {
          parentCounts[parentID] = 0;
        }
        parentCounts[parentID]++;
        
        const eIds = config.elecIds.split(';').map(s => s.trim()).filter(s => s);
        const eDescs = config.elecDesc.split(';').map(s => s.trim());
        
        const count = parentCounts[parentID];
        const index = (count - 1) % eIds.length;
        const targetId = eIds[index];
        const targetDesc = eDescs[index] || '';
        
        // The electrical child should be the next row after the parent
        const childRowAbs = startRow + i + 1;
        
        // Safety check
        if (childRowAbs > endRow + 10) continue;
        
        const actualChildID = orderSheet.getRange(childRowAbs, 4).getValue();
        
        // Only update if the child is an electrical kit and it's different
        if (eIds.includes(actualChildID) && actualChildID !== targetId) {
          orderSheet.getRange(childRowAbs, 4).setValue(targetId);
          orderSheet.getRange(childRowAbs, 5).setValue(targetDesc);
          updatedCount++;
          
          updatedModules.push({
            parentId: parentID,
            instance: count,
            oldElectricalKit: actualChildID,
            newElectricalKit: targetId
          });
        }
      }
    }
    
    return {
      success: true,
      data: {
        updatedCount: updatedCount,
        modules: updatedModules,
        message: `Renumbering complete. ${updatedCount} electrical kit(s) updated.`
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('renumberKits error:', error);
    return {
      success: false,
      error: {
        code: 'RENUMBER_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Trigger full master sync from external BOM spreadsheet
 * Pulls data from the source BOM and updates REF_DATA and sections
 * 
 * @returns {Object} { success, data: { message }, error? }
 */
function triggerMasterSync() {
  try {
    const sourceSpreadsheetId = CONFIG.SOURCE_BOM_ID;
    const sourceTabName = 'BOM Structure Tree Diagram';
    
    // Try to open source spreadsheet
    let sourceSS;
    try {
      sourceSS = SpreadsheetApp.openById(sourceSpreadsheetId);
    } catch (e) {
      throw new Error('Could not open Source Spreadsheet. Check access permissions. ' + e.message);
    }
    
    const sourceSheet = sourceSS.getSheetByName(sourceTabName);
    if (!sourceSheet) {
      throw new Error(`Source tab '${sourceTabName}' not found.`);
    }
    
    const destSS = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const destSheet = destSS.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    const refSheet = destSS.getSheetByName(CONFIG.SHEETS.REF_DATA);
    
    if (!destSheet || !refSheet) {
      throw new Error('Destination sheets not found');
    }
    
    // Call the existing runMasterSync logic
    // Note: This assumes ORDERING_LIST.gs functions are available
    // If not, we need to replicate the sync logic here
    
    // For now, we'll call the existing function if it exists
    if (typeof runMasterSync === 'function') {
      runMasterSync();
      
      return {
        success: true,
        data: {
          message: 'Master sync completed successfully. REF_DATA and sections have been updated.'
        },
        timestamp: new Date().toISOString()
      };
    } else {
      throw new Error('runMasterSync function not found. Please ensure ORDERING_LIST.gs is included.');
    }
    
  } catch (error) {
    console.error('triggerMasterSync error:', error);
    return {
      success: false,
      error: {
        code: 'MASTER_SYNC_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Get sync status and metadata
 * Returns information about the last sync and current state
 * 
 * @returns {Object} { success, data: { lastSync, itemCounts }, error? }
 */
function getSyncStatus() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const refSheet = ss.getSheetByName(CONFIG.SHEETS.REF_DATA);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!refSheet || !orderSheet) {
      throw new Error('Required sheets not found');
    }
    
    // Count items in REF_DATA
    const configData = refSheet.getRange('A:A').getValues();
    const moduleData = refSheet.getRange('C:C').getValues();
    const visionData = refSheet.getRange('AF:AF').getValues();
    
    const configCount = configData.filter(row => row[0] && row[0] !== 'Part ID').length;
    const moduleCount = moduleData.filter(row => row[0] && row[0] !== 'Part ID').length;
    const visionCount = visionData.filter(row => row[0] && row[0] !== 'Part ID').length;
    
    // Try to get last modified time (this is approximate)
    const lastModified = refSheet.getLastUpdated();
    
    return {
      success: true,
      data: {
        lastSync: lastModified ? lastModified.toISOString() : null,
        itemCounts: {
          config: configCount,
          module: moduleCount,
          vision: visionCount
        },
        spreadsheetId: CONFIG.MAIN_SPREADSHEET_ID,
        sourceBomId: CONFIG.SOURCE_BOM_ID
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('getSyncStatus error:', error);
    return {
      success: false,
      error: {
        code: 'SYNC_STATUS_ERROR',
        message: error.message
      }
    };
  }
}

// =========================================
// PHASE 6: CHECKBOX OPERATIONS
// =========================================

/**
 * Check an item (add timestamp)
 * 
 * @param {number} rowNumber - Absolute row number in sheet (1-indexed)
 * @returns {Object} { success, data: { timestamp }, error? }
 */
function checkItem(rowNumber) {
  try {
    if (!rowNumber || rowNumber < 1) {
      throw new Error('Invalid row number');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!orderSheet) {
      throw new Error('ORDERING LIST sheet not found');
    }
    
    // Column G (7) - Checkbox
    // Column H (8) - Date timestamp
    
    // Set checkbox to TRUE
    orderSheet.getRange(rowNumber, 7).setValue(true);
    
    // Add timestamp to Column H
    const timestamp = Utilities.formatDate(
      new Date(), 
      ss.getSpreadsheetTimeZone(), 
      'dd/MM/yyyy'
    );
    orderSheet.getRange(rowNumber, 8).setValue(timestamp);
    
    return {
      success: true,
      data: {
        timestamp: timestamp,
        isChecked: true
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('checkItem error:', error);
    return {
      success: false,
      error: {
        code: 'CHECK_ITEM_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Uncheck an item (requires password)
 * 
 * @param {number} rowNumber - Absolute row number in sheet (1-indexed)
 * @param {string} password - User entered password
 * @returns {Object} { success, data?: { message }, error? }
 */
function uncheckItem(rowNumber, password) {
  try {
    if (!rowNumber || rowNumber < 1) {
      throw new Error('Invalid row number');
    }
    
    if (!password) {
      throw new Error('Password is required');
    }
    
    // Validate password
    const correctPassword = getUncheckPassword();
    
    if (password !== correctPassword) {
      return {
        success: false,
        error: {
          code: 'INVALID_PASSWORD',
          message: 'Incorrect password'
        }
      };
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!orderSheet) {
      throw new Error('ORDERING LIST sheet not found');
    }
    
    // Column G (7) - Checkbox
    // Column H (8) - Date timestamp
    // Column I (9) - Release Type
    
    // Set checkbox to FALSE
    orderSheet.getRange(rowNumber, 7).setValue(false);
    
    // Clear date timestamp
    orderSheet.getRange(rowNumber, 8).clearContent();
    
    // Clear release type
    orderSheet.getRange(rowNumber, 9).clearContent();
    
    return {
      success: true,
      data: {
        message: 'Item unchecked successfully',
        isChecked: false
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('uncheckItem error:', error);
    return {
      success: false,
      error: {
        code: 'UNCHECK_ITEM_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Update release type for a checked item
 * 
 * @param {number} rowNumber - Absolute row number in sheet (1-indexed)
 * @param {string} releaseType - 'CHARGE OUT' | 'MRP' | ''
 * @returns {Object} { success, error? }
 */
function updateReleaseType(rowNumber, releaseType) {
  try {
    if (!rowNumber || rowNumber < 1) {
      throw new Error('Invalid row number');
    }
    
    // Validate release type
    const validTypes = ['CHARGE OUT', 'MRP', ''];
    if (!validTypes.includes(releaseType)) {
      throw new Error('Invalid release type. Must be "CHARGE OUT", "MRP", or empty.');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!orderSheet) {
      throw new Error('ORDERING LIST sheet not found');
    }
    
    // Column I (9) - Release Type
    orderSheet.getRange(rowNumber, 9).setValue(releaseType);
    
    return {
      success: true,
      data: {
        releaseType: releaseType
      },
      timestamp: new Date().toISOString()
    };
    
  } catch (error) {
    console.error('updateReleaseType error:', error);
    return {
      success: false,
      error: {
        code: 'RELEASE_TYPE_UPDATE_ERROR',
        message: error.message
      }
    };
  }
}

/**
 * Toggle checkbox (smart toggle - checks if currently checked)
 * 
 * @param {number} rowNumber - Absolute row number in sheet (1-indexed)
 * @param {string} password - Password (required only if unchecking)
 * @returns {Object} { success, data: { isChecked, needsPassword? }, error? }
 */
function toggleCheckbox(rowNumber, password) {
  try {
    if (!rowNumber || rowNumber < 1) {
      throw new Error('Invalid row number');
    }
    
    const ss = SpreadsheetApp.openById(CONFIG.MAIN_SPREADSHEET_ID);
    const orderSheet = ss.getSheetByName(CONFIG.SHEETS.ORDERING_LIST);
    
    if (!orderSheet) {
      throw new Error('ORDERING LIST sheet not found');
    }
    
    // Get current checkbox state
    const isCurrentlyChecked = orderSheet.getRange(rowNumber, 7).getValue() === true;
    
    if (isCurrentlyChecked) {
      // Trying to uncheck - need password
      if (!password) {
        return {
          success: false,
          error: {
            code: 'PASSWORD_REQUIRED',
            message: 'Password required to uncheck item'
          },
          data: {
            needsPassword: true
          }
        };
      }
      
      // Call uncheckItem with password
      return uncheckItem(rowNumber, password);
      
    } else {
      // Checking - no password needed
      return checkItem(rowNumber);
    }
    
  } catch (error) {
    console.error('toggleCheckbox error:', error);
    return {
      success: false,
      error: {
        code: 'TOGGLE_CHECKBOX_ERROR',
        message: error.message
      }
    };
  }
}

// =========================================
// TEST FUNCTIONS (for development)
// =========================================

/**
 * Test function to verify API works
 * Run this from the GAS editor to test
 */
function testGetFullState() {
  const result = getFullState();
  console.log('Test Result:', JSON.stringify(result, null, 2));
  return result;
}

/**
 * Test addConfigItem
 */
function testAddConfigItem() {
  // Test adding a regular config item
  const result1 = addConfigItem(0, '430001-A366');
  console.log('Add regular config:', JSON.stringify(result1, null, 2));
  
  // Test adding a special trigger
  const result2 = addConfigItem(1, '430001-A714');
  console.log('Add special trigger:', JSON.stringify(result2, null, 2));
  
  return { regular: result1, special: result2 };
}

/**
 * Test removeConfigItem
 */
function testRemoveConfigItem() {
  const result = removeConfigItem(1);
  console.log('Remove config:', JSON.stringify(result, null, 2));
  return result;
}

/**
 * Test addModule
 */
function testAddModule() {
  const result = addModule(0, '430000-A961'); // Taping module
  console.log('Add module:', JSON.stringify(result, null, 2));
  return result;
}

/**
 * Test addVisionItem
 */
function testAddVisionItem() {
  const result = addVisionItem(0, '430001-A013'); // Example vision ID
  console.log('Add vision item:', JSON.stringify(result, null, 2));
  return result;
}

