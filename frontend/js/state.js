/**
 * ViTrox BOM Configurator - State Management
 * Centralized state store for the application
 */

// =========================================
// CONSTANTS - SPECIAL TRIGGER IDs
// =========================================
const SPECIAL_TRIGGERS = {
    '430001-A378': {
        name: 'Basic Tool Kit',
        icon: '🧰',
        shoppingListSize: 10,
        shoppingListKey: 'basicTool',
        bgColor: 'amber'
    },
    '430001-A714': {
        name: 'Pneumatic Kit',
        icon: '🌬',
        shoppingListSize: 3,
        shoppingListKey: 'pneumatic',
        bgColor: 'sky'
    }
};

// Helper to check if an ID is a special trigger
function isSpecialTrigger(id) {
    return SPECIAL_TRIGGERS.hasOwnProperty(id);
}

// Helper to get special trigger config
function getSpecialTriggerConfig(id) {
    return SPECIAL_TRIGGERS[id] || null;
}

// =========================================
// APPLICATION STATE
// =========================================
const AppState = {
    // Connection Status
    connection: {
        status: 'connected', // 'connected' | 'syncing' | 'error'
        lastSync: new Date(),
        spreadsheetId: '1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA'
    },

    // Reference Data (loaded from REF_DATA sheet)
    refData: {
        configItems: [],      // From REF_DATA!A:B - All config options
        moduleItems: [],      // From REF_DATA!C:D - Module options
        visionItems: [],      // From REF_DATA!AF:AH - Vision options
        toolingOptions: {},   // From REF_DATA!P:S - Tooling options
        shoppingLists: {
            basicTool: [],    // From REF_DATA!I:J (for 430001-A378)
            pneumatic: []     // From REF_DATA!K:L (for 430001-A714)
        }
    },

    // User Configuration (CONFIG section) - Flexible array, max 10 items
    configItems: [],

    // Core Items (fixed, read-only)
    coreItems: [],

    // Module Slots (max 10)
    modules: [],

    // Vision Items (standalone, max 10)
    visionItems: [],

    // Order List (computed from above)
    orderList: [],

    // UI State
    ui: {
        expandedSections: {
            core: false,
            config: true,
            module: true,
            vision: true
        },
        selectedOrderLine: null,
        filters: {
            search: '',
            section: 'all',
            status: 'all'
        },
        modals: {
            password: { open: false, targetLine: null },
            settings: { open: false },
            override: { open: false, context: null }
        }
    }
};

// =========================================
// CONFIG ITEM FUNCTIONS
// =========================================

/**
 * Add a new config item slot (local state only - for UI)
 * Actual selection triggers API call via updateConfigSelection
 */
function addConfigItem() {
    if (AppState.configItems.length >= 10) {
        showToast('Maximum 10 config items allowed', 'warning');
        return;
    }
    
    AppState.configItems.push({
        slotIndex: AppState.configItems.length,
        selectedId: null,
        description: '',
        isSpecialTrigger: false,
        shoppingListSelections: []
    });
    
    updateConfigCount();
    renderConfigSection();
}

/**
 * Remove a config item slot
 * Calls GAS API if connected, otherwise updates local state only
 */
async function removeConfigItem(slotIndex) {
    const configItem = AppState.configItems[slotIndex];
    
    // Only call API if item has a selection (is saved to sheet)
    if (configItem && configItem.selectedId && GAS_API.isGASEnvironment()) {
        try {
            showToast('Removing config item...', 'info');
            const result = await GAS_API.removeConfigItem(slotIndex);
            
            if (!result.success) {
                showToast('Failed to remove: ' + result.error.message, 'error');
                return;
            }
        } catch (error) {
            console.error('API error removing config item:', error);
            showToast('Failed to remove config item', 'error');
            return;
        }
    }
    
    // Update local state
    AppState.configItems.splice(slotIndex, 1);
    
    // Recalculate slot indices
    AppState.configItems.forEach((item, idx) => {
        item.slotIndex = idx;
    });
    
    updateConfigCount();
    rebuildOrderList();
    renderConfigSection();
    renderOrderList();
    showToast('Config item removed', 'success');
}

/**
 * Update config item selection (when user selects from dropdown)
 * Calls GAS API if connected
 */
async function updateConfigSelection(slotIndex, partId) {
    const configItem = AppState.configItems[slotIndex];
    if (!configItem) return;
    
    const oldId = configItem.selectedId;
    
    // Find the config item in refData
    const refItem = AppState.refData.configItems.find(item => item.id === partId);
    
    if (!partId || !refItem) {
        // Clearing selection - if GAS connected and had previous value, call remove
        if (oldId && GAS_API.isGASEnvironment()) {
            try {
                showToast('Clearing config item...', 'info');
                const result = await GAS_API.removeConfigItem(slotIndex);
                if (!result.success) {
                    showToast('Failed to clear: ' + result.error.message, 'error');
                    return;
                }
            } catch (error) {
                console.error('API error clearing config item:', error);
            }
        }
        
        configItem.selectedId = null;
        configItem.description = '';
        configItem.isSpecialTrigger = false;
        configItem.shoppingListSelections = [];
    } else {
        // Setting/changing selection - call API if connected
        if (GAS_API.isGASEnvironment()) {
            try {
                showToast('Saving config item...', 'info');
                const result = await GAS_API.addConfigItem(slotIndex, partId);
                
                if (!result.success) {
                    showToast('Failed to save: ' + result.error.message, 'error');
                    return;
                }
                
                // Update local state from API response
                const apiData = result.data.configItem;
                configItem.selectedId = apiData.selectedId;
                configItem.description = apiData.description;
                configItem.isSpecialTrigger = apiData.isSpecialTrigger;
                configItem.shoppingListSelections = apiData.shoppingListSelections || [];
                
                showToast('Config item saved', 'success');
            } catch (error) {
                console.error('API error saving config item:', error);
                showToast('Failed to save config item', 'error');
                return;
            }
        } else {
            // No GAS environment - update local state only
            configItem.selectedId = partId;
            configItem.description = refItem.description;
            
            // Check if this is a special trigger
            const triggerConfig = getSpecialTriggerConfig(partId);
            if (triggerConfig) {
                configItem.isSpecialTrigger = true;
                configItem.shoppingListSelections = new Array(triggerConfig.shoppingListSize).fill(null);
            } else {
                configItem.isSpecialTrigger = false;
                configItem.shoppingListSelections = [];
            }
        }
    }
    
    rebuildOrderList();
    renderConfigSection();
    renderOrderList();
}

/**
 * Update shopping list selection for a special trigger config item
 * Calls GAS API if connected
 */
async function updateConfigShoppingSelection(slotIndex, shoppingIndex, partId) {
    const configItem = AppState.configItems[slotIndex];
    if (!configItem || !configItem.isSpecialTrigger) return;
    
    // Call GAS API if connected
    if (GAS_API.isGASEnvironment()) {
        try {
            const result = await GAS_API.updateConfigShoppingList(slotIndex, shoppingIndex, partId);
            
            if (!result.success) {
                showToast('Failed to update shopping list: ' + result.error.message, 'error');
                return;
            }
        } catch (error) {
            console.error('API error updating shopping list:', error);
            showToast('Failed to update shopping list', 'error');
            return;
        }
    }
    
    // Update local state
    configItem.shoppingListSelections[shoppingIndex] = partId || null;
    
    rebuildOrderList();
    renderOrderList();
}

/**
 * Get available config options (excluding already selected ones)
 */
function getAvailableConfigOptions(currentSlotIndex) {
    const selectedIds = AppState.configItems
        .filter((item, idx) => idx !== currentSlotIndex && item.selectedId)
        .map(item => item.selectedId);
    
    return AppState.refData.configItems.filter(item => !selectedIds.includes(item.id));
}

// =========================================
// MODULE FUNCTIONS
// =========================================

/**
 * Add a new module slot (local state only - for UI)
 * Actual selection triggers API call via updateModuleParent
 */
function addModuleSlot() {
    if (AppState.modules.length >= 10) {
        showToast('Maximum 10 modules allowed', 'warning');
        return;
    }
    
    AppState.modules.push({
        slotIndex: AppState.modules.length,
        parentId: null,
        parentDescription: '',
        instanceNumber: 0,
        instanceTotal: 0,
        children: {
            electrical: null,
            tooling: [],
            jigs: [],
            vision: null
        }
    });
    
    updateModuleCount();
    renderModuleSection();
}

/**
 * Remove a module slot
 * Calls GAS API if connected, otherwise updates local state only
 */
async function removeModuleSlot(slotIndex) {
    const module = AppState.modules[slotIndex];
    
    // Only call API if module has a selection (is saved to sheet)
    if (module && module.parentId && GAS_API.isGASEnvironment()) {
        try {
            showToast('Removing module...', 'info');
            const result = await GAS_API.removeModule(slotIndex);
            
            if (!result.success) {
                showToast('Failed to remove: ' + result.error.message, 'error');
                return;
            }
        } catch (error) {
            console.error('API error removing module:', error);
            showToast('Failed to remove module', 'error');
            return;
        }
    }
    
    // Update local state
    AppState.modules.splice(slotIndex, 1);
    
    // Recalculate slot indices
    AppState.modules.forEach((mod, idx) => {
        mod.slotIndex = idx;
    });
    
    // Recalculate instance numbers for all modules
    recalculateInstanceNumbers();
    
    updateModuleCount();
    rebuildOrderList();
    renderModuleSection();
    renderOrderList();
    showToast('Module removed', 'success');
}

/**
 * Update module parent selection
 * Calls GAS API if connected
 */
async function updateModuleParent(slotIndex, parentId) {
    const module = AppState.modules[slotIndex];
    const oldParentId = module.parentId;
    
    // Find parent config from refData
    const parentConfig = AppState.refData.moduleItems.find(item => item.id === parentId);
    
    if (!parentConfig) {
        // Clearing selection - if GAS connected and had previous value, call remove
        if (oldParentId && GAS_API.isGASEnvironment()) {
            try {
                showToast('Clearing module...', 'info');
                const result = await GAS_API.removeModule(slotIndex);
                if (!result.success) {
                    showToast('Failed to clear: ' + result.error.message, 'error');
                    return;
                }
            } catch (error) {
                console.error('API error clearing module:', error);
            }
        }
        
        module.parentId = null;
        module.parentDescription = '';
        module.children = { electrical: null, tooling: [], jigs: [], vision: null };
    } else {
        // Setting/changing selection - call API if connected
        if (GAS_API.isGASEnvironment()) {
            try {
                showToast('Saving module...', 'info');
                const result = await GAS_API.addModule(slotIndex, parentId);
                
                if (!result.success) {
                    showToast('Failed to save: ' + result.error.message, 'error');
                    return;
                }
                
                // Update local state from API response
                const apiModule = result.data.module;
                module.parentId = apiModule.parentId;
                module.parentDescription = apiModule.parentDescription;
                module.instanceNumber = apiModule.instanceNumber;
                
                // Recalculate instance totals for all modules with same parent
                recalculateInstanceNumbers();
                
                // Parse children from API response (simplified for now - will need full parsing)
                populateModuleChildrenFromAPI(module, apiModule.children);
                
                showToast('Module saved', 'success');
            } catch (error) {
                console.error('API error saving module:', error);
                showToast('Failed to save module', 'error');
                return;
            }
        } else {
            // No GAS environment - update local state only
            module.parentId = parentId;
            module.parentDescription = parentConfig.description;
            
            // Calculate instance number (for rotation)
            recalculateInstanceNumbers();
            
            // Auto-populate children based on mappings
            populateModuleChildren(module, parentConfig);
        }
    }
    
    rebuildOrderList();
    renderModuleSection();
    renderOrderList();
}

/**
 * Populate module children from API response
 */
function populateModuleChildrenFromAPI(module, children) {
    // Reset children
    module.children = {
        electrical: null,
        tooling: [],
        jigs: [],
        vision: null
    };
    
    // Parse children array from API
    for (const child of children) {
        if (child.type === 'electrical') {
            module.children.electrical = {
                id: child.id,
                description: child.description,
                isAutoSelected: true,
                options: [] // Will be populated from refData if needed
            };
        } else if (child.type === 'tooling') {
            module.children.tooling.push({
                id: child.id,
                description: child.description,
                hasOptions: false, // Will be determined by next child
                options: []
            });
        } else if (child.type === 'toolingOption') {
            // Add options to the last tooling item
            if (module.children.tooling.length > 0) {
                const lastTooling = module.children.tooling[module.children.tooling.length - 1];
                lastTooling.hasOptions = true;
                lastTooling.selectedOption = null;
                // Options will be fetched from refData if needed
            }
        } else if (child.type === 'rubberTip') {
            // Add rubber tip to the last tooling item
            if (module.children.tooling.length > 0) {
                const lastTooling = module.children.tooling[module.children.tooling.length - 1];
                lastTooling.hasRubberTip = true;
                lastTooling.selectedRubberTip = null;
            }
        } else if (child.type === 'jig') {
            module.children.jigs.push({
                id: child.id,
                description: child.description
            });
        } else if (child.type === 'visionFixed') {
            module.children.vision = {
                id: child.id,
                type: 'fixed',
                options: [child.id]
            };
        } else if (child.type === 'visionSelect') {
            module.children.vision = {
                id: null,
                type: 'select',
                options: child.ids
            };
        }
    }
}

/**
 * Recalculate instance numbers for rotation logic
 */
function recalculateInstanceNumbers() {
    const parentCounts = {};
    
    AppState.modules.forEach(module => {
        if (module.parentId) {
            if (!parentCounts[module.parentId]) {
                parentCounts[module.parentId] = { count: 0, total: 0 };
            }
            parentCounts[module.parentId].total++;
        }
    });
    
    // Reset counts for second pass
    Object.keys(parentCounts).forEach(key => {
        parentCounts[key].count = 0;
    });
    
    AppState.modules.forEach(module => {
        if (module.parentId && parentCounts[module.parentId]) {
            parentCounts[module.parentId].count++;
            module.instanceNumber = parentCounts[module.parentId].count;
            module.instanceTotal = parentCounts[module.parentId].total;
        } else {
            module.instanceNumber = 0;
            module.instanceTotal = 0;
        }
    });
}

/**
 * Populate module children based on parent config
 */
function populateModuleChildren(module, parentConfig) {
    const mappings = parentConfig.mappings || {};
    
    // Reset children
    module.children = {
        electrical: null,
        tooling: [],
        jigs: [],
        vision: null
    };
    
    // Electrical Kit (Rotational)
    if (mappings.elecIds) {
        const elecIds = mappings.elecIds.split(';').map(s => s.trim()).filter(s => s);
        const elecDescs = (mappings.elecDesc || '').split(';').map(s => s.trim());
        
        if (elecIds.length > 0) {
            const rotationIndex = (module.instanceNumber - 1) % elecIds.length;
            module.children.electrical = {
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
    
    // Tooling Kits (Stacked)
    if (mappings.toolIds) {
        const toolIds = mappings.toolIds.split(';').map(s => s.trim()).filter(s => s);
        const toolDescs = (mappings.toolDesc || '').split(';').map(s => s.trim());
        
        module.children.tooling = toolIds.map((id, i) => ({
            id,
            description: toolDescs[i] || '',
            selectedOption: null,
            optionChoices: AppState.refData.toolingOptions[id] || []
        }));
    }
    
    // Jigs (Stacked)
    if (mappings.jigIds) {
        const jigIds = mappings.jigIds.split(';').map(s => s.trim()).filter(s => s);
        const jigDescs = (mappings.jigDesc || '').split(';').map(s => s.trim());
        
        module.children.jigs = jigIds.map((id, i) => ({
            id,
            description: jigDescs[i] || ''
        }));
    }
    
    // Vision (Fixed or Select)
    if (mappings.visionIds) {
        const visionIds = mappings.visionIds.split(';').map(s => s.trim()).filter(s => s);
        
        if (visionIds.length === 1) {
            // Fixed single vision
            const visionItem = AppState.refData.visionItems.find(v => v.id === visionIds[0]);
            module.children.vision = {
                type: 'fixed',
                selectedId: visionIds[0],
                description: visionItem?.description || '',
                category: visionItem?.category || '',
                options: []
            };
        } else if (visionIds.length > 1) {
            // Multiple options (dropdown)
            module.children.vision = {
                type: 'select',
                selectedId: null,
                description: '',
                category: '',
                options: visionIds.map(id => {
                    const visionItem = AppState.refData.visionItems.find(v => v.id === id);
                    return { id, description: visionItem?.description || '', category: visionItem?.category || '' };
                })
            };
        }
    }
}

/**
 * Override electrical kit selection
 */
function overrideElectricalKit(slotIndex, newKitId) {
    const module = AppState.modules[slotIndex];
    if (!module || !module.children.electrical) return;
    
    const elec = module.children.electrical;
    const newOption = elec.options.find(opt => opt.id === newKitId);
    
    if (newOption) {
        elec.currentId = newKitId;
        elec.description = newOption.description;
        elec.isOverridden = (newKitId !== elec.autoSelectedId);
    }
    
    rebuildOrderList();
    renderModuleSection();
    renderOrderList();
    autoSave();
}

/**
 * Update tooling option selection
 */
function updateToolingOption(slotIndex, toolingIndex, optionId) {
    const module = AppState.modules[slotIndex];
    if (!module || !module.children.tooling[toolingIndex]) return;
    
    module.children.tooling[toolingIndex].selectedOption = optionId;
    
    rebuildOrderList();
    renderOrderList();
    autoSave();
}

/**
 * Update module vision selection
 */
function updateModuleVision(slotIndex, visionId) {
    const module = AppState.modules[slotIndex];
    if (!module || !module.children.vision) return;
    
    const visionItem = AppState.refData.visionItems.find(v => v.id === visionId);
    module.children.vision.selectedId = visionId;
    module.children.vision.description = visionItem?.description || '';
    module.children.vision.category = visionItem?.category || '';
    
    rebuildOrderList();
    renderOrderList();
    autoSave();
}

// =========================================
// VISION FUNCTIONS
// =========================================

/**
 * Add standalone vision item (local state only - for UI)
 * Actual selection triggers API call via updateStandaloneVision
 */
function addVisionItem() {
    if (AppState.visionItems.length >= 10) {
        showToast('Maximum 10 vision items allowed', 'warning');
        return;
    }
    
    AppState.visionItems.push({
        slotIndex: AppState.visionItems.length,
        selectedId: null,
        description: '',
        category: ''
    });
    
    updateVisionCount();
    renderVisionSection();
}

/**
 * Remove standalone vision item
 * Calls GAS API if connected, otherwise updates local state only
 */
async function removeVisionItem(slotIndex) {
    const visionItem = AppState.visionItems[slotIndex];
    
    // Only call API if item has a selection (is saved to sheet)
    if (visionItem && visionItem.selectedId && GAS_API.isGASEnvironment()) {
        try {
            showToast('Removing vision item...', 'info');
            const result = await GAS_API.removeVisionItem(slotIndex);
            
            if (!result.success) {
                showToast('Failed to remove: ' + result.error.message, 'error');
                return;
            }
        } catch (error) {
            console.error('API error removing vision item:', error);
            showToast('Failed to remove vision item', 'error');
            return;
        }
    }
    
    // Update local state
    AppState.visionItems.splice(slotIndex, 1);
    
    // Recalculate slot indices
    AppState.visionItems.forEach((item, idx) => {
        item.slotIndex = idx;
    });
    
    updateVisionCount();
    rebuildOrderList();
    renderVisionSection();
    renderOrderList();
    showToast('Vision item removed', 'success');
}

/**
 * Update standalone vision selection
 * Calls GAS API if connected
 */
async function updateStandaloneVision(slotIndex, visionId) {
    const visionSlot = AppState.visionItems[slotIndex];
    if (!visionSlot) return;
    
    const oldId = visionSlot.selectedId;
    const visionItem = AppState.refData.visionItems.find(v => v.id === visionId);
    
    if (!visionId || !visionItem) {
        // Clearing selection - if GAS connected and had previous value, call remove
        if (oldId && GAS_API.isGASEnvironment()) {
            try {
                showToast('Clearing vision item...', 'info');
                const result = await GAS_API.removeVisionItem(slotIndex);
                if (!result.success) {
                    showToast('Failed to clear: ' + result.error.message, 'error');
                    return;
                }
            } catch (error) {
                console.error('API error clearing vision item:', error);
            }
        }
        
        visionSlot.selectedId = null;
        visionSlot.description = '';
        visionSlot.category = '';
    } else {
        // Setting/changing selection - call API if connected
        if (GAS_API.isGASEnvironment()) {
            try {
                showToast('Saving vision item...', 'info');
                const result = await GAS_API.addVisionItem(slotIndex, visionId);
                
                if (!result.success) {
                    showToast('Failed to save: ' + result.error.message, 'error');
                    return;
                }
                
                // Update local state from API response
                const apiData = result.data.visionItem;
                visionSlot.selectedId = apiData.id;
                visionSlot.description = apiData.description;
                visionSlot.category = apiData.category;
                
                showToast('Vision item saved', 'success');
            } catch (error) {
                console.error('API error saving vision item:', error);
                showToast('Failed to save vision item', 'error');
                return;
            }
        } else {
            // No GAS environment - update local state only
            visionSlot.selectedId = visionId;
            visionSlot.description = visionItem.description;
            visionSlot.category = visionItem.category;
        }
    }
    
    rebuildOrderList();
    renderOrderList();
}

// =========================================
// CHECKBOX FUNCTIONS
// =========================================

/**
 * Toggle checkbox on order item
 * Calls GAS API if connected
 */
async function toggleOrderCheckbox(lineNumber) {
    const item = AppState.orderList.find(i => i.lineNumber === lineNumber);
    if (!item) return;
    
    if (item.isChecked) {
        // Unchecking - requires password
        openPasswordModal(lineNumber);
    } else {
        // Checking - add timestamp
        if (GAS_API.isGASEnvironment()) {
            try {
                // Get actual row number in sheet (need to map lineNumber to sheet row)
                const sheetRow = getSheetRowFromLineNumber(lineNumber);
                
                if (!sheetRow || !item.sheetRow) {
                    showToast('Unable to sync checkbox. Please refresh the page to reload data.', 'warning');
                    console.warn('Sheet row mapping not found for line', lineNumber);
                    return;
                }
                
                const result = await GAS_API.checkItem(sheetRow);
                
                if (!result.success) {
                    showToast('Failed to check item: ' + result.error.message, 'error');
                    return;
                }
                
                // Update local state from API response
                item.isChecked = true;
                item.checkDate = result.data.timestamp;
                
            } catch (error) {
                console.error('API error checking item:', error);
                showToast('Failed to check item', 'error');
                return;
            }
        } else {
            // No GAS environment - update local state only
            item.isChecked = true;
            item.checkDate = formatDate(new Date());
        }
        
        rebuildOrderList();
        renderOrderList();
        updateSummary();
        showToast('Item checked', 'success');
    }
}

/**
 * Confirm uncheck after password
 * Calls GAS API if connected
 */
async function confirmUncheckItem(lineNumber, password) {
    const item = AppState.orderList.find(i => i.lineNumber === lineNumber);
    if (!item) return;
    
    if (GAS_API.isGASEnvironment()) {
        try {
            // Get actual row number in sheet
            const sheetRow = getSheetRowFromLineNumber(lineNumber);
            
            if (!sheetRow || !item.sheetRow) {
                showToast('Unable to sync checkbox. Please refresh the page to reload data.', 'warning');
                console.warn('Sheet row mapping not found for line', lineNumber);
                closePasswordModal();
                return;
            }
            
            const result = await GAS_API.uncheckItem(sheetRow, password);
            
            if (!result.success) {
                if (result.error.code === 'INVALID_PASSWORD') {
                    showToast('Incorrect password', 'error');
                } else {
                    showToast('Failed to uncheck item: ' + result.error.message, 'error');
                }
                return;
            }
            
            // Update local state from API response
            item.isChecked = false;
            item.checkDate = null;
            item.releaseType = null;
            
        } catch (error) {
            console.error('API error unchecking item:', error);
            showToast('Failed to uncheck item', 'error');
            return;
        }
    } else {
        // No GAS environment - update local state only
        item.isChecked = false;
        item.checkDate = null;
        item.releaseType = null;
    }
    
    closePasswordModal();
    rebuildOrderList();
    renderOrderList();
    updateSummary();
    showToast('Item unchecked', 'success');
}

/**
 * Update release type
 * Calls GAS API if connected
 */
async function updateReleaseType(lineNumber, releaseType) {
    const item = AppState.orderList.find(i => i.lineNumber === lineNumber);
    if (!item) return;
    
    if (GAS_API.isGASEnvironment()) {
        try {
            // Get actual row number in sheet
            const sheetRow = getSheetRowFromLineNumber(lineNumber);
            
            if (!sheetRow || !item.sheetRow) {
                showToast('Unable to sync release type. Please refresh the page to reload data.', 'warning');
                console.warn('Sheet row mapping not found for line', lineNumber);
                return;
            }
            
            const result = await GAS_API.updateReleaseType(sheetRow, releaseType);
            
            if (!result.success) {
                showToast('Failed to update release type: ' + result.error.message, 'error');
                return;
            }
            
        } catch (error) {
            console.error('API error updating release type:', error);
            showToast('Failed to update release type', 'error');
            return;
        }
    }
    
    // Update local state
    item.releaseType = releaseType;
}

/**
 * Get sheet row number from line number
 * Looks up the actual sheet row from the order list item
 */
function getSheetRowFromLineNumber(lineNumber) {
    const item = AppState.orderList.find(i => i.lineNumber === lineNumber);
    
    if (!item || !item.sheetRow) {
        console.warn(`Sheet row not found for line number ${lineNumber}`);
        // Fallback: estimate based on line number
        // This is a rough approximation and may not be accurate
        return lineNumber + 3; // +3 to skip typical header rows
    }
    
    return item.sheetRow;
}

// =========================================
// ORDER LIST BUILDER
// =========================================

/**
 * Rebuild the complete order list from state
 */
function rebuildOrderList() {
    const list = [];
    let lineNumber = 1;
    
    // CORE Section
    AppState.coreItems.forEach(item => {
        list.push({
            section: 'CORE',
            lineNumber: lineNumber++,
            sheetRow: item.sheetRow || null, // Actual sheet row (for API calls)
            partId: item.id,
            description: item.description,
            quantity: item.quantity || 1,
            isChecked: item.isChecked || false,
            checkDate: item.checkDate || null,
            releaseType: item.releaseType || null,
            parentLineNumber: null,
            depth: 0,
            itemType: 'core'
        });
    });
    
    // CONFIG Section - Now using flexible configItems array
    AppState.configItems.forEach((configItem) => {
        if (configItem.selectedId) {
            const parentLine = lineNumber++;
            const triggerConfig = getSpecialTriggerConfig(configItem.selectedId);
            
            list.push({
                section: 'CONFIG',
                lineNumber: parentLine,
                partId: configItem.selectedId,
                description: configItem.description,
                quantity: 1,
                isChecked: false,
                checkDate: null,
                releaseType: null,
                parentLineNumber: null,
                depth: 0,
                itemType: configItem.isSpecialTrigger ? 'config-parent-special' : 'config-parent'
            });
            
            // If it's a special trigger, add shopping list items
            if (configItem.isSpecialTrigger && triggerConfig) {
                const shoppingListKey = triggerConfig.shoppingListKey;
                const shoppingList = AppState.refData.shoppingLists[shoppingListKey] || [];
                
                configItem.shoppingListSelections.forEach((selection, idx) => {
                    if (selection) {
                        const shopItem = shoppingList.find(i => i.id === selection);
                        list.push({
                            section: 'CONFIG',
                            lineNumber: lineNumber++,
                            partId: selection,
                            description: shopItem?.description || '',
                            quantity: 1,
                            isChecked: false,
                            checkDate: null,
                            releaseType: null,
                            parentLineNumber: parentLine,
                            depth: 1,
                            itemType: 'config-child'
                        });
                    }
                });
            }
        }
    });
    
    // MODULE Section
    AppState.modules.forEach((module, moduleIdx) => {
        if (module.parentId) {
            const parentLine = lineNumber++;
            const instanceLabel = module.instanceTotal > 1 ? ` (#${module.instanceNumber})` : '';
            
            list.push({
                section: 'MODULE',
                lineNumber: parentLine,
                partId: module.parentId,
                description: module.parentDescription + instanceLabel,
                quantity: 1,
                isChecked: false,
                checkDate: null,
                releaseType: null,
                parentLineNumber: null,
                depth: 0,
                itemType: 'module-parent',
                moduleSlot: moduleIdx
            });
            
            // Electrical Kit
            if (module.children.electrical) {
                list.push({
                    section: 'MODULE',
                    lineNumber: lineNumber++,
                    partId: module.children.electrical.currentId,
                    description: `⚡ ${module.children.electrical.description}`,
                    quantity: 1,
                    isChecked: false,
                    checkDate: null,
                    releaseType: null,
                    parentLineNumber: parentLine,
                    depth: 1,
                    itemType: 'electrical'
                });
            }
            
            // Tooling Kits
            module.children.tooling.forEach((tool, toolIdx) => {
                const toolLine = lineNumber++;
                list.push({
                    section: 'MODULE',
                    lineNumber: toolLine,
                    partId: tool.id,
                    description: `🔧 ${tool.description}`,
                    quantity: 1,
                    isChecked: false,
                    checkDate: null,
                    releaseType: null,
                    parentLineNumber: parentLine,
                    depth: 1,
                    itemType: 'tooling'
                });
                
                // Tooling Option (grandchild)
                if (tool.selectedOption) {
                    const optItem = tool.optionChoices.find(o => o.id === tool.selectedOption);
                    list.push({
                        section: 'MODULE',
                        lineNumber: lineNumber++,
                        partId: tool.selectedOption,
                        description: optItem?.description || '',
                        quantity: 1,
                        isChecked: false,
                        checkDate: null,
                        releaseType: null,
                        parentLineNumber: toolLine,
                        depth: 2,
                        itemType: 'tooling-option'
                    });
                }
            });
            
            // Jigs
            module.children.jigs.forEach(jig => {
                list.push({
                    section: 'MODULE',
                    lineNumber: lineNumber++,
                    partId: jig.id,
                    description: `📐 ${jig.description}`,
                    quantity: 1,
                    isChecked: false,
                    checkDate: null,
                    releaseType: null,
                    parentLineNumber: parentLine,
                    depth: 1,
                    itemType: 'jig'
                });
            });
            
            // Vision (within module)
            if (module.children.vision && module.children.vision.selectedId) {
                list.push({
                    section: 'MODULE',
                    lineNumber: lineNumber++,
                    partId: module.children.vision.selectedId,
                    description: `👁 ${module.children.vision.description}`,
                    quantity: 1,
                    isChecked: false,
                    checkDate: null,
                    releaseType: null,
                    parentLineNumber: parentLine,
                    depth: 1,
                    itemType: 'vision'
                });
            }
        }
    });
    
    // VISION Section (standalone)
    AppState.visionItems.forEach(visionSlot => {
        if (visionSlot.selectedId) {
            list.push({
                section: 'VISION',
                lineNumber: lineNumber++,
                partId: visionSlot.selectedId,
                description: visionSlot.description,
                quantity: 1,
                isChecked: false,
                checkDate: null,
                releaseType: null,
                parentLineNumber: null,
                depth: 0,
                itemType: 'vision-standalone',
                category: visionSlot.category
            });
        }
    });
    
    AppState.orderList = list;
    updateSummary();
}

// =========================================
// UTILITY FUNCTIONS
// =========================================

function formatDate(date) {
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    return `${day}/${month}/${year}`;
}

function updateConfigCount() {
    const countEl = document.getElementById('configCount');
    if (countEl) {
        countEl.textContent = `${AppState.configItems.length}/10`;
    }
}

function updateModuleCount() {
    const countEl = document.getElementById('moduleCount');
    if (countEl) {
        countEl.textContent = `${AppState.modules.length}/10`;
    }
}

function updateVisionCount() {
    const countEl = document.getElementById('visionCount');
    if (countEl) {
        countEl.textContent = `${AppState.visionItems.length}/10`;
    }
}

function updateCoreCount() {
    const countEl = document.getElementById('coreCount');
    if (countEl) {
        countEl.textContent = AppState.coreItems.length;
    }
}

function updateSummary() {
    const total = AppState.orderList.length;
    const checked = AppState.orderList.filter(i => i.isChecked).length;
    const pending = total - checked;
    
    document.getElementById('totalItems').textContent = total;
    document.getElementById('checkedItems').textContent = checked;
    document.getElementById('pendingItems').textContent = pending;
}

/**
 * Auto-save state (debounced)
 */
let saveTimeout = null;
function autoSave() {
    if (saveTimeout) clearTimeout(saveTimeout);
    
    saveTimeout = setTimeout(() => {
        // TODO: Implement actual save to Google Sheets via GAS
        console.log('Auto-saving state...', AppState);
        updateLastSyncTime();
    }, 500);
}

function updateLastSyncTime() {
    const el = document.getElementById('lastSyncTime');
    if (el) {
        el.textContent = `Last sync: Just now`;
    }
}
