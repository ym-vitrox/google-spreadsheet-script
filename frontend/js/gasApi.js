/**
 * ViTrox BOM Configurator - Google Apps Script API Wrapper
 * Provides promise-based interface for GAS calls with error handling
 */

const GAS_API = {
    
    // =========================================
    // CONFIGURATION
    // =========================================
    
    config: {
        // Set to true when testing locally (uses mock data)
        useMockData: false,
        
        // Timeout for API calls (ms)
        timeout: 30000,
        
        // Retry settings
        maxRetries: 2,
        retryDelay: 1000
    },
    
    // =========================================
    // CORE API CALL METHOD
    // =========================================
    
    /**
     * Call a GAS function with error handling and timeout
     * @param {string} functionName - Name of GAS function
     * @param {...any} args - Arguments to pass
     * @returns {Promise} Resolves with result or rejects with error
     */
    call: function(functionName, ...args) {
        // Check if running outside GAS (local development)
        if (typeof google === 'undefined' || !google.script) {
            if (this.config.useMockData) {
                console.log(`[GAS_API] Mock mode: ${functionName}`, args);
                return this._mockCall(functionName, args);
            } else {
                return Promise.reject(new Error('Not running in Google Apps Script environment. Enable mock mode for local testing.'));
            }
        }
        
        return new Promise((resolve, reject) => {
            const timeoutId = setTimeout(() => {
                reject(new Error(`API call timeout: ${functionName}`));
            }, this.config.timeout);
            
            google.script.run
                .withSuccessHandler((result) => {
                    clearTimeout(timeoutId);
                    console.log(`[GAS_API] Success: ${functionName}`, result);
                    resolve(result);
                })
                .withFailureHandler((error) => {
                    clearTimeout(timeoutId);
                    console.error(`[GAS_API] Error: ${functionName}`, error);
                    reject(error);
                })
                [functionName](...args);
        });
    },
    
    /**
     * Call with retry logic
     */
    callWithRetry: async function(functionName, ...args) {
        let lastError;
        
        for (let attempt = 0; attempt <= this.config.maxRetries; attempt++) {
            try {
                return await this.call(functionName, ...args);
            } catch (error) {
                lastError = error;
                console.warn(`[GAS_API] Attempt ${attempt + 1} failed for ${functionName}:`, error);
                
                if (attempt < this.config.maxRetries) {
                    await this._delay(this.config.retryDelay * (attempt + 1));
                }
            }
        }
        
        throw lastError;
    },
    
    // =========================================
    // PHASE 1: READ OPERATIONS
    // =========================================
    
    /**
     * Get complete application state on startup
     * @returns {Promise<Object>} { success, data: { refData, coreItems, configItems, modules, visionItems } }
     */
    getFullState: function() {
        return this.callWithRetry('getFullState');
    },
    
    /**
     * Get reference data only (for refresh)
     * @returns {Promise<Object>} { success, data: { configItems, moduleItems, visionItems, toolingOptions, shoppingLists } }
     */
    getRefData: function() {
        return this.callWithRetry('getRefData');
    },
    
    // =========================================
    // PHASE 2: CONFIG OPERATIONS (Placeholder)
    // =========================================
    
    /**
     * Add a config item to a slot
     * @param {number} slotIndex - 0-based slot index
     * @param {string} partId - Part ID to add
     */
    addConfigItem: function(slotIndex, partId) {
        return this.call('addConfigItem', slotIndex, partId);
    },
    
    /**
     * Remove a config item from a slot
     * @param {number} slotIndex - 0-based slot index
     */
    removeConfigItem: function(slotIndex) {
        return this.call('removeConfigItem', slotIndex);
    },
    
    /**
     * Update shopping list selection
     */
    updateConfigShoppingList: function(slotIndex, shoppingIndex, partId) {
        return this.call('updateConfigShoppingList', slotIndex, shoppingIndex, partId);
    },
    
    // =========================================
    // PHASE 3: MODULE OPERATIONS (Placeholder)
    // =========================================
    
    addModule: function(slotIndex, parentId) {
        return this.call('addModule', slotIndex, parentId);
    },
    
    removeModule: function(slotIndex) {
        return this.call('removeModule', slotIndex);
    },
    
    updateModuleChild: function(slotIndex, childType, childIndex, partId) {
        return this.call('updateModuleChild', slotIndex, childType, childIndex, partId);
    },
    
    overrideElectricalKit: function(slotIndex, electricalId) {
        return this.call('overrideElectricalKit', slotIndex, electricalId);
    },
    
    // =========================================
    // PHASE 4: VISION OPERATIONS (Placeholder)
    // =========================================
    
    addVisionItem: function(slotIndex, visionId) {
        return this.call('addVisionItem', slotIndex, visionId);
    },
    
    removeVisionItem: function(slotIndex) {
        return this.call('removeVisionItem', slotIndex);
    },
    
    updateVisionItem: function(slotIndex, visionId) {
        return this.call('updateVisionItem', slotIndex, visionId);
    },
    
    // =========================================
    // PHASE 5: ADMIN OPERATIONS
    // =========================================
    
    renumberKits: function() {
        return this.callWithRetry('renumberKits');
    },
    
    triggerMasterSync: function() {
        return this.callWithRetry('triggerMasterSync');
    },
    
    getSyncStatus: function() {
        return this.callWithRetry('getSyncStatus');
    },
    
    // =========================================
    // PHASE 6: CHECKBOX OPERATIONS
    // =========================================
    
    checkItem: function(rowNumber) {
        return this.callWithRetry('checkItem', rowNumber);
    },
    
    uncheckItem: function(rowNumber, password) {
        return this.callWithRetry('uncheckItem', rowNumber, password);
    },
    
    updateReleaseType: function(rowNumber, releaseType) {
        return this.callWithRetry('updateReleaseType', rowNumber, releaseType);
    },
    
    toggleCheckbox: function(rowNumber, password) {
        return this.callWithRetry('toggleCheckbox', rowNumber, password);
    },
    
    validatePassword: function(password) {
        return this.callWithRetry('validatePassword', password);
    },
    
    // =========================================
    // UTILITY METHODS
    // =========================================
    
    /**
     * Check if running in GAS environment
     */
    isGASEnvironment: function() {
        return typeof google !== 'undefined' && google.script;
    },
    
    /**
     * Delay helper
     */
    _delay: function(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    },
    
    /**
     * Mock API calls for local development
     */
    _mockCall: function(functionName, args) {
        return new Promise((resolve) => {
            setTimeout(() => {
                console.log(`[GAS_API] Mock response for: ${functionName}`);
                
                // Return mock data based on function name
                switch (functionName) {
                    case 'getFullState':
                        resolve(this._getMockFullState());
                        break;
                    case 'getRefData':
                        resolve(this._getMockRefData());
                        break;
                    case 'validatePassword':
                        resolve({ valid: args[0] === '123' });
                        break;
                    default:
                        resolve({ success: true, message: 'Mock response' });
                }
            }, 500); // Simulate network delay
        });
    },
    
    /**
     * Mock full state data
     */
    _getMockFullState: function() {
        return {
            success: true,
            data: {
                refData: this._getMockRefData().data,
                coreItems: [
                    { lineNumber: 1, id: '430000-A557', description: 'Main Frame Assembly', quantity: 1, isChecked: false },
                    { lineNumber: 2, id: '430000-A558', description: 'Power Distribution Unit', quantity: 1, isChecked: false },
                    { lineNumber: 3, id: '430000-A559', description: 'Control Panel Assembly', quantity: 1, isChecked: false },
                ],
                configItems: [],
                modules: [],
                visionItems: []
            },
            timestamp: new Date().toISOString()
        };
    },
    
    /**
     * Mock reference data
     */
    _getMockRefData: function() {
        return {
            success: true,
            data: {
                configItems: [
                    { id: '430001-A378', description: 'Basic Tool Kit' },
                    { id: '430001-A714', description: 'Pneumatic Kit' },
                    { id: '430001-A712', description: 'Optional Module Package' },
                    { id: '430001-A715', description: 'ESD Protection Kit' },
                ],
                moduleItems: [
                    {
                        id: '430001-A276',
                        description: 'Rotary Module V2',
                        mappings: {
                            elecIds: '430001-A501;430001-A502;430001-A503',
                            elecDesc: 'Connector Set A;Connector Set B;Connector Set C',
                            toolIds: '430001-A602',
                            toolDesc: 'Standard Tool Kit',
                            jigIds: '',
                            jigDesc: '',
                            visionIds: '',
                            visionDesc: ''
                        }
                    }
                ],
                visionItems: [
                    { id: '430001-A756', description: 'B2K Camera Setup', category: 'Position Check Vision' },
                    { id: '430001-A757', description: 'S2K Camera Setup', category: 'Position Check Vision' },
                ],
                toolingOptions: {},
                shoppingLists: {
                    basicTool: [
                        { id: '430002-A001', description: 'Screwdriver Set' },
                        { id: '430002-A002', description: 'Hex Key Set' },
                    ],
                    pneumatic: [
                        { id: '430002-A101', description: 'Air Fitting Set' },
                    ]
                }
            },
            timestamp: new Date().toISOString()
        };
    }
};

// Export for module systems (if used)
if (typeof module !== 'undefined' && module.exports) {
    module.exports = GAS_API;
}

