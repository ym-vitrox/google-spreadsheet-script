/**
 * ViTrox BOM Configurator - Main Application
 * Entry point and initialization
 */

// =========================================
// INITIALIZATION
// =========================================

document.addEventListener('DOMContentLoaded', function() {
    console.log('ViTrox BOM Configurator initializing...');
    
    // Initialize Lucide icons
    lucide.createIcons();
    
    // Set up keyboard shortcuts
    setupKeyboardShortcuts();

    // Initialize Panel Resizer
    initResizer();
    
    // Initialize the application (load data)
    initializeApp();
    
    console.log('Initialization started...');
});

/**
 * Initialize the application - loads data from GAS or falls back to mock
 */
async function initializeApp() {
    // Show loading state
    showLoadingState();
    updateConnectionStatus('syncing', 'Loading data...');
    
    try {
        // Try to load from GAS API
        if (GAS_API.isGASEnvironment()) {
            console.log('Running in GAS environment, loading real data...');
            await loadDataFromGAS();
        } else {
            console.log('Not in GAS environment, checking mock mode...');
            
            // Check if mock mode is enabled
            if (GAS_API.config.useMockData) {
                console.log('Mock mode enabled, loading mock data via API...');
                await loadDataFromGAS();
            } else {
                console.log('Using local mock data (development mode)...');
                loadMockData();
            }
        }
        
        // Initial render
        renderAllSections();
        updateConnectionStatus('connected', 'Last sync: Just now');
        showToast('Data loaded successfully', 'success');
        
        // Load sync status (non-blocking)
        loadSyncStatus();
        
    } catch (error) {
        console.error('Failed to initialize app:', error);
        updateConnectionStatus('error', 'Connection failed');
        showErrorState(error.message);
        
        // Fall back to mock data so the UI is still usable
        console.log('Falling back to mock data...');
        loadMockData();
        renderAllSections();
        showToast('Loaded with demo data (connection failed)', 'warning');
    }
    
    // Hide loading state
    hideLoadingState();
}

/**
 * Load data from Google Apps Script API
 */
async function loadDataFromGAS() {
    const result = await GAS_API.getFullState();
    
    if (!result.success) {
        throw new Error(result.error?.message || 'Failed to load data from server');
    }
    
    const data = result.data;
    
    // Populate AppState with real data
    AppState.refData = data.refData;
    AppState.coreItems = data.coreItems || [];
    AppState.configItems = data.configItems || [];
    AppState.modules = data.modules || [];
    AppState.visionItems = data.visionItems || [];
    
    // Update connection info
    AppState.connection.status = 'connected';
    AppState.connection.lastSync = new Date();
    
    // Rebuild order list
    rebuildOrderList();
    
    console.log('Data loaded from GAS:', {
        coreItems: AppState.coreItems.length,
        configItems: AppState.configItems.length,
        modules: AppState.modules.length,
        visionItems: AppState.visionItems.length,
        refData: {
            configOptions: AppState.refData.configItems.length,
            moduleOptions: AppState.refData.moduleItems.length,
            visionOptions: AppState.refData.visionItems.length
        }
    });
}

/**
 * Refresh reference data only (not current state)
 */
async function refreshRefData() {
    try {
        updateConnectionStatus('syncing', 'Refreshing...');
        
        const result = await GAS_API.getRefData();
        
        if (result.success) {
            AppState.refData = result.data;
            AppState.connection.lastSync = new Date();
            updateConnectionStatus('connected', 'Last sync: Just now');
            showToast('Reference data refreshed', 'success');
            renderAllSections();
        } else {
            throw new Error(result.error?.message || 'Refresh failed');
        }
    } catch (error) {
        console.error('Refresh error:', error);
        updateConnectionStatus('error', 'Refresh failed');
        showToast('Failed to refresh data', 'error');
    }
}

// =========================================
// LOADING / ERROR STATES
// =========================================

function showLoadingState() {
    // Show loading overlay
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
        loadingOverlay.classList.remove('hidden');
    }
    
    // Update sections to show skeleton loading
    const sections = ['coreContent', 'configContent', 'moduleContent', 'visionContent'];
    sections.forEach(sectionId => {
        const section = document.getElementById(sectionId);
        if (section) {
            section.innerHTML = `
                <div class="space-y-3">
                    <div class="skeleton h-12 w-full"></div>
                    <div class="skeleton h-12 w-full"></div>
                    <div class="skeleton h-12 w-3/4"></div>
                </div>
            `;
        }
    });
    
    // Update order table
    const orderBody = document.getElementById('orderTableBody');
    if (orderBody) {
        orderBody.innerHTML = `
            <tr>
                <td colspan="7" class="px-4 py-12 text-center">
                    <div class="flex flex-col items-center gap-3">
                        <div class="w-8 h-8 border-3 border-brand-600 border-t-transparent rounded-full animate-spin"></div>
                        <p class="text-slate-500">Loading order data...</p>
                    </div>
                </td>
            </tr>
        `;
    }
}

function hideLoadingState() {
    const loadingOverlay = document.getElementById('loadingOverlay');
    if (loadingOverlay) {
        loadingOverlay.classList.add('hidden');
    }
}

function showErrorState(message) {
    const orderBody = document.getElementById('orderTableBody');
    if (orderBody) {
        orderBody.innerHTML = `
            <tr>
                <td colspan="7" class="px-4 py-12 text-center">
                    <div class="flex flex-col items-center gap-3">
                        <div class="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center">
                            <i data-lucide="alert-circle" class="w-8 h-8 text-red-500"></i>
                        </div>
                        <p class="text-slate-700 font-medium">Failed to load data</p>
                        <p class="text-sm text-slate-500">${message}</p>
                        <button onclick="initializeApp()" class="px-4 py-2 bg-brand-600 text-white rounded-lg hover:bg-brand-700 transition-colors">
                            Retry
                        </button>
                    </div>
                </td>
            </tr>
        `;
        lucide.createIcons();
    }
}

function updateConnectionStatus(status, message) {
    const statusDot = document.querySelector('header .relative.flex.h-2\\.5');
    const statusText = document.querySelector('header .text-slate-300');
    const lastSyncText = document.getElementById('lastSyncTime');
    
    // Update AppState
    AppState.connection.status = status;
    
    // Update status dot
    if (statusDot) {
        const pingSpan = statusDot.querySelector('.animate-ping');
        const dotSpan = statusDot.querySelector('.relative.inline-flex');
        
        if (status === 'connected') {
            if (pingSpan) pingSpan.className = 'animate-ping absolute inline-flex h-full w-full rounded-full bg-emerald-400 opacity-75';
            if (dotSpan) dotSpan.className = 'relative inline-flex rounded-full h-2.5 w-2.5 bg-emerald-500';
            if (statusText) statusText.textContent = 'Connected';
        } else if (status === 'syncing') {
            if (pingSpan) pingSpan.className = 'animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75';
            if (dotSpan) dotSpan.className = 'relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500';
            if (statusText) statusText.textContent = 'Syncing...';
        } else if (status === 'error') {
            if (pingSpan) pingSpan.className = 'absolute inline-flex h-full w-full rounded-full bg-red-400 opacity-0';
            if (dotSpan) dotSpan.className = 'relative inline-flex rounded-full h-2.5 w-2.5 bg-red-500';
            if (statusText) statusText.textContent = 'Disconnected';
        }
    }
    
    // Update last sync text
    if (lastSyncText && message) {
        lastSyncText.textContent = message;
    }
    
    // Update footer status
    const footerStatus = document.querySelector('footer .flex.items-center.gap-2 span:last-child');
    const footerDot = document.querySelector('footer .w-2.h-2');
    if (footerStatus) {
        footerStatus.textContent = status === 'connected' ? 'System Ready' : 
                                   status === 'syncing' ? 'Syncing...' : 'Connection Error';
    }
    if (footerDot) {
        footerDot.className = `w-2 h-2 rounded-full ${
            status === 'connected' ? 'bg-emerald-500' : 
            status === 'syncing' ? 'bg-amber-500 animate-pulse' : 'bg-red-500'
        }`;
    }
}

// =========================================
// PANEL RESIZER LOGIC
// =========================================

function initResizer() {
    const resizer = document.getElementById('panelResizer');
    const leftPanel = document.getElementById('leftPanel');
    let isResizing = false;

    if (!resizer || !leftPanel) return;

    resizer.addEventListener('mousedown', function(e) {
        isResizing = true;
        document.body.style.cursor = 'col-resize';
        resizer.classList.add('resizing');
        
        // Add overlay to prevent iframe/select interference if any
        const overlay = document.createElement('div');
        overlay.id = 'resize-overlay';
        overlay.style.position = 'fixed';
        overlay.style.inset = '0';
        overlay.style.zIndex = '9999';
        overlay.style.cursor = 'col-resize';
        document.body.appendChild(overlay);
    });

    document.addEventListener('mousemove', function(e) {
        if (!isResizing) return;

        // Calculate new width
        let newWidth = e.clientX;
        
        // Boundaries
        if (newWidth < 300) newWidth = 300;
        if (newWidth > 800) newWidth = 800;

        leftPanel.style.width = `${newWidth}px`;
    });

    document.addEventListener('mouseup', function() {
        if (!isResizing) return;
        
        isResizing = false;
        document.body.style.cursor = 'default';
        resizer.classList.remove('resizing');
        
        const overlay = document.getElementById('resize-overlay');
        if (overlay) overlay.remove();
        
        // Trigger resize event for any charts/tables that need to re-adjust
        window.dispatchEvent(new Event('resize'));
    });
}

// =========================================
// MOCK DATA LOADER (Fallback for local dev)
// =========================================

function loadMockData() {
    // Mock Core Items
    AppState.coreItems = [
        { id: '430000-A557', description: 'Main Frame Assembly', quantity: 1 },
        { id: '430000-A558', description: 'Power Distribution Unit', quantity: 1 },
        { id: '430000-A559', description: 'Control Panel Assembly', quantity: 1 },
        { id: '430000-A560', description: 'Safety Enclosure', quantity: 1 },
        { id: '430000-A561', description: 'Base Platform', quantity: 1 },
    ];
    
    // Mock Config Items from REF_DATA!A:B
    AppState.refData.configItems = [
        { id: '430001-A378', description: 'Basic Tool Kit' },
        { id: '430001-A714', description: 'Pneumatic Kit' },
        { id: '430001-A712', description: 'Optional Module Package' },
        { id: '430001-A715', description: 'ESD Protection Kit' },
        { id: '430001-A716', description: 'Spare Parts Kit' },
        { id: '430001-A717', description: 'Calibration Standards Set' },
        { id: '430001-A718', description: 'Documentation Package' },
        { id: '430001-A719', description: 'Training Materials' },
        { id: '430001-A720', description: 'Maintenance Kit' },
        { id: '430001-A721', description: 'Software License Pack' },
    ];
    
    // Mock Module Items with mappings
    AppState.refData.moduleItems = [
        {
            id: '430001-A276',
            description: 'Rotary Module V2',
            mappings: {
                elecIds: '430001-A501;430001-A502;430001-A503',
                elecDesc: 'Connector Set A;Connector Set B;Connector Set C',
                toolIds: '430001-A602;430001-A689',
                toolDesc: 'Standard Tool Kit;Vacuum Pickup Assembly',
                jigIds: '430001-A801',
                jigDesc: 'Calibration Jig',
                visionIds: '430001-A756;430001-A757',
                visionDesc: 'B2K Camera Setup;S2K Camera Setup'
            }
        },
        {
            id: '430001-A277',
            description: 'Linear Transfer Module',
            mappings: {
                elecIds: '430001-A511;430001-A512',
                elecDesc: 'Linear Connector A;Linear Connector B',
                toolIds: '430001-A603',
                toolDesc: 'Linear Tool Kit',
                jigIds: '',
                jigDesc: '',
                visionIds: '430001-A758',
                visionDesc: 'Position Check Vision'
            }
        },
        {
            id: '430001-A278',
            description: 'Wafer Handler Module',
            mappings: {
                elecIds: '430001-A521',
                elecDesc: 'Handler Connector Set',
                toolIds: '430001-A604;430001-A605',
                toolDesc: 'Handler Tool Set;Gripper Kit',
                jigIds: '430001-A802;430001-A803',
                jigDesc: 'Alignment Jig;Test Jig',
                visionIds: ''
            }
        },
        {
            id: '430001-A279',
            description: 'Inspection Station',
            mappings: {
                elecIds: '430001-A531;430001-A532;430001-A533;430001-A534',
                elecDesc: 'Station Cable A;Station Cable B;Station Cable C;Station Cable D',
                toolIds: '430001-A606',
                toolDesc: 'Inspection Tool Kit',
                jigIds: '',
                jigDesc: '',
                visionIds: '430001-A759;430001-A760;430001-A761',
                visionDesc: 'Top Vision;Side Vision;Bottom Vision'
            }
        },
        {
            id: '430001-A280',
            description: 'Output Conveyor Module',
            mappings: {
                elecIds: '430001-A541',
                elecDesc: 'Conveyor Connector',
                toolIds: '',
                toolDesc: '',
                jigIds: '',
                jigDesc: '',
                visionIds: ''
            }
        }
    ];
    
    // Mock Vision Items
    AppState.refData.visionItems = [
        { id: '430001-A756', description: 'B2K Camera Setup - 4MP Basic', category: 'Position Check Vision' },
        { id: '430001-A757', description: 'S2K Camera Setup - 4MP Standard', category: 'Position Check Vision' },
        { id: '430001-A758', description: 'S4K Camera Setup - 16MP High Res', category: 'Position Check Vision' },
        { id: '430001-A759', description: 'Top View Camera Assembly', category: 'In-Pocket Vision' },
        { id: '430001-A760', description: 'Side View Camera Assembly', category: 'In-Pocket Vision' },
        { id: '430001-A761', description: 'Bottom View Camera Assembly', category: 'In-Pocket Vision' },
        { id: '430001-A762', description: '3D Scanner - BGA Package', category: '3D BGA Pad Package' },
        { id: '430001-A763', description: '3D Scanner - QFN Package', category: '3D BGA Pad Package' },
        { id: '430001-A764', description: '3D Scanner - Flip Chip', category: '3D BGA Pad Package' },
    ];
    
    // Mock Tooling Options
    AppState.refData.toolingOptions = {
        '430001-A602': [
            { id: '430001-A602-01', description: 'Tip Size 0.8mm', category: 'Standard Tips' },
            { id: '430001-A602-02', description: 'Tip Size 1.0mm', category: 'Standard Tips' },
            { id: '430001-A602-03', description: 'Tip Size 1.2mm', category: 'Standard Tips' },
            { id: '430001-A602-04', description: 'Tip Size 1.5mm', category: 'Large Tips' },
        ],
        '430001-A689': [
            { id: '430001-A689-01', description: 'Rubber Tip 0.8x0.8', category: 'Rubber Tips' },
            { id: '430001-A689-02', description: 'Rubber Tip 2x2', category: 'Rubber Tips' },
            { id: '430001-A689-03', description: 'Rubber Tip 3x3', category: 'Rubber Tips' },
            { id: '430001-A689-04', description: 'Solder Bump Tip', category: 'Special Tips' },
        ],
        '430001-A603': [
            { id: '430001-A603-01', description: 'Linear Guide A', category: 'Guides' },
            { id: '430001-A603-02', description: 'Linear Guide B', category: 'Guides' },
        ],
        '430001-A604': [
            { id: '430001-A604-01', description: 'Gripper Type A', category: 'Grippers' },
            { id: '430001-A604-02', description: 'Gripper Type B', category: 'Grippers' },
            { id: '430001-A604-03', description: 'Gripper Type C', category: 'Grippers' },
        ],
        '430001-A606': [
            { id: '430001-A606-01', description: 'Inspection Lens 50mm', category: 'Lenses' },
            { id: '430001-A606-02', description: 'Inspection Lens 65mm', category: 'Lenses' },
            { id: '430001-A606-03', description: 'Inspection Lens 110mm', category: 'Lenses' },
        ]
    };
    
    // Mock Shopping Lists
    AppState.refData.shoppingLists.basicTool = [
        { id: '430002-A001', description: 'Screwdriver Set (Metric)' },
        { id: '430002-A002', description: 'Hex Key Set' },
        { id: '430002-A003', description: 'Torque Wrench' },
        { id: '430002-A004', description: 'Precision Tweezers' },
        { id: '430002-A005', description: 'ESD Wrist Strap' },
        { id: '430002-A006', description: 'Multimeter' },
        { id: '430002-A007', description: 'Caliper (Digital)' },
        { id: '430002-A008', description: 'Flashlight (LED)' },
        { id: '430002-A009', description: 'Safety Glasses' },
        { id: '430002-A010', description: 'Cleaning Kit' },
        { id: '430002-A011', description: 'Anti-Static Mat' },
        { id: '430002-A012', description: 'Cable Ties Set' },
    ];
    
    AppState.refData.shoppingLists.pneumatic = [
        { id: '430002-A101', description: 'Air Fitting Set' },
        { id: '430002-A102', description: 'Pressure Regulator' },
        { id: '430002-A103', description: 'Quick Disconnect Set' },
        { id: '430002-A104', description: 'Air Line Tubing (5m)' },
        { id: '430002-A105', description: 'Filter/Regulator Unit' },
    ];
    
    // Initialize empty state arrays
    AppState.configItems = [];
    AppState.modules = [];
    AppState.visionItems = [];
    
    // Build initial order list
    rebuildOrderList();
    
    console.log('Mock data loaded successfully');
}

// =========================================
// RENDER ALL SECTIONS
// =========================================

function renderAllSections() {
    renderCoreSection();
    renderConfigSection();
    renderModuleSection();
    renderVisionSection();
    renderOrderList();
    updateSummary();
    updateConfigCount();
    updateModuleCount();
    updateVisionCount();
    
    // Re-init icons after render
    lucide.createIcons();
}

// =========================================
// BUTTON HANDLERS (Called from HTML)
// =========================================

function addModule() {
    addModuleSlot();
}

function addVision() {
    addVisionItem();
}

function addConfig() {
    addConfigItem();
}

// =========================================
// KEYBOARD SHORTCUTS
// =========================================

function setupKeyboardShortcuts() {
    document.addEventListener('keydown', function(e) {
        // Escape to close modals
        if (e.key === 'Escape') {
            if (AppState.ui.modals.password.open) {
                closePasswordModal();
            }
            if (AppState.ui.modals.override.context) {
                closeOverrideModal();
            }
            closeSettingsModal();
        }
        
        // Enter to confirm password
        if (e.key === 'Enter' && AppState.ui.modals.password.open) {
            confirmUncheck();
        }
        
        // Ctrl+S to trigger sync
        if (e.ctrlKey && e.key === 's') {
            e.preventDefault();
            triggerSync();
        }
        
        // Ctrl+R to refresh data
        if (e.ctrlKey && e.key === 'r' && !e.shiftKey) {
            e.preventDefault();
            refreshRefData();
        }
    });
}

// =========================================
// UTILITY: Update last sync display
// =========================================

function updateLastSyncDisplay() {
    const lastSync = AppState.connection.lastSync;
    const now = new Date();
    const diffMs = now - lastSync;
    const diffMins = Math.floor(diffMs / 60000);
    
    let text = 'Just now';
    if (diffMins >= 1 && diffMins < 60) {
        text = `${diffMins} min ago`;
    } else if (diffMins >= 60) {
        const diffHours = Math.floor(diffMins / 60);
        text = `${diffHours} hour${diffHours > 1 ? 's' : ''} ago`;
    }
    
    const el = document.getElementById('lastSyncTime');
    if (el) {
        el.textContent = `Last sync: ${text}`;
    }
}

// Update sync time display every minute
setInterval(updateLastSyncDisplay, 60000);

// =========================================
// ADMIN FUNCTIONS
// =========================================

/**
 * Trigger sync to refresh data from REF_DATA
 * This is the "quick sync" that reloads reference data
 */
async function triggerSync() {
    try {
        showToast('Syncing data...', 'info');
        
        if (GAS_API.isGASEnvironment()) {
            // Get fresh REF_DATA from GAS
            const result = await GAS_API.getRefData();
            
            if (!result.success) {
                showToast('Sync failed: ' + result.error.message, 'error');
                return;
            }
            
            // Update refData in AppState
            AppState.refData = result.data;
            
            // Re-render all sections with fresh data
            renderConfigSection();
            renderModuleSection();
            renderVisionSection();
            
            // Update last sync time
            document.getElementById('lastSyncTime').textContent = 'Last sync: Just now';
            
            showToast('Sync completed successfully', 'success');
        } else {
            showToast('Sync is only available in GAS environment', 'warning');
        }
    } catch (error) {
        console.error('Sync error:', error);
        showToast('Sync failed', 'error');
    }
}

/**
 * Trigger renumbering of electrical kits
 * Recalculates rotation based on instance counts
 */
async function triggerRenumber() {
    try {
        const confirmed = confirm('This will recalculate electrical kit assignments for all modules. Continue?');
        if (!confirmed) return;
        
        showToast('Renumbering kits...', 'info');
        
        if (GAS_API.isGASEnvironment()) {
            const result = await GAS_API.renumberKits();
            
            if (!result.success) {
                showToast('Renumber failed: ' + result.error.message, 'error');
                return;
            }
            
            const data = result.data;
            showToast(`Renumbering complete. ${data.updatedCount} kit(s) updated.`, 'success');
            
            // Reload the order list to show updated kits
            const stateResult = await GAS_API.getFullState();
            if (stateResult.success) {
                // Update modules in AppState
                AppState.modules = stateResult.data.modules || [];
                rebuildOrderList();
                renderModuleSection();
                renderOrderList();
            }
            
        } else {
            showToast('Renumber is only available in GAS environment', 'warning');
        }
    } catch (error) {
        console.error('Renumber error:', error);
        showToast('Renumber failed', 'error');
    }
}

/**
 * Load and display sync status
 */
async function loadSyncStatus() {
    try {
        if (GAS_API.isGASEnvironment()) {
            const result = await GAS_API.getSyncStatus();
            
            if (result.success) {
                const data = result.data;
                
                // Update last sync time
                if (data.lastSync) {
                    const lastSyncDate = new Date(data.lastSync);
                    const timeAgo = formatTimeAgo(lastSyncDate);
                    document.getElementById('lastSyncTime').textContent = `Last sync: ${timeAgo}`;
                }
                
                // Update spreadsheet ID display
                const spreadsheetIdEl = document.getElementById('spreadsheetId');
                if (spreadsheetIdEl) {
                    spreadsheetIdEl.textContent = data.spreadsheetId;
                }
            }
        }
    } catch (error) {
        console.error('Error loading sync status:', error);
    }
}

/**
 * Format time ago (e.g., "5 minutes ago", "2 hours ago")
 */
function formatTimeAgo(date) {
    const seconds = Math.floor((new Date() - date) / 1000);
    
    if (seconds < 60) return 'Just now';
    
    const minutes = Math.floor(seconds / 60);
    if (minutes < 60) return `${minutes} minute${minutes > 1 ? 's' : ''} ago`;
    
    const hours = Math.floor(minutes / 60);
    if (hours < 24) return `${hours} hour${hours > 1 ? 's' : ''} ago`;
    
    const days = Math.floor(hours / 24);
    return `${days} day${days > 1 ? 's' : ''} ago`;
}

// =========================================
// DEBUG HELPERS (Remove in production)
// =========================================

window.debugState = function() {
    console.log('Current AppState:', JSON.parse(JSON.stringify(AppState)));
};

window.debugAddTestModules = function() {
    // Add 3 modules for testing rotation
    addModuleSlot();
    updateModuleParent(0, '430001-A276');
    
    addModuleSlot();
    updateModuleParent(1, '430001-A276');
    
    addModuleSlot();
    updateModuleParent(2, '430001-A276');
    
    console.log('Added 3 test modules with same parent to test rotation');
};

window.debugAddTestConfig = function() {
    // Add test config items including special triggers
    addConfigItem();
    updateConfigSelection(0, '430001-A378');
    
    addConfigItem();
    updateConfigSelection(1, '430001-A714');
    
    addConfigItem();
    updateConfigSelection(2, '430001-A715');
    
    console.log('Added 3 test config items (2 special triggers + 1 regular)');
};

window.debugReloadFromGAS = async function() {
    console.log('Reloading data from GAS...');
    await initializeApp();
};

window.debugEnableMockMode = function() {
    GAS_API.config.useMockData = true;
    console.log('Mock mode enabled');
};

console.log('Debug helpers available: debugState(), debugAddTestModules(), debugAddTestConfig(), debugReloadFromGAS(), debugEnableMockMode()');
