/**
 * ViTrox BOM Configurator - UI Components
 * Render functions for each UI component
 */

// =========================================
// CONNECTION CHECK
// =========================================

/**
 * Check if editing is allowed (online mode)
 */
function canEdit() {
    return AppState.connection.status === 'connected';
}

/**
 * Show warning if offline
 */
function checkConnectionBeforeEdit(action) {
    if (!canEdit()) {
        showToast('Editing disabled while offline', 'warning');
        return false;
    }
    return true;
}

// =========================================
// SECTION TOGGLE
// =========================================

function toggleSection(sectionName) {
    const content = document.getElementById(`${sectionName}Content`);
    const chevron = document.getElementById(`${sectionName}Chevron`);
    
    if (content && chevron) {
        content.classList.toggle('hidden');
        chevron.classList.toggle('rotated');
        AppState.ui.expandedSections[sectionName] = !content.classList.contains('hidden');
    }
}

// =========================================
// CORE SECTION RENDERER
// =========================================

function renderCoreSection() {
    const container = document.getElementById('coreContent');
    if (!container) return;
    
    if (AppState.coreItems.length === 0) {
        container.innerHTML = `
            <div class="text-sm text-slate-400 text-center py-4">
                No core components loaded
            </div>
        `;
        return;
    }
    
    let html = '';
    AppState.coreItems.forEach(item => {
        html += `
            <div class="core-item">
                <div>
                    <p class="font-mono text-sm text-slate-600">${item.id}</p>
                    <p class="text-sm text-slate-800">${item.description}</p>
                </div>
                <span class="text-xs bg-slate-100 text-slate-500 px-2 py-1 rounded">Qty: ${item.quantity || 1}</span>
            </div>
        `;
    });
    
    container.innerHTML = html;
    updateCoreCount();
}

// =========================================
// CONFIG SECTION RENDERER (Flexible)
// =========================================

function renderConfigSection() {
    const container = document.getElementById('configContent');
    if (!container) return;
    
    if (AppState.configItems.length === 0) {
        container.innerHTML = `
            <div class="border-2 border-dashed border-slate-300 rounded-lg p-6 text-center">
                <i data-lucide="package-plus" class="w-8 h-8 text-slate-300 mx-auto mb-2"></i>
                <p class="text-sm text-slate-400">No config items added</p>
                <p class="text-xs text-slate-400 mt-1">Click "+ Add Config" to begin</p>
            </div>
        `;
        lucide.createIcons();
        updateConfigCount();
        return;
    }
    
    let html = '';
    AppState.configItems.forEach((configItem, idx) => {
        html += renderConfigItemCard(configItem, idx);
    });
    
    container.innerHTML = html;
    lucide.createIcons();
    updateConfigCount();
}

/**
 * Render a single config item card
 */
function renderConfigItemCard(configItem, slotIndex) {
    const availableOptions = getAvailableConfigOptions(slotIndex);
    const triggerConfig = configItem.selectedId ? getSpecialTriggerConfig(configItem.selectedId) : null;
    
    // Determine card styling based on special trigger or regular
    let cardClass = 'config-item-card';
    let bgClass = 'bg-white';
    let borderClass = 'border-slate-200';
    let iconBgClass = 'bg-slate-100';
    let iconContent = '<i data-lucide="package" class="w-5 h-5 text-slate-500"></i>';
    
    if (triggerConfig) {
        if (triggerConfig.bgColor === 'amber') {
            bgClass = 'bg-amber-50';
            borderClass = 'border-amber-200';
            iconBgClass = 'bg-amber-100';
            iconContent = `<span class="text-xl">${triggerConfig.icon}</span>`;
        } else if (triggerConfig.bgColor === 'sky') {
            bgClass = 'bg-sky-50';
            borderClass = 'border-sky-200';
            iconBgClass = 'bg-sky-100';
            iconContent = `<span class="text-xl">${triggerConfig.icon}</span>`;
        }
    }
    
    return `
        <div class="${cardClass} ${bgClass} border ${borderClass} rounded-lg overflow-hidden" data-slot="${slotIndex}">
            <!-- Card Header -->
            <div class="flex items-center justify-between p-3">
                <div class="flex items-center gap-3">
                    <div class="w-10 h-10 ${iconBgClass} rounded-lg flex items-center justify-center">
                        ${iconContent}
                    </div>
                    <div class="flex-1">
                        <p class="text-xs text-slate-500 uppercase tracking-wide mb-1">Config Item ${slotIndex + 1}</p>
                        <select class="w-full text-sm bg-white border border-slate-200 rounded px-2 py-1.5 custom-select min-w-[200px]"
                                onchange="updateConfigSelection(${slotIndex}, this.value)">
                            <option value="">Select config item...</option>
                            ${configItem.selectedId ? `<option value="${configItem.selectedId}" selected>${configItem.selectedId} - ${configItem.description}</option>` : ''}
                            ${availableOptions
                                .filter(opt => opt.id !== configItem.selectedId)
                                .map(opt => `
                                    <option value="${opt.id}">
                                        ${opt.id} - ${opt.description}
                                        ${isSpecialTrigger(opt.id) ? ' ★' : ''}
                                    </option>
                                `).join('')}
                        </select>
                    </div>
                </div>
                <button onclick="removeConfigItem(${slotIndex})" 
                        class="p-1.5 hover:bg-slate-200 rounded transition-colors ml-2" 
                        title="Remove Config Item">
                    <i data-lucide="x" class="w-4 h-4 text-slate-400"></i>
                </button>
            </div>
            
            <!-- Special Trigger Shopping List (if applicable) -->
            ${triggerConfig ? renderSpecialTriggerShoppingList(configItem, slotIndex, triggerConfig) : ''}
        </div>
    `;
}

/**
 * Render shopping list for special trigger config items
 */
function renderSpecialTriggerShoppingList(configItem, slotIndex, triggerConfig) {
    const shoppingListKey = triggerConfig.shoppingListKey;
    const shoppingList = AppState.refData.shoppingLists[shoppingListKey] || [];
    const selections = configItem.shoppingListSelections || [];
    const count = triggerConfig.shoppingListSize;
    
    let borderColor = triggerConfig.bgColor === 'amber' ? 'border-amber-200' : 'border-sky-200';
    let labelColor = triggerConfig.bgColor === 'amber' ? 'text-amber-700' : 'text-sky-700';
    
    let html = `
        <div class="border-t ${borderColor} p-3">
            <p class="text-xs ${labelColor} uppercase tracking-wide mb-2 font-medium">
                ${triggerConfig.icon} Shopping List (${count} items)
            </p>
            <div class="space-y-2">
    `;
    
    for (let i = 0; i < count; i++) {
        const selectedId = selections[i] || '';
        const selectedItem = shoppingList.find(o => o.id === selectedId);
        
        html += `
            <div class="flex items-center gap-2">
                <span class="text-xs text-slate-400 w-5 text-right">${i + 1}.</span>
                <select class="flex-1 text-sm bg-white border border-slate-200 rounded px-2 py-1.5 custom-select"
                        onchange="updateConfigShoppingSelection(${slotIndex}, ${i}, this.value)">
                    <option value="">Select item...</option>
                    ${shoppingList.map(opt => `
                        <option value="${opt.id}" ${opt.id === selectedId ? 'selected' : ''}>
                            ${opt.id} - ${opt.description}
                        </option>
                    `).join('')}
                </select>
            </div>
        `;
    }
    
    html += '</div></div>';
    return html;
}

// =========================================
// MODULE SECTION RENDERER
// =========================================

function renderModuleSection() {
    const container = document.getElementById('moduleContent');
    if (!container) return;
    
    if (AppState.modules.length === 0) {
        container.innerHTML = `
            <div class="border-2 border-dashed border-slate-300 rounded-lg p-6 text-center">
                <i data-lucide="plus-circle" class="w-8 h-8 text-slate-300 mx-auto mb-2"></i>
                <p class="text-sm text-slate-400">Click "+ Add Module" to begin</p>
            </div>
        `;
        lucide.createIcons();
        return;
    }
    
    let html = '';
    AppState.modules.forEach((module, idx) => {
        html += renderModuleCard(module, idx);
    });
    
    container.innerHTML = html;
    lucide.createIcons();
    updateModuleCount();
}

function renderModuleCard(module, slotIndex) {
    const instanceBadge = module.instanceTotal > 1 
        ? `<span class="instance-badge">Instance ${module.instanceNumber} of ${module.instanceTotal}</span>`
        : '';
    
    return `
        <div class="module-card ${module.parentId ? 'selected' : ''}" data-slot="${slotIndex}">
            <div class="module-card-header flex items-center justify-between">
                <div class="flex items-center gap-2">
                    <span class="text-sm font-semibold text-slate-600">Module ${slotIndex + 1}</span>
                    ${instanceBadge}
                </div>
                <button onclick="removeModuleSlot(${slotIndex})" class="p-1 hover:bg-slate-200 rounded transition-colors" title="Remove Module">
                    <i data-lucide="x" class="w-4 h-4 text-slate-400"></i>
                </button>
            </div>
            <div class="module-card-body">
                <!-- Parent Selector -->
                <div class="mb-4">
                    <label class="block text-xs font-medium text-slate-500 mb-1">Parent Module</label>
                    <select class="w-full text-sm bg-white border border-slate-200 rounded-lg px-3 py-2 custom-select"
                            onchange="updateModuleParent(${slotIndex}, this.value)">
                        <option value="">Select a module...</option>
                        ${AppState.refData.moduleItems.map(item => `
                            <option value="${item.id}" ${item.id === module.parentId ? 'selected' : ''}>
                                ${item.id} - ${item.description}
                            </option>
                        `).join('')}
                    </select>
                </div>
                
                ${module.parentId ? renderModuleChildren(module, slotIndex) : `
                    <div class="text-center py-4 text-slate-400 text-sm">
                        <i data-lucide="arrow-up" class="w-5 h-5 mx-auto mb-1"></i>
                        Select a parent module above
                    </div>
                `}
            </div>
        </div>
    `;
}

function renderModuleChildren(module, slotIndex) {
    let html = `<div class="space-y-3 border-t border-slate-100 pt-3 mt-3">
        <p class="text-xs text-slate-500 uppercase tracking-wide">Auto-Generated Children</p>`;
    
    // Electrical Kit (with Option D UI)
    if (module.children.electrical) {
        const elec = module.children.electrical;
        const isOverridden = elec.isOverridden;
        
        html += `
            <div class="child-item electrical">
                <div class="flex items-center justify-between mb-2">
                    <div class="flex items-center gap-2">
                        <i data-lucide="zap" class="w-4 h-4 text-amber-600"></i>
                        <span class="text-sm font-medium text-amber-800">Electrical Kit</span>
                    </div>
                    <span class="auto-badge">
                        ${isOverridden ? '⚠️ Overridden' : `Auto: ${elec.rotationIndex + 1} of ${elec.rotationTotal}`}
                    </span>
                </div>
                <select class="w-full text-sm bg-white border border-amber-200 rounded px-2 py-1.5 custom-select"
                        onchange="handleElectricalChange(${slotIndex}, this.value)">
                    ${elec.options.map(opt => `
                        <option value="${opt.id}" ${opt.id === elec.currentId ? 'selected' : ''}>
                            ${opt.id} - ${opt.description} ${opt.id === elec.autoSelectedId ? '✓ (Auto)' : ''}
                        </option>
                    `).join('')}
                </select>
                ${!isOverridden ? `
                    <p class="text-xs text-amber-700 mt-1 flex items-center gap-1">
                        <i data-lucide="info" class="w-3 h-3"></i>
                        Auto-selected based on rotation. Override if needed.
                    </p>
                ` : ''}
            </div>
        `;
    }
    
    // Tooling Kits
    if (module.children.tooling.length > 0) {
        module.children.tooling.forEach((tool, toolIdx) => {
            html += `
                <div class="child-item tooling">
                    <div class="flex items-center gap-2 mb-2">
                        <i data-lucide="wrench" class="w-4 h-4 text-emerald-600"></i>
                        <span class="text-sm font-medium text-emerald-800">${tool.description || 'Tooling Kit'}</span>
                    </div>
                    <p class="text-xs font-mono text-slate-500 mb-2">${tool.id}</p>
                    ${tool.optionChoices.length > 0 ? `
                        <div class="mt-2 pl-4 border-l-2 border-emerald-200">
                            <label class="block text-xs text-slate-500 mb-1">Select Option:</label>
                            <select class="w-full text-sm bg-white border border-emerald-200 rounded px-2 py-1.5 custom-select"
                                    onchange="updateToolingOption(${slotIndex}, ${toolIdx}, this.value)">
                                <option value="">Choose option...</option>
                                ${tool.optionChoices.map(opt => `
                                    <option value="${opt.id}" ${opt.id === tool.selectedOption ? 'selected' : ''}>
                                        ${opt.id} - ${opt.description}
                                    </option>
                                `).join('')}
                            </select>
                        </div>
                    ` : ''}
                </div>
            `;
        });
    }
    
    // Jigs
    if (module.children.jigs.length > 0) {
        module.children.jigs.forEach(jig => {
            html += `
                <div class="child-item jig">
                    <div class="flex items-center gap-2">
                        <i data-lucide="ruler" class="w-4 h-4 text-sky-600"></i>
                        <div>
                            <span class="text-sm font-medium text-sky-800">${jig.description || 'Jig'}</span>
                            <p class="text-xs font-mono text-slate-500">${jig.id}</p>
                        </div>
                    </div>
                </div>
            `;
        });
    }
    
    // Vision (within module)
    if (module.children.vision) {
        const vision = module.children.vision;
        
        html += `
            <div class="child-item vision">
                <div class="flex items-center gap-2 mb-2">
                    <i data-lucide="eye" class="w-4 h-4 text-violet-600"></i>
                    <span class="text-sm font-medium text-violet-800">Vision Layer</span>
                </div>
                ${vision.type === 'fixed' ? `
                    <div class="bg-white rounded px-3 py-2 border border-violet-200">
                        <p class="text-sm font-mono text-slate-700">${vision.selectedId}</p>
                        <p class="text-xs text-slate-500">${vision.description}</p>
                        <span class="text-xs bg-violet-100 text-violet-600 px-2 py-0.5 rounded mt-1 inline-block">${vision.category}</span>
                    </div>
                ` : `
                    <select class="w-full text-sm bg-white border border-violet-200 rounded px-2 py-1.5 custom-select"
                            onchange="updateModuleVision(${slotIndex}, this.value)">
                        <option value="">Select vision component...</option>
                        ${vision.options.map(opt => `
                            <option value="${opt.id}" ${opt.id === vision.selectedId ? 'selected' : ''}>
                                ${opt.id} - ${opt.description}
                            </option>
                        `).join('')}
                    </select>
                    ${vision.selectedId ? `
                        <span class="text-xs bg-violet-100 text-violet-600 px-2 py-0.5 rounded mt-2 inline-block">${vision.category}</span>
                    ` : ''}
                `}
            </div>
        `;
    }
    
    html += '</div>';
    return html;
}

// Handle electrical kit change with override warning
function handleElectricalChange(slotIndex, newKitId) {
    const module = AppState.modules[slotIndex];
    if (!module || !module.children.electrical) return;
    
    const elec = module.children.electrical;
    
    // Check if changing away from auto-selected
    if (newKitId !== elec.autoSelectedId && elec.currentId === elec.autoSelectedId) {
        // Show override warning
        AppState.ui.modals.override.context = { slotIndex, newKitId, oldKitId: elec.currentId };
        openOverrideModal(elec.currentId, newKitId);
    } else {
        // Direct update (reverting to auto or already overridden)
        overrideElectricalKit(slotIndex, newKitId);
    }
}

// =========================================
// VISION SECTION RENDERER
// =========================================

function renderVisionSection() {
    const container = document.getElementById('visionContent');
    if (!container) return;
    
    if (AppState.visionItems.length === 0) {
        container.innerHTML = `
            <div class="border-2 border-dashed border-slate-300 rounded-lg p-6 text-center">
                <i data-lucide="scan-eye" class="w-8 h-8 text-slate-300 mx-auto mb-2"></i>
                <p class="text-sm text-slate-400">No vision components added</p>
            </div>
        `;
        lucide.createIcons();
        return;
    }
    
    let html = '';
    AppState.visionItems.forEach((item, idx) => {
        html += `
            <div class="vision-item-card">
                <div class="flex items-center justify-between mb-2">
                    <span class="text-sm font-medium text-slate-600">Vision Item ${idx + 1}</span>
                    <button onclick="removeVisionItem(${idx})" class="p-1 hover:bg-slate-100 rounded transition-colors">
                        <i data-lucide="x" class="w-4 h-4 text-slate-400"></i>
                    </button>
                </div>
                <select class="w-full text-sm bg-white border border-slate-200 rounded-lg px-3 py-2 custom-select"
                        onchange="updateStandaloneVision(${idx}, this.value)">
                    <option value="">Select vision component...</option>
                    ${renderVisionOptionsGrouped(item.selectedId)}
                </select>
                ${item.selectedId ? `
                    <div class="mt-2 flex items-center justify-between">
                        <span class="text-xs text-slate-500">${item.description}</span>
                        <span class="text-xs bg-violet-100 text-violet-600 px-2 py-0.5 rounded">${item.category}</span>
                    </div>
                ` : ''}
            </div>
        `;
    });
    
    container.innerHTML = html;
    lucide.createIcons();
    updateVisionCount();
}

function renderVisionOptionsGrouped(selectedId) {
    // Group vision items by category
    const grouped = {};
    AppState.refData.visionItems.forEach(item => {
        const cat = item.category || 'Uncategorized';
        if (!grouped[cat]) grouped[cat] = [];
        grouped[cat].push(item);
    });
    
    let html = '';
    Object.entries(grouped).forEach(([category, items]) => {
        html += `<optgroup label="${category}">`;
        items.forEach(item => {
            html += `<option value="${item.id}" ${item.id === selectedId ? 'selected' : ''}>${item.id} - ${item.description}</option>`;
        });
        html += `</optgroup>`;
    });
    
    return html;
}

// =========================================
// ORDER LIST RENDERER
// =========================================

function renderOrderList() {
    const tbody = document.getElementById('orderTableBody');
    if (!tbody) return;
    
    const filteredList = getFilteredOrderList();
    
    if (filteredList.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="px-4 py-12 text-center text-slate-400">
                    <i data-lucide="inbox" class="w-12 h-12 mx-auto mb-3 text-slate-300"></i>
                    <p>No items match your filters</p>
                </td>
            </tr>
        `;
        lucide.createIcons();
        return;
    }
    
    let html = '';
    let currentSection = '';
    
    filteredList.forEach(item => {
        // Section header
        if (item.section !== currentSection) {
            currentSection = item.section;
            html += `
                <tr class="order-row section-header">
                    <td colspan="7" class="px-4 py-2 text-xs font-bold uppercase tracking-wider text-slate-500">
                        ${getSectionIcon(item.section)} ${item.section}
                    </td>
                </tr>
            `;
        }
        
        const rowClass = item.isChecked ? 'checked' : '';
        const depthClass = item.depth === 1 ? 'child-row' : item.depth === 2 ? 'grandchild-row' : '';
        const indentStyle = item.depth > 0 ? `padding-left: ${16 + (item.depth * 20)}px` : '';
        
        html += `
            <tr class="order-row ${rowClass} ${depthClass}" data-line="${item.lineNumber}">
                <td class="px-4 py-3 text-sm text-slate-400">${item.lineNumber}</td>
                <td class="px-4 py-3 part-id" style="${indentStyle}">
                    ${item.depth > 0 ? '<span class="text-slate-300 mr-1">└</span>' : ''}
                    <span class="font-mono text-sm ${item.isChecked ? 'text-emerald-600' : 'text-slate-700'}">${item.partId}</span>
                </td>
                <td class="px-4 py-3 text-sm text-slate-600">${item.description}</td>
                <td class="px-4 py-3 text-center text-sm text-slate-500">${item.quantity}</td>
                <td class="px-4 py-3 text-center">
                    <div class="custom-checkbox ${item.isChecked ? 'checked' : ''} mx-auto" 
                         onclick="toggleOrderCheckbox(${item.lineNumber})"></div>
                </td>
                <td class="px-4 py-3 text-sm text-slate-500">${item.checkDate || '—'}</td>
                <td class="px-4 py-3">
                    <select class="text-xs bg-slate-50 border border-slate-200 rounded px-2 py-1 ${!item.isChecked ? 'opacity-50' : ''}"
                            ${!item.isChecked ? 'disabled' : ''}
                            onchange="updateReleaseType(${item.lineNumber}, this.value)">
                        <option value="">—</option>
                        <option value="CHARGE OUT" ${item.releaseType === 'CHARGE OUT' ? 'selected' : ''}>CHARGE OUT</option>
                        <option value="MRP" ${item.releaseType === 'MRP' ? 'selected' : ''}>MRP</option>
                    </select>
                </td>
            </tr>
        `;
    });
    
    tbody.innerHTML = html;
    lucide.createIcons();
}

function getSectionIcon(section) {
    const icons = {
        'CORE': '📦',
        'CONFIG': '⚙️',
        'MODULE': '🔧',
        'VISION': '👁'
    };
    return icons[section] || '';
}

function getFilteredOrderList() {
    const { search, section, status } = AppState.ui.filters;
    
    return AppState.orderList.filter(item => {
        // Search filter
        if (search) {
            const searchLower = search.toLowerCase();
            if (!item.partId.toLowerCase().includes(searchLower) && 
                !item.description.toLowerCase().includes(searchLower)) {
                return false;
            }
        }
        
        // Section filter
        if (section !== 'all' && item.section !== section) {
            return false;
        }
        
        // Status filter
        if (status === 'checked' && !item.isChecked) return false;
        if (status === 'pending' && item.isChecked) return false;
        
        return true;
    });
}

function filterOrderList() {
    AppState.ui.filters.search = document.getElementById('searchInput')?.value || '';
    AppState.ui.filters.section = document.getElementById('sectionFilter')?.value || 'all';
    AppState.ui.filters.status = document.getElementById('statusFilter')?.value || 'all';
    
    renderOrderList();
}

// =========================================
// MODAL FUNCTIONS
// =========================================

function openPasswordModal(lineNumber) {
    const item = AppState.orderList.find(i => i.lineNumber === lineNumber);
    if (!item) return;
    
    document.getElementById('uncheckItemName').textContent = `${item.partId} - ${item.description}`;
    document.getElementById('uncheckItemDate').textContent = item.checkDate || 'Unknown';
    document.getElementById('passwordInput').value = '';
    
    AppState.ui.modals.password.open = true;
    AppState.ui.modals.password.targetLine = lineNumber;
    
    document.getElementById('passwordModal').classList.remove('hidden');
    document.getElementById('passwordInput').focus();
}

function closePasswordModal() {
    AppState.ui.modals.password.open = false;
    AppState.ui.modals.password.targetLine = null;
    document.getElementById('passwordModal').classList.add('hidden');
}

async function confirmUncheck() {
    const password = document.getElementById('passwordInput').value;
    
    if (!password) {
        showToast('Please enter a password', 'warning');
        return;
    }
    
    // Call the async confirmUncheckItem with password
    // The function will handle password validation via API
    await confirmUncheckItem(AppState.ui.modals.password.targetLine, password);
}

function openSettingsModal() {
    document.getElementById('settingsModal').classList.remove('hidden');
}

function closeSettingsModal() {
    document.getElementById('settingsModal').classList.add('hidden');
}

function openOverrideModal(currentKit, newKit) {
    document.getElementById('overrideCurrentKit').textContent = currentKit;
    document.getElementById('overrideNewKit').textContent = newKit;
    document.getElementById('overrideModal').classList.remove('hidden');
}

function closeOverrideModal() {
    AppState.ui.modals.override.context = null;
    document.getElementById('overrideModal').classList.add('hidden');
}

function confirmOverride() {
    const ctx = AppState.ui.modals.override.context;
    if (ctx) {
        overrideElectricalKit(ctx.slotIndex, ctx.newKitId);
    }
    closeOverrideModal();
}

// =========================================
// TOAST NOTIFICATIONS
// =========================================

function showToast(message, type = 'success') {
    const toast = document.getElementById('toast');
    const toastMessage = document.getElementById('toastMessage');
    const toastIcon = document.getElementById('toastIcon');
    
    toastMessage.textContent = message;
    
    // Update icon based on type
    const iconMap = {
        success: 'check-circle',
        error: 'x-circle',
        warning: 'alert-triangle',
        info: 'info'
    };
    const colorMap = {
        success: 'text-emerald-400',
        error: 'text-red-400',
        warning: 'text-amber-400',
        info: 'text-blue-400'
    };
    
    toastIcon.setAttribute('data-lucide', iconMap[type] || 'check-circle');
    toastIcon.className = `w-5 h-5 ${colorMap[type] || 'text-emerald-400'}`;
    lucide.createIcons();
    
    toast.classList.remove('hidden', 'hiding');
    
    // Auto-hide after 3 seconds
    setTimeout(() => {
        toast.classList.add('hiding');
        setTimeout(() => {
            toast.classList.add('hidden');
            toast.classList.remove('hiding');
        }, 300);
    }, 3000);
}

// =========================================
// PDF EXPORT
// =========================================

function exportToPDF() {
    showToast('Generating PDF...', 'info');
    
    try {
        // Create a temporary container for PDF content
        const pdfContainer = createPDFContent();
        
        // Add to document temporarily (hidden)
        pdfContainer.style.position = 'absolute';
        pdfContainer.style.left = '-9999px';
        pdfContainer.style.top = '0';
        document.body.appendChild(pdfContainer);
        
        // Generate filename with timestamp
        const now = new Date();
        const timestamp = `${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}_${String(now.getHours()).padStart(2,'0')}${String(now.getMinutes()).padStart(2,'0')}`;
        const filename = `BOM_OrderList_${timestamp}.pdf`;
        
        // PDF options
        const opt = {
            margin: [10, 10, 10, 10],
            filename: filename,
            image: { type: 'jpeg', quality: 0.98 },
            html2canvas: { 
                scale: 2,
                useCORS: true,
                logging: false
            },
            jsPDF: { 
                unit: 'mm', 
                format: 'a4', 
                orientation: 'landscape',
                compress: true
            },
            pagebreak: { mode: ['avoid-all', 'css', 'legacy'] }
        };
        
        // Generate and save PDF
        html2pdf().set(opt).from(pdfContainer).save().then(() => {
            // Clean up
            document.body.removeChild(pdfContainer);
            showToast('PDF exported successfully', 'success');
        }).catch(err => {
            // Clean up on error
            if (document.body.contains(pdfContainer)) {
                document.body.removeChild(pdfContainer);
            }
            showToast('Failed to export PDF', 'error');
            console.error('PDF Export Error:', err);
        });
        
    } catch (error) {
        showToast('Failed to generate PDF', 'error');
        console.error('PDF Generation Error:', error);
    }
}

/**
 * Create formatted PDF content with header and styling
 */
function createPDFContent() {
    const container = document.createElement('div');
    container.style.width = '277mm'; // A4 landscape width
    container.style.padding = '10mm';
    container.style.backgroundColor = 'white';
    container.style.fontFamily = 'Arial, sans-serif';
    
    // Calculate summary statistics
    const totalItems = AppState.orderList.length;
    const checkedItems = AppState.orderList.filter(i => i.isChecked).length;
    const uncheckedItems = totalItems - checkedItems;
    
    // PDF Header
    const header = document.createElement('div');
    header.style.marginBottom = '15px';
    header.style.borderBottom = '2px solid #1e293b';
    header.style.paddingBottom = '10px';
    header.innerHTML = `
        <div style="display: flex; justify-content: space-between; align-items: flex-start;">
            <div>
                <h1 style="margin: 0; font-size: 24px; color: #1e293b; font-weight: bold;">
                    BOM Order List
                </h1>
                <p style="margin: 5px 0 0 0; font-size: 12px; color: #64748b;">
                    Bill of Materials Configuration
                </p>
            </div>
            <div style="text-align: right;">
                <p style="margin: 0; font-size: 11px; color: #64748b;">
                    Generated: ${formatDateTime(new Date())}
                </p>
                <p style="margin: 3px 0 0 0; font-size: 11px; color: #64748b;">
                    Total Items: ${totalItems} | Checked: ${checkedItems} | Unchecked: ${uncheckedItems}
                </p>
            </div>
        </div>
    `;
    container.appendChild(header);
    
    // Section Summary
    const sectionSummary = createSectionSummary();
    if (sectionSummary) {
        container.appendChild(sectionSummary);
    }
    
    // Order List Table
    const table = createPDFTable();
    container.appendChild(table);
    
    // Footer
    const footer = document.createElement('div');
    footer.style.marginTop = '15px';
    footer.style.paddingTop = '10px';
    footer.style.borderTop = '1px solid #e2e8f0';
    footer.style.fontSize = '9px';
    footer.style.color = '#94a3b8';
    footer.style.textAlign = 'center';
    footer.innerHTML = `
        <p style="margin: 0;">
            BOM Configurator System | Export Date: ${formatDate(new Date())}
        </p>
    `;
    container.appendChild(footer);
    
    return container;
}

/**
 * Create section summary box
 */
function createSectionSummary() {
    const sections = {
        'CORE': AppState.coreItems.length,
        'CONFIG': AppState.configItems.filter(c => c.selectedId).length,
        'MODULE': AppState.modules.filter(m => m.parentId).length,
        'VISION': AppState.visionItems.filter(v => v.selectedId).length
    };
    
    const hasSections = Object.values(sections).some(count => count > 0);
    if (!hasSections) return null;
    
    const summaryBox = document.createElement('div');
    summaryBox.style.marginBottom = '15px';
    summaryBox.style.padding = '10px';
    summaryBox.style.backgroundColor = '#f8fafc';
    summaryBox.style.borderRadius = '4px';
    summaryBox.style.border = '1px solid #e2e8f0';
    
    let summaryHTML = '<div style="display: flex; gap: 20px; font-size: 11px;">';
    
    for (const [section, count] of Object.entries(sections)) {
        if (count > 0) {
            const icon = section === 'CORE' ? '📦' : section === 'CONFIG' ? '⚙️' : section === 'MODULE' ? '🔧' : '👁';
            summaryHTML += `
                <div style="flex: 1;">
                    <span style="font-weight: bold; color: #475569;">${icon} ${section}:</span>
                    <span style="color: #64748b;">${count} item${count !== 1 ? 's' : ''}</span>
                </div>
            `;
        }
    }
    
    summaryHTML += '</div>';
    summaryBox.innerHTML = summaryHTML;
    
    return summaryBox;
}

/**
 * Create formatted table for PDF
 */
function createPDFTable() {
    const tableWrapper = document.createElement('div');
    tableWrapper.style.fontSize = '9px';
    
    const table = document.createElement('table');
    table.style.width = '100%';
    table.style.borderCollapse = 'collapse';
    table.style.marginTop = '10px';
    
    // Table Header
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr style="background-color: #334155; color: white;">
            <th style="padding: 6px 8px; text-align: left; border: 1px solid #475569; width: 5%;">#</th>
            <th style="padding: 6px 8px; text-align: left; border: 1px solid #475569; width: 15%;">Part ID</th>
            <th style="padding: 6px 8px; text-align: left; border: 1px solid #475569; width: 40%;">Description</th>
            <th style="padding: 6px 8px; text-align: center; border: 1px solid #475569; width: 8%;">Qty</th>
            <th style="padding: 6px 8px; text-align: center; border: 1px solid #475569; width: 8%;">Checked</th>
            <th style="padding: 6px 8px; text-align: center; border: 1px solid #475569; width: 12%;">Date</th>
            <th style="padding: 6px 8px; text-align: center; border: 1px solid #475569; width: 12%;">Release Type</th>
        </tr>
    `;
    table.appendChild(thead);
    
    // Table Body
    const tbody = document.createElement('tbody');
    const filteredList = getFilteredOrderList();
    
    let currentSection = '';
    
    filteredList.forEach((item, index) => {
        // Section header row
        if (item.section !== currentSection) {
            currentSection = item.section;
            const sectionRow = document.createElement('tr');
            sectionRow.style.backgroundColor = '#f1f5f9';
            sectionRow.innerHTML = `
                <td colspan="7" style="padding: 5px 8px; border: 1px solid #cbd5e1; font-weight: bold; color: #475569; font-size: 10px;">
                    ${getSectionIcon(item.section)} ${item.section}
                </td>
            `;
            tbody.appendChild(sectionRow);
        }
        
        // Item row
        const row = document.createElement('tr');
        const bgColor = item.isChecked ? '#ecfdf5' : (index % 2 === 0 ? '#ffffff' : '#f8fafc');
        row.style.backgroundColor = bgColor;
        
        const indent = item.depth > 0 ? `padding-left: ${8 + (item.depth * 10)}px` : '';
        const depthMarker = item.depth > 0 ? '└ ' : '';
        const checkedIcon = item.isChecked ? '✓' : '—';
        const checkedColor = item.isChecked ? '#059669' : '#cbd5e1';
        
        row.innerHTML = `
            <td style="padding: 4px 8px; border: 1px solid #e2e8f0; color: #94a3b8;">${item.lineNumber}</td>
            <td style="padding: 4px 8px; border: 1px solid #e2e8f0; font-family: monospace; ${indent}">
                ${depthMarker}${item.partId}
            </td>
            <td style="padding: 4px 8px; border: 1px solid #e2e8f0; color: #475569;">${item.description}</td>
            <td style="padding: 4px 8px; border: 1px solid #e2e8f0; text-align: center; color: #64748b;">${item.quantity}</td>
            <td style="padding: 4px 8px; border: 1px solid #e2e8f0; text-align: center; color: ${checkedColor}; font-weight: bold;">${checkedIcon}</td>
            <td style="padding: 4px 8px; border: 1px solid #e2e8f0; text-align: center; color: #64748b;">${item.checkDate || '—'}</td>
            <td style="padding: 4px 8px; border: 1px solid #e2e8f0; text-align: center; color: #64748b; font-size: 8px;">${item.releaseType || '—'}</td>
        `;
        
        tbody.appendChild(row);
    });
    
    table.appendChild(tbody);
    tableWrapper.appendChild(table);
    
    return tableWrapper;
}

/**
 * Format date and time for display
 */
function formatDateTime(date) {
    const d = new Date(date);
    const dateStr = formatDate(d);
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    return `${dateStr} ${hours}:${minutes}`;
}

// =========================================
// SYNC & ADMIN FUNCTIONS
// =========================================

function triggerSync() {
    showToast('Syncing data...', 'info');
    
    // TODO: Implement actual sync with Google Sheets
    setTimeout(() => {
        AppState.connection.lastSync = new Date();
        updateLastSyncTime();
        showToast('Sync completed', 'success');
    }, 1500);
}

function triggerRenumber() {
    showToast('Renumbering kits...', 'info');
    
    // Recalculate all instance numbers
    recalculateInstanceNumbers();
    
    // Update electrical kits based on new rotation
    AppState.modules.forEach((module, idx) => {
        if (module.parentId && module.children.electrical) {
            const parentConfig = AppState.refData.moduleItems.find(item => item.id === module.parentId);
            if (parentConfig) {
                populateModuleChildren(module, parentConfig);
            }
        }
    });
    
    rebuildOrderList();
    renderModuleSection();
    renderOrderList();
    
    setTimeout(() => {
        showToast('Renumbering complete', 'success');
    }, 500);
}

