# Backend Integration Plan: ViTrox BOM Configurator

## Executive Summary

This document outlines the complete plan to integrate the ViTrox BOM Configurator frontend with Google Sheets backend via Google Apps Script (GAS). The frontend will mirror the exact business logic currently implemented in `ORDERING_LIST.gs`.

---

## Table of Contents

1. [Architecture Overview](#1-architecture-overview)
2. [Data Sources & Structure](#2-data-sources--structure)
3. [API Design](#3-api-design)
4. [Business Logic Replication](#4-business-logic-replication)
5. [Implementation Phases](#5-implementation-phases)
6. [File Structure](#6-file-structure)
7. [Error Handling](#7-error-handling)
8. [Testing Strategy](#8-testing-strategy)

---

## 1. Architecture Overview

### 1.1 System Diagram

```
┌─────────────────────────────────────────────────────────────────────────┐
│                           FRONTEND (Web App)                             │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐ │
│  │    CONFIG    │  │    MODULE    │  │    VISION    │  │  ORDER LIST  │ │
│  │   Section    │  │   Section    │  │   Section    │  │   Display    │ │
│  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘  └──────┬───────┘ │
│         │                 │                 │                 │         │
│         └─────────────────┴─────────────────┴─────────────────┘         │
│                                    │                                     │
│                            google.script.run                             │
└────────────────────────────────────┼─────────────────────────────────────┘
                                     │
                                     ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                    GOOGLE APPS SCRIPT (Backend API)                      │
│  ┌──────────────────────────────────────────────────────────────────┐   │
│  │                         api.gs (NEW)                              │   │
│  │  - getFullState()           - addModule()                        │   │
│  │  - getRefData()             - removeModule()                     │   │
│  │  - addConfigItem()          - updateModuleChild()                │   │
│  │  - removeConfigItem()       - addVisionItem()                    │   │
│  │  - updateConfigShopping()   - removeVisionItem()                 │   │
│  │  - renumberKits()           - toggleCheckbox()                   │   │
│  │  - triggerMasterSync()      - validatePassword()                 │   │
│  └──────────────────────────────────────────────────────────────────┘   │
│                                    │                                     │
│  ┌─────────────────────┐  ┌───────┴────────┐  ┌─────────────────────┐   │
│  │  ORDERING_LIST.gs   │  │   REF_DATA     │  │  handleCheckBox.gs  │   │
│  │  (Existing Logic)   │  │   (Database)   │  │  (Password Logic)   │   │
│  └─────────────────────┘  └────────────────┘  └─────────────────────┘   │
└─────────────────────────────────────────────────────────────────────────┘
                                     │
                                     ▼
┌─────────────────────────────────────────────────────────────────────────┐
│                         GOOGLE SPREADSHEETS                              │
│  ┌────────────────────────────────────────────────────────────────┐     │
│  │  Layout Configuration - 430000-LXXX-X (ID: 1a4fDQd1U7E650...)  │     │
│  │  ├── ORDERING LIST (User Data)                                 │     │
│  │  └── REF_DATA (Reference Database)                             │     │
│  └────────────────────────────────────────────────────────────────┘     │
│                                                                          │
│  ┌────────────────────────────────────────────────────────────────┐     │
│  │  BOM Structure Tree Diagram (ID: 1nTSOqK4nGRkUEHGFnUF30...)    │     │
│  │  ├── BOM Structure Tree Diagram (Source Data)                  │     │
│  │  └── Tooling Illustration (Tooling Options)                    │     │
│  └────────────────────────────────────────────────────────────────┘     │
└─────────────────────────────────────────────────────────────────────────┘
```

### 1.2 Key Design Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Deployment | Web App (standalone) | User accesses via URL, not tied to Sheets UI |
| Execute As | Me (Script Owner) | Ensures data access permissions |
| Access | Anyone (no login) | External users without Google accounts |
| Sync Approach | API-driven (Approach C) | Server-side logic, reliable state management |
| State Loading | Re-validate against REF_DATA | Catches invalid/outdated entries |
| Offline Mode | Block editing | Online-only application |

### 1.3 Spreadsheet IDs

| Spreadsheet | ID | Purpose |
|-------------|-----|---------|
| Main (ORDERING LIST + REF_DATA) | `1a4fDQd1U7E650gdxzEIIoQIDv2n38gP735BEbB86mT8` | User data and reference database |
| Source BOM | `1nTSOqK4nGRkUEHGFnUF30gRCGFQMo6I2l8vhZB-NkSA` | Master sync source |

---

## 2. Data Sources & Structure

### 2.1 ORDERING LIST Sheet Structure

| Column | Header | Content | Editable |
|--------|--------|---------|----------|
| A | Section | CORE, CONFIG, MODULE, VISION | No |
| B | Category | Category for VISION, "PART #:" for others | No |
| C | ITEM | Line number within section | No |
| D | Part ID | The actual part number | Yes (via dropdowns) |
| E | DESCRIPTION | Part description | Auto (formula/lookup) |
| F | QTY | Quantity | Yes |
| G | RELEASED | TRUE/FALSE checkbox | Yes (with password for uncheck) |
| H | DATE RELEASED | Date string | Auto (on check) |
| I | RELEASE TYPE | CHARGE OUT / MRP | Yes (after check) |
| J | REMARK | Notes | Yes |

### 2.2 REF_DATA Sheet Structure

| Columns | Purpose | Used By |
|---------|---------|---------|
| A:B | Config Items (Part ID, Description) | CONFIG section dropdowns |
| C:D | Module Items (Part ID, Description) | MODULE section dropdowns |
| I:J | Shopping List for 430001-A378 | Basic Tool Kit (10 items) |
| K:L | Shopping List for 430001-A714 | Pneumatic Kit (3 items) |
| P:Q | Tooling Options Database (Parent ID, Child ID) | Deletion checks |
| R:S | Tooling Options (Category, Description) | Option lookups |
| U:V | Shadow Menu (Parent ID, Display Value) | Dropdown rendering with categories |
| W:X | Electrical Mappings (IDs; Descriptions) | Module children |
| Y:Z | Tooling Mappings (IDs; Descriptions) | Module children |
| AA:AB | Jig Mappings (IDs; Descriptions) | Module children |
| AC:AD | Vision Mappings (IDs; Descriptions) | Module children |
| AF:AH | Vision Database (Part ID, Description, Category) | VISION section |
| AJ:AK | Vision Shadow Menu | Optional reference |

### 2.3 Section Boundaries

```
ORDERING LIST Structure:
┌─────────────────────────────────────┐
│ Row 1-2: Headers (PC Section)       │ ← Ignored by frontend
├─────────────────────────────────────┤
│ CORE Section                        │ ← Read-only (30 items)
│ - Fixed items from runMasterSync()  │
├─────────────────────────────────────┤
│ CONFIG Section                      │ ← Editable (10 slots)
│ - Config items with shopping lists  │
├─────────────────────────────────────┤
│ MODULE Section                      │ ← Editable (10 slots)
│ - Parent modules with children      │
│ - Children: Elec, Tool, Jig, Vision │
├─────────────────────────────────────┤
│ VISION Section                      │ ← Editable (10 slots)
│ - Standalone vision components      │
├─────────────────────────────────────┤
│ TOOLING, VCM, OTHERS, etc.          │ ← Ignored by frontend
└─────────────────────────────────────┘
```

---

## 3. API Design

### 3.1 API Functions Overview

```javascript
// ═══════════════════════════════════════════════════════════════
// READ OPERATIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Get complete application state on startup
 * @returns {Object} { refData, coreItems, configItems, modules, visionItems, orderList }
 */
function getFullState()

/**
 * Get reference data only (for refresh)
 * @returns {Object} { configOptions, moduleOptions, visionOptions, toolingOptions, shoppingLists }
 */
function getRefData()

// ═══════════════════════════════════════════════════════════════
// CONFIG OPERATIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Add a config item to a slot
 * @param {number} slotIndex - 0-based slot index
 * @param {string} partId - Part ID to add
 * @returns {Object} { success, configItem, shoppingList?, error? }
 */
function addConfigItem(slotIndex, partId)

/**
 * Remove a config item from a slot
 * @param {number} slotIndex - 0-based slot index
 * @returns {Object} { success, error? }
 */
function removeConfigItem(slotIndex)

/**
 * Update shopping list selection for special config triggers
 * @param {number} slotIndex - Config slot index
 * @param {number} shoppingIndex - Shopping list item index
 * @param {string} partId - Selected part ID
 * @returns {Object} { success, error? }
 */
function updateConfigShoppingList(slotIndex, shoppingIndex, partId)

// ═══════════════════════════════════════════════════════════════
// MODULE OPERATIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Add a module to a slot (triggers child insertion)
 * @param {number} slotIndex - 0-based slot index
 * @param {string} parentId - Parent module Part ID
 * @returns {Object} { success, module: { parentId, children: { electrical, tooling[], jigs[], vision } }, error? }
 */
function addModule(slotIndex, parentId)

/**
 * Remove a module from a slot (removes all children)
 * @param {number} slotIndex - 0-based slot index
 * @returns {Object} { success, error? }
 */
function removeModule(slotIndex)

/**
 * Update a module child selection (tooling option, vision selection, etc.)
 * @param {number} slotIndex - Module slot index
 * @param {string} childType - 'toolingOption' | 'rubberTip' | 'vision'
 * @param {number} childIndex - Index within child type
 * @param {string} partId - Selected part ID
 * @returns {Object} { success, error? }
 */
function updateModuleChild(slotIndex, childType, childIndex, partId)

/**
 * Override electrical kit (user manual override)
 * @param {number} slotIndex - Module slot index
 * @param {string} electricalId - New electrical kit Part ID
 * @returns {Object} { success, error? }
 */
function overrideElectricalKit(slotIndex, electricalId)

// ═══════════════════════════════════════════════════════════════
// VISION OPERATIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Add a standalone vision item
 * @param {number} slotIndex - 0-based slot index
 * @param {string} visionId - Vision Part ID
 * @returns {Object} { success, visionItem: { id, description, category }, error? }
 */
function addVisionItem(slotIndex, visionId)

/**
 * Remove a standalone vision item
 * @param {number} slotIndex - 0-based slot index
 * @returns {Object} { success, error? }
 */
function removeVisionItem(slotIndex)

/**
 * Update a standalone vision item
 * @param {number} slotIndex - 0-based slot index
 * @param {string} visionId - New Vision Part ID
 * @returns {Object} { success, visionItem: { id, description, category }, error? }
 */
function updateVisionItem(slotIndex, visionId)

// ═══════════════════════════════════════════════════════════════
// CHECKBOX OPERATIONS (Phase 2)
// ═══════════════════════════════════════════════════════════════

/**
 * Toggle checkbox (check an item)
 * @param {number} lineNumber - Line number in order list
 * @returns {Object} { success, timestamp, error? }
 */
function checkItem(lineNumber)

/**
 * Uncheck an item (requires password)
 * @param {number} lineNumber - Line number in order list
 * @param {string} password - User entered password
 * @returns {Object} { success, error? }
 */
function uncheckItem(lineNumber, password)

/**
 * Update release type for a checked item
 * @param {number} lineNumber - Line number in order list
 * @param {string} releaseType - 'CHARGE OUT' | 'MRP'
 * @returns {Object} { success, error? }
 */
function updateReleaseType(lineNumber, releaseType)

// ═══════════════════════════════════════════════════════════════
// ADMIN OPERATIONS
// ═══════════════════════════════════════════════════════════════

/**
 * Renumber electrical kits based on rotation
 * @returns {Object} { success, updatedModules: [], error? }
 */
function renumberKits()

/**
 * Trigger full master sync from external BOM
 * @returns {Object} { success, message, error? }
 */
function triggerMasterSync()

/**
 * Validate password for uncheck operation
 * @param {string} password - User entered password
 * @returns {Object} { valid: boolean }
 */
function validatePassword(password)

/**
 * Get password from Script Properties (admin only)
 * @returns {string} Current password
 */
function getUncheckPassword()

/**
 * Set password in Script Properties (admin only)
 * @param {string} newPassword - New password
 * @returns {Object} { success }
 */
function setUncheckPassword(newPassword)
```

### 3.2 Response Formats

#### Success Response
```javascript
{
  success: true,
  data: { /* operation-specific data */ },
  timestamp: "2025-01-02T10:30:00.000Z"
}
```

#### Error Response
```javascript
{
  success: false,
  error: {
    code: "INVALID_PART_ID",
    message: "Part ID 430000-XXXX not found in reference data",
    details: { partId: "430000-XXXX", section: "MODULE" }
  }
}
```

### 3.3 Full State Response Structure

```javascript
{
  success: true,
  data: {
    // Reference Data (for dropdowns)
    refData: {
      configItems: [
        { id: "430001-A714", description: "List-Optional Pneumatic Module" },
        { id: "430001-A378", description: "List-Optional Basic Tool Module" },
        // ...
      ],
      moduleItems: [
        { 
          id: "430000-A960", 
          description: "Module-Reject Bin with check present",
          mappings: {
            elecIds: "430001-A529;430001-A530;430001-A531;430001-A532",
            elecDesc: "Kit-Misc. Ele. Reject Bin 1;Kit-Misc. Ele. Reject Bin 2;...",
            toolIds: "",
            toolDesc: "",
            jigIds: "",
            jigDesc: "",
            visionIds: "",
            visionDesc: ""
          }
        },
        // ...
      ],
      visionItems: [
        { id: "430000-A756", description: "Position Check Vision 1", category: "Vision Fixed Spec" },
        // ...
      ],
      shoppingLists: {
        basicTool: [ // For 430001-A378
          { id: "430000-A556", description: "Assy-Ground Master Panel" },
          // ... up to 10 items
        ],
        pneumatic: [ // For 430001-A714
          { id: "430001-A286", description: "Assy-Module Precision Regulator Set" },
          // ... up to 3 items
        ]
      },
      toolingOptions: {
        "430001-A689": [
          { id: "430002-N656", description: "4 POS-Nozzle-1.7x1.7", category: "RUBBER TIP INTERFACE" },
          // ...
        ],
        // ...
      }
    },
    
    // Current State (from ORDERING LIST)
    coreItems: [
      { lineNumber: 1, id: "430000-S062-R02", description: "Machine Base Structure", quantity: 1, isChecked: true, checkDate: "30/12/2025", releaseType: "CHARGE OUT" },
      // ... 30 items
    ],
    
    configItems: [
      {
        slotIndex: 0,
        id: "430001-A366",
        description: "Assy-PX Vision PC Standard",
        isSpecialTrigger: false,
        shoppingListSelections: []
      },
      {
        slotIndex: 1,
        id: "430001-A714",
        description: "List-Optional Pneumatic Module",
        isSpecialTrigger: true,
        shoppingListSelections: ["", "", ""] // 3 slots for pneumatic
      },
      // ... up to 10 items
    ],
    
    modules: [
      {
        slotIndex: 0,
        instanceNumber: 1,
        instanceTotal: 2, // Same parent appears twice
        parentId: "430000-A961",
        parentDescription: "Module-Taping-v2.0",
        children: {
          electrical: {
            autoSelectedId: "...",
            currentId: "...",
            isOverridden: false,
            options: [...]
          },
          tooling: [
            {
              id: "430001-A704",
              description: "List-Tooling Tape And Reel",
              selectedOption: null,
              options: [...]
            }
          ],
          jigs: [],
          vision: null
        }
      },
      // ... up to 10 modules
    ],
    
    visionItems: [
      {
        slotIndex: 0,
        id: "430001-A013",
        description: "Module-Mark Package Vision-B2K",
        category: "Vision B2K"
      },
      // ... up to 10 items
    ]
  }
}
```

---

## 4. Business Logic Replication

### 4.1 CONFIG Section Logic

#### Special Triggers
```javascript
const SPECIAL_TRIGGERS = {
  '430001-A378': {
    name: 'Basic Tool Kit',
    shoppingListKey: 'basicTool',
    shoppingListSize: 10,
    shoppingListColumns: { ids: 'I', descriptions: 'J' }
  },
  '430001-A714': {
    name: 'Pneumatic Kit',
    shoppingListKey: 'pneumatic',
    shoppingListSize: 3,
    shoppingListColumns: { ids: 'K', descriptions: 'L' }
  }
};
```

#### Config Selection Flow
```
User selects Part ID in CONFIG slot
    │
    ├── Is it a special trigger?
    │   ├── YES: Expand shopping list (10 or 3 items)
    │   │        Insert dropdown rows in sheet
    │   │        Return expanded state to frontend
    │   │
    │   └── NO: Just update Part ID
    │           Return simple config item
    │
    └── Was previous value a special trigger?
        └── YES: Delete shopping list rows
```

### 4.2 MODULE Section Logic

#### Constants
```javascript
const RUBBER_TIP_PARENTS = ["430001-A689", "430001-A690", "430001-A691", "430001-A692"];
const RUBBER_TIP_SOURCE_ID = "430001-A380";
```

#### Electrical Kit Rotation
```javascript
function calculateElectricalKit(parentId, moduleSection, currentRowIndex) {
  // Count instances of same parent BEFORE and INCLUDING current row
  let instanceCount = 0;
  for (let i = 0; i <= currentRowIndex; i++) {
    if (moduleSection[i].parentId === parentId) {
      instanceCount++;
    }
  }
  
  // Get electrical options from REF_DATA
  const elecIds = getElectricalIds(parentId); // Array
  
  // Calculate rotation index
  const index = (instanceCount - 1) % elecIds.length;
  
  return elecIds[index];
}
```

#### Module Insertion Order
```
1. ELECTRICAL KIT (Single, auto-selected via rotation)
   └── Type: 'child'
   └── Static value, no dropdown

2. TOOLING KITS (Multiple, stacked)
   ├── Type: 'child' for each tooling kit
   │
   └── For each tooling kit:
       ├── Check for tooling OPTIONS
       │   └── Type: 'grandchild' (dropdown from shadow menu)
       │
       └── Check for RUBBER TIP (if parent in RUBBER_TIP_PARENTS)
           └── Type: 'rubber_tip' (dropdown from 430001-A380 options)

3. JIGS (Multiple, stacked)
   └── Type: 'jig'
   └── Static value, no dropdown

4. VISION (Single, last)
   ├── visionIds.length === 0: No vision row
   ├── visionIds.length === 1: Type 'vision_fixed' (static)
   └── visionIds.length > 1: Type 'vision_select' (dropdown)
```

#### Child Deletion Logic
```javascript
function deleteModuleChildren(sheet, parentRow, parentConfig) {
  // Collect all possible child IDs
  const possibleChildren = [
    ...parseIds(parentConfig.elecIds),
    ...parseIds(parentConfig.toolIds),
    ...parseIds(parentConfig.jigIds),
    ...parseIds(parentConfig.visionIds),
    ...getToolingOptionIds(parentConfig.toolIds),
    ...getRubberTipOptionIds(parentConfig.toolIds)
  ];
  
  // Delete rows from parentRow + 1 until hitting unrelated content
  let checkRow = parentRow + 1;
  while (checkRow <= sheet.getMaxRows()) {
    const childPartId = sheet.getRange(checkRow, 4).getValue();
    const itemNumber = sheet.getRange(checkRow, 3).getValue();
    
    // Stop if:
    // - Row has an ITEM number (next parent)
    // - Part ID is not in possible children AND not empty
    if (itemNumber !== "" || 
        (childPartId !== "" && !possibleChildren.includes(childPartId))) {
      break;
    }
    
    // Delete this row
    sheet.deleteRow(checkRow);
    // Don't increment checkRow since rows shift up
  }
}
```

### 4.3 VISION Section Logic

#### Standalone Vision
```javascript
// Uses categorized dropdown from REF_DATA!AF:AF
// Category lookup: VLOOKUP(D, REF_DATA!AF:AH, 3)
// Description lookup: VLOOKUP(D, REF_DATA!AF:AH, 2)
```

### 4.4 Renumber Kits Logic

```javascript
function renumberKits() {
  // 1. Get all modules in MODULE section
  // 2. For each unique parent ID, track instance counts
  // 3. Recalculate expected electrical kit for each instance
  // 4. If current != expected, update the electrical child row
  
  const parentCounts = {};
  
  for (const module of modules) {
    const parentId = module.parentId;
    const config = getParentConfig(parentId);
    
    if (config && config.elecIds) {
      parentCounts[parentId] = (parentCounts[parentId] || 0) + 1;
      const count = parentCounts[parentId];
      
      const elecIds = parseIds(config.elecIds);
      const index = (count - 1) % elecIds.length;
      const expectedElecId = elecIds[index];
      
      // Update if different
      if (module.children.electrical.currentId !== expectedElecId) {
        updateElectricalKit(module.row, expectedElecId);
      }
    }
  }
}
```

---

## 5. Implementation Phases

### Phase 1: Foundation (Priority 1)
**Goal**: Basic read/write operations working

| Task | Description | Files |
|------|-------------|-------|
| 1.1 | Create `api.gs` with basic structure | `api.gs` |
| 1.2 | Implement `getFullState()` | `api.gs` |
| 1.3 | Implement `getRefData()` | `api.gs` |
| 1.4 | Update frontend to call GAS on load | `app.js` |
| 1.5 | Create loading/error states in UI | `components.js`, `custom.css` |

**Deliverables**:
- Frontend loads real data from spreadsheet
- REF_DATA populates all dropdowns
- Current ORDERING LIST state displayed

### Phase 2: CONFIG Operations (Priority 1)
**Goal**: Full CONFIG section functionality

| Task | Description | Files |
|------|-------------|-------|
| 2.1 | Implement `addConfigItem()` | `api.gs` |
| 2.2 | Implement `removeConfigItem()` | `api.gs` |
| 2.3 | Implement `updateConfigShoppingList()` | `api.gs` |
| 2.4 | Handle special triggers (A378, A714) | `api.gs` |
| 2.5 | Update frontend state management | `state.js` |
| 2.6 | Connect frontend actions to API | `components.js` |

**Deliverables**:
- Add/remove config items
- Special triggers expand shopping lists
- Shopping list selections save to sheet

### Phase 3: MODULE Operations (Priority 1)
**Goal**: Full MODULE section functionality

| Task | Description | Files |
|------|-------------|-------|
| 3.1 | Implement `addModule()` with all children | `api.gs` |
| 3.2 | Implement `removeModule()` with cleanup | `api.gs` |
| 3.3 | Implement `updateModuleChild()` | `api.gs` |
| 3.4 | Implement `overrideElectricalKit()` | `api.gs` |
| 3.5 | Implement electrical rotation logic | `api.gs` |
| 3.6 | Handle tooling options + rubber tips | `api.gs` |
| 3.7 | Update frontend module rendering | `components.js` |

**Deliverables**:
- Add/remove modules with automatic children
- Electrical kit rotation works
- Tooling options selectable
- Vision within modules works

### Phase 4: VISION Operations (Priority 1)
**Goal**: Full VISION section functionality

| Task | Description | Files |
|------|-------------|-------|
| 4.1 | Implement `addVisionItem()` | `api.gs` |
| 4.2 | Implement `removeVisionItem()` | `api.gs` |
| 4.3 | Implement `updateVisionItem()` | `api.gs` |
| 4.4 | Update frontend vision rendering | `components.js` |

**Deliverables**:
- Add/remove/update standalone vision items
- Category display works

### Phase 5: Admin Functions (Priority 1)
**Goal**: Sync and renumber functionality

| Task | Description | Files |
|------|-------------|-------|
| 5.1 | Implement `renumberKits()` | `api.gs` |
| 5.2 | Implement `triggerMasterSync()` | `api.gs` |
| 5.3 | Add admin buttons to frontend | `index.html`, `components.js` |

**Deliverables**:
- Renumber Kits button works
- Master Sync button works
- Sync status displayed

### Phase 6: Checkbox Operations (Priority 2)
**Goal**: Check/uncheck with password protection

| Task | Description | Files |
|------|-------------|-------|
| 6.1 | Implement `checkItem()` | `api.gs` |
| 6.2 | Implement `uncheckItem()` with password | `api.gs` |
| 6.3 | Implement `updateReleaseType()` | `api.gs` |
| 6.4 | Implement password validation | `api.gs` |
| 6.5 | Store password in Script Properties | `api.gs` |
| 6.6 | Update frontend checkbox handling | `components.js`, `state.js` |

**Deliverables**:
- Checking items adds timestamp
- Unchecking requires password
- Release type selectable after check

### Phase 7: PDF Export (Priority 3)
**Goal**: Export order list to PDF

| Task | Description | Files |
|------|-------------|-------|
| 7.1 | Enhance client-side PDF generation | `components.js` |
| 7.2 | (Optional) Server-side PDF via GAS | `api.gs` |

**Deliverables**:
- PDF export button generates downloadable PDF

---

## 6. File Structure

### 6.1 Backend (Google Apps Script)

```
Google Apps Script Project
├── ORDERING_LIST.gs      (Existing - DO NOT MODIFY)
├── handleCheckBox.gs     (Existing - DO NOT MODIFY)
├── api.gs                (NEW - Frontend API layer)
├── helpers.gs            (NEW - Shared utilities)
└── config.gs             (NEW - Configuration constants)
```

### 6.2 Frontend

```
frontend/
├── index.html            (Main HTML - add GAS include)
├── css/
│   └── custom.css        (Styles)
└── js/
    ├── app.js            (Main entry - update for GAS)
    ├── state.js          (State management - update)
    ├── components.js     (UI rendering - update)
    └── gasApi.js         (NEW - GAS API wrapper)
```

### 6.3 New File: gasApi.js

```javascript
/**
 * Google Apps Script API Wrapper
 * Provides promise-based interface for GAS calls
 */

const GAS_API = {
  /**
   * Call a GAS function with error handling
   * @param {string} functionName - Name of GAS function
   * @param {...any} args - Arguments to pass
   * @returns {Promise} Resolves with result or rejects with error
   */
  call: function(functionName, ...args) {
    return new Promise((resolve, reject) => {
      google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject)
        [functionName](...args);
    });
  },
  
  // Convenience methods
  getFullState: () => GAS_API.call('getFullState'),
  getRefData: () => GAS_API.call('getRefData'),
  addConfigItem: (slot, id) => GAS_API.call('addConfigItem', slot, id),
  removeConfigItem: (slot) => GAS_API.call('removeConfigItem', slot),
  // ... etc
};
```

---

## 7. Error Handling

### 7.1 Error Types

| Code | Description | User Message |
|------|-------------|--------------|
| `CONNECTION_ERROR` | Cannot reach GAS backend | "Connection lost. Please check your internet connection." |
| `INVALID_PART_ID` | Part ID not in REF_DATA | "Part ID not found in reference data." |
| `SLOT_OCCUPIED` | Trying to add to occupied slot | "This slot already has an item." |
| `SLOT_EMPTY` | Trying to remove from empty slot | "This slot is already empty." |
| `INVALID_PASSWORD` | Wrong password for uncheck | "Incorrect password." |
| `SHEET_ERROR` | Spreadsheet operation failed | "Failed to update spreadsheet. Please try again." |
| `SYNC_ERROR` | Master sync failed | "Sync failed. Please check source spreadsheet access." |

### 7.2 Frontend Error Handling

```javascript
async function handleApiCall(apiFunction, loadingMessage) {
  try {
    // Show loading state
    showLoading(loadingMessage);
    
    // Make API call
    const result = await apiFunction();
    
    if (result.success) {
      return result.data;
    } else {
      showError(result.error.message);
      return null;
    }
  } catch (error) {
    showError('Connection error. Please try again.');
    console.error('API Error:', error);
    return null;
  } finally {
    hideLoading();
  }
}
```

### 7.3 Offline Detection

```javascript
// Block editing when offline
window.addEventListener('online', () => {
  AppState.connection.status = 'online';
  updateConnectionStatus();
  showToast('Connection restored', 'success');
});

window.addEventListener('offline', () => {
  AppState.connection.status = 'offline';
  updateConnectionStatus();
  showToast('Connection lost. Editing disabled.', 'error');
});

function canEdit() {
  return AppState.connection.status === 'online';
}
```

---

## 8. Testing Strategy

### 8.1 Unit Tests (GAS)

| Test | Description |
|------|-------------|
| `testGetFullState` | Verify all sections are loaded correctly |
| `testConfigSpecialTriggers` | Verify shopping list expansion |
| `testModuleChildInsertion` | Verify correct children are inserted |
| `testElectricalRotation` | Verify rotation math |
| `testRenumberKits` | Verify renumbering updates correctly |

### 8.2 Integration Tests

| Test | Description |
|------|-------------|
| Add config item | Frontend → API → Sheet updated → Frontend updated |
| Add module | Children appear in correct order |
| Remove module | All children deleted |
| Master sync | REF_DATA and sections updated |

### 8.3 Manual Testing Checklist

- [ ] Load page - all data displays correctly
- [ ] Add CONFIG item (regular)
- [ ] Add CONFIG item (special trigger - verify shopping list)
- [ ] Remove CONFIG item (special trigger - verify cleanup)
- [ ] Add MODULE - verify electrical rotation
- [ ] Add same MODULE again - verify different electrical kit
- [ ] Select tooling option
- [ ] Remove MODULE - verify all children deleted
- [ ] Add VISION item
- [ ] Check item - verify timestamp
- [ ] Uncheck item - verify password required
- [ ] Renumber Kits button
- [ ] Master Sync button
- [ ] Export PDF

---

## Appendix A: Deployment Instructions

### A.1 Deploy Web App

1. Open Google Apps Script editor
2. Click **Deploy** → **New deployment**
3. Select type: **Web app**
4. Settings:
   - Description: "ViTrox BOM Configurator v1.0"
   - Execute as: **Me**
   - Who has access: **Anyone**
5. Click **Deploy**
6. Copy the Web App URL

### A.2 Update Frontend for Deployment

The frontend HTML must be served via GAS `HtmlService`:

```javascript
// In api.gs
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ViTrox BOM Configurator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

### A.3 Include Frontend Files in GAS

Convert frontend files to GAS HTML files:
- `index.html` → Keep as main entry
- `custom.css` → Include via `<?!= include('css') ?>`
- `*.js` → Include via `<?!= include('js-app') ?>`

---

## Appendix B: Quick Reference

### B.1 Special IDs

| ID | Purpose |
|----|---------|
| `430001-A378` | Basic Tool Kit trigger (10 shopping items) |
| `430001-A714` | Pneumatic Kit trigger (3 shopping items) |
| `430001-A380` | Rubber Tip source ID |
| `430001-A689` | Rubber Tip parent 1 |
| `430001-A690` | Rubber Tip parent 2 |
| `430001-A691` | Rubber Tip parent 3 |
| `430001-A692` | Rubber Tip parent 4 |

### B.2 REF_DATA Quick Reference

| What | Columns |
|------|---------|
| Config options | A:B |
| Module options | C:D |
| Basic Tool shopping list | I:J |
| Pneumatic shopping list | K:L |
| Tooling options database | P:Q (parent-child), R:S (cat-desc) |
| Tooling shadow menu | U:V |
| Electrical mappings | W:X |
| Tooling mappings | Y:Z |
| Jig mappings | AA:AB |
| Vision mappings | AC:AD |
| Vision database | AF:AH |

---

*Document Version: 1.0*
*Last Updated: January 2, 2026*
*Author: AI Assistant (Claude)*

