# Microsoft Fabric Capacity Extension - Complete Regeneration Instructions

This document contains comprehensive instructions for recreating the Microsoft Fabric Capacity Extension exactly as implemented.

## Overview

Create a Microsoft Edge browser extension that provides a GUI interface for managing Microsoft Fabric capacities across Azure subscriptions. The extension features OAuth2 authentication, cross-subscription discovery, tenant display, real-time status updates, and comprehensive logging.

## Extension Structure

- **Manifest Version**: 3 (modern Chrome/Edge extension format)
- **Extension Name**: "Fabric Capacity Extension"
- **Description**: "An Edge extension for managing Fabric Capacities"
- **Version**: 1.0

## UI Design and Layout

### Header Section
- **Title**: "Microsoft Fabric" displayed as h2
- **Tenant Display**: Shows authenticated tenant name below title (initially hidden)
- **Loading Indicator**: Right-justified loading message next to tenant info

### Main Controls
- **Capacity Selector**: Dropdown showing available Fabric capacities with status indicators and SKU information
- **Refresh Button**: Manual refresh button (⟳ symbol) next to dropdown
- **SKU Management Section**: Hidden by default, shown when capacity is selected
  - **SKU Dropdown**: Shows current SKU and available options
  - **Update SKU Button**: Blue button to initiate SKU changes
- **Action Buttons**: "Start Capacity" and "Stop Capacity" buttons

### Logging Section
- **Log Area**: Read-only textarea (rows=20, cols=80) with monospace font
- **Debug Toggle**: Checkbox labeled "Enable Debug Logging"

### Visual Design Standards
- **Width**: 500px popup window
- **Font**: 'Segoe UI' (Microsoft standard)
- **Background**: #f3f2f1 (Microsoft neutral background)
- **Colors**: 
  - Running status: #107C10 (Microsoft green)
  - Stopped status: #D83B01 (Microsoft red)
  - Loading/secondary text: #605e5c (Microsoft neutral)
- **Spacing**: 10px gaps between major elements, 8px for smaller elements
- **Buttons**: Rounded corners (2px), proper hover states, disabled states with 60% opacity

## Authentication Implementation

### OAuth2 Configuration
- **Client ID**: `b2f9922d-47b3-45de-be16-72911e143fa4`
- **Authentication Method**: `chrome.identity.launchWebAuthFlow`
- **Scope**: `https://management.azure.com/user_impersonation`
- **Tenant**: `common` (multi-tenant support)
- **Response Type**: `token` (implicit flow)

### Authentication Flow
1. Check for cached valid token first
2. If no valid token, launch OAuth2 flow using Azure AD v2.0 endpoints
3. Extract access token from response URL fragment
4. Decode JWT token to extract tenant information
5. Cache token with expiry time
6. Update UI with tenant display

### Token Management
- **Caching**: Store in `chrome.storage.local` with expiry timestamp
- **Validation**: Test token with lightweight API call
- **Refresh**: Automatic refresh on 401 responses
- **Expiry Buffer**: 5-minute buffer before token expiry
- **Security**: No sensitive data in localStorage, proper state validation

### Tenant Information Extraction
- Decode JWT access token to extract tenant details
- Display hierarchy: `tenant_display_name` → domain from UPN → `tid`
- Update UI to show tenant context
- Clear display when authentication is cleared

## Azure API Integration

### API Endpoints and Versions
- **Base URL**: `https://management.azure.com`
- **Subscriptions**: `/subscriptions?api-version=2022-12-01`
- **Fabric Capacities**: `/subscriptions/{id}/providers/Microsoft.Fabric/capacities?api-version=2023-11-01`
- **Capacity Operations**: `{capacityId}/{action}?api-version=2023-11-01`

### API Operations
1. **List Subscriptions**: Get all accessible Azure subscriptions
2. **List Capacities**: Query Fabric capacities across all subscriptions with SKU information
3. **Start Capacity**: POST to `{capacityId}/resume`
4. **Stop Capacity**: POST to `{capacityId}/suspend`
5. **Update SKU**: PATCH to `{capacityId}` with new SKU in request body

### Error Handling
- Handle 404 for subscriptions without Fabric provider
- Retry on 401 with token refresh
- Graceful handling of network errors
- User-friendly error messages
- Comprehensive debug logging

### Key Features Implementation

#### 1. Initialization and DOM Management
- Initialize all DOM element references with validation
- Setup event listeners with error handling
- Verify all required elements exist before proceeding

#### 2. Authentication with Tenant Display
- Complete OAuth2 flow with Azure AD
- JWT token decoding for tenant information
- Token caching with intelligent expiry management
- UI updates for authentication state

#### 3. Capacity Discovery and Management
- Cross-subscription capacity discovery with SKU information display
- Real-time status display in dropdown with SKU details
- State-based button enabling/disabling
- Manual refresh capability
- SKU management interface that appears when capacity is selected

#### 3.1. SKU Management Implementation
- Display current SKU in capacity dropdown (e.g., "MyCapacity - F8 (Running)")
- Show/hide SKU container based on capacity selection
- Populate SKU dropdown with available options (F2, F4, F8, F16, F32, F64, F128, F256, F512)
- Enable "Update SKU" button only when different SKU is selected
- Confirmation dialog for SKU changes on running capacities
- PATCH API call to update capacity with new SKU
- Automatic refresh after SKU change with re-selection of capacity

### HTML Structure Details

#### Required DOM Elements
```html
<!-- Existing elements -->
<select id="capacitySelect" style="flex: 1;">
    <option value="">Select a capacity...</option>
</select>
<button id="refreshButton">⟳</button>

<!-- New SKU management section -->
<div id="skuContainer" style="display: none;">
    <div style="display: flex; gap: 10px; align-items: center; margin-top: 10px;">
        <label for="skuSelect" style="font-size: 14px; color: #323130; font-weight: 600; min-width: 80px;">Current SKU:</label>
        <select id="skuSelect" style="flex: 1;">
            <option value="">Loading SKUs...</option>
        </select>
        <button id="updateSkuButton" style="flex: 0 0 auto; padding: 8px 12px; background-color: #0078d4; color: white;" disabled>
            Update SKU
        </button>
    </div>
</div>

<!-- Existing action buttons -->
<button id="startButton" class="start-button" disabled>Start Capacity</button>
<button id="stopButton" class="stop-button" disabled>Stop Capacity</button>
```

#### Additional CSS Styles
```css
.update-sku-button {
    background-color: #0078d4;
    color: white;
}

.update-sku-button:hover:not(:disabled) {
    background-color: #106ebe;
}

.update-sku-button:disabled {
    background-color: #8a8886;
    color: #a19f9d;
}

.sku-container {
    display: flex;
    flex-direction: column;
    gap: 8px;
    padding: 12px;
    background-color: #faf9f8;
    border: 1px solid #e1dfdd;
    border-radius: 2px;
    margin: 8px 0;
}

.current-sku {
    font-size: 14px;
    color: #0078d4;
    font-weight: 600;
    background-color: #f3f9fd;
    padding: 4px 8px;
    border-radius: 2px;
    border: 1px solid #deecf9;
}
```

### JavaScript Implementation Details

#### Required Class Properties (add to constructor)
```javascript
this.skuContainer = null;
this.skuSelect = null;
this.updateSkuButton = null;
this.availableSkus = [];
```

#### Required Event Listeners
```javascript
this.updateSkuButton.addEventListener('click', () => {
    this.updateCapacitySku();
});

this.skuSelect.addEventListener('change', () => {
    this.onSkuSelectionChange();
});
```

#### Required Methods
1. `loadAvailableSkus(capacity)` - Load available SKU options
2. `populateSkuDropdown(capacity)` - Populate SKU dropdown with current and available options
3. `onSkuSelectionChange()` - Handle SKU selection changes and enable/disable update button
4. `updateCapacitySku()` - Perform PATCH API call to update capacity SKU

#### SKU Options
Standard Fabric SKUs: F2, F4, F8, F16, F32, F64, F128, F256, F512
Display format: "F8 (8 vCores, 16 GB RAM)" with descriptions
