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
- **Capacity Selector**: Dropdown showing available Fabric capacities with status indicators
- **Refresh Button**: Manual refresh button (⟳ symbol) next to dropdown
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
2. **List Capacities**: Query Fabric capacities across all subscriptions
3. **Start Capacity**: POST to `{capacityId}/resume`
4. **Stop Capacity**: POST to `{capacityId}/suspend`

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
- Cross-subscription capacity discovery
- Real-time status display in dropdown
- State-based button enabling/disabling
- Manual refresh capability

#### 4. Operation Logging
- Timestamped logging to textarea
- Debug mode toggle with persistent preference
- Success/error message differentiation
- Fallback console logging

#### 5. UI State Management
- Loading indicators during operations
- Button state management based on selection
- Dropdown population with status indicators
- Error handling and user feedback

## Advanced Features

### Manual Refresh System
- Refresh button for manual capacity list updates
- Preserves user selection after refresh
- Disabled during loading operations
- Subtle UI feedback during refresh

### Debug Capabilities
- Comprehensive debug logging toggle
- Detailed API call logging
- Token validation information
- State change tracking

### Cache Management
- Double-click title to clear authentication cache
- Intelligent token expiry detection
- Automatic cache cleanup

### Error Recovery
- Automatic token refresh on API failures
- Graceful handling of network issues
- User-friendly error messaging
- Fallback operations for edge cases

## Complete Feature Set

### Production Features
- ✅ OAuth2 authentication with Azure AD
- ✅ Tenant name display and context
- ✅ Cross-subscription capacity discovery
- ✅ Real-time capacity status indicators
- ✅ One-click start/stop operations
- ✅ Manual refresh capability
- ✅ Comprehensive logging system
- ✅ Debug mode toggle
- ✅ Intelligent token caching
- ✅ Error handling and recovery
- ✅ Microsoft design compliance
- ✅ Responsive layout
- ✅ State persistence

### Security Features
- ✅ Secure OAuth2 implementation
- ✅ Token validation and refresh
- ✅ CSRF protection with state parameter
- ✅ No credentials stored locally
- ✅ Proper error boundary handling

### User Experience Features
- ✅ Loading indicators
- ✅ Visual status feedback
- ✅ Persistent user preferences
- ✅ Intuitive button states
- ✅ Clear error messaging
- ✅ Professional Microsoft styling

## Implementation Notes

1. **No Test/Mock Code**: Production-ready with real Azure API integration only
2. **Token Security**: JWT tokens properly decoded and validated
3. **Error Boundaries**: Comprehensive error handling at all levels
4. **Performance**: Efficient API calls with caching and validation
5. **Accessibility**: Proper labeling and keyboard navigation support
6. **Responsive Design**: Flexible layout accommodating various content lengths

## Deployment Checklist

- [ ] All files present and correctly named
- [ ] Icon files generated (PNG and SVG)
- [ ] No development/test code remaining
- [ ] All API endpoints using production URLs
- [ ] Error handling tested for all scenarios
- [ ] Authentication flow tested end-to-end
- [ ] UI tested with various data states
- [ ] Debug logging verified in both modes
- [ ] Token caching and expiry tested
- [ ] Cross-subscription functionality verified

This extension provides a complete, production-ready solution for managing Microsoft Fabric capacities with enterprise-grade authentication, comprehensive error handling, and a polished user experience following Microsoft design standards.
