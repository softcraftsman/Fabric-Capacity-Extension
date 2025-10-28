# Copilot Instructions for Fabric Capacity Extension

## Project Overview
- **Purpose:** Edge extension for managing Microsoft Fabric capacities across Azure subscriptions.
- **Key Features:**
  - Azure AD OAuth2 authentication (with token caching, refresh, and silent renewal)
  - Cross-subscription discovery of Fabric capacities
  - Start/stop and SKU management for capacities
  - Real-time status, logging, and debug mode
  - UI built in `popup.html` with logic in `popup.js`

## Architecture
- **popup.js:** Main logic for authentication, API calls, token management, and UI event handling. Uses a single class (`FabricCapacityManager`).
- **popup.html:** UI elements (dropdowns, buttons, log area) referenced by ID in JS.
- **manifest.json:** Declares permissions, popup, and host permissions for Azure and Microsoft login endpoints.
- **No background scripts or content scripts.**

## Authentication & Permissions
- Uses `chrome.identity.launchWebAuthFlow` for OAuth2 with Azure AD v2.0 endpoints.
- Requires these scopes:
  - `https://management.core.windows.net/user_impersonation`
  - `https://graph.microsoft.com/User.Read`
  - `offline_access`
- Tokens are cached in `chrome.storage.local` with expiry and session info.
- Token refresh is proactive (background timer) and silent when possible.
- User/tenant info is displayed in the UI header.

## API Integration
- **Azure Management API** for subscriptions and Fabric capacities (API versions: `2022-12-01`, `2023-11-01`).
- **Microsoft Graph** for user info.
- All API calls use bearer tokens from the authentication flow.

## UI/UX Patterns
- Dropdown auto-refreshes on open; status is color-coded (green for running, red for stopped).
- SKU changes prompt for confirmation if the capacity is running.
- Double-clicking the title clears authentication cache (for troubleshooting).
- Debug mode toggle persists in storage and increases log verbosity.

## Developer Workflows
- **Local testing:**
  - Load unpacked extension in Edge via `edge://extensions/`.
  - Click refresh on the extension card after code changes.
- **Distribution:**
  - Zip all files (including `icon.png`) for Edge Add-ons submission.
- **No build step** (plain JS/HTML/CSS).

## Error Handling & Troubleshooting
- Comprehensive error handling for auth, network, API, and permission issues.
- Debug logs are available in the UI and browser console.
- Common issues and solutions are documented in `README.md` (see Troubleshooting section).

## Project Conventions
- All DOM elements referenced by ID; keep IDs in sync between HTML and JS.
- All API endpoints and versions are constants in `popup.js`.
- No sensitive data is stored locally; only tokens (with expiry) in extension storage.
- All user-facing messages are clear and actionable.

## Key Files
- `popup.js`: Main logic, patterns, and conventions
- `popup.html`: UI structure and element IDs
- `manifest.json`: Permissions and extension config
- `README.md`: Full feature list, troubleshooting, and developer workflow

---

**For AI agents:**
- Follow the patterns in `popup.js` for API calls, token management, and UI updates.
- Reference `README.md` for feature behavior and troubleshooting.
- Use only the permissions and APIs declared in `manifest.json`.
- Maintain the user experience conventions (status colors, auto-refresh, debug logging, etc.).
