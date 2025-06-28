# Microsoft Fabric Capacity Extension

A Microsoft Edge browser extension that provides a GUI interface for managing Microsoft Fabric capacities across your Azure subscriptions.

## Features

- 🔐 **Azure AD Authentication** - Secure OAuth2 flow using `chrome.identity.launchWebAuthFlow` with token caching
- 🌐 **Cross-Subscription Discovery** - Automatically finds Fabric capacities across all your subscriptions
- ▶️ **Start/Stop Controls** - Easy one-click capacity management
- 📊 **Real-time Status** - Live status updates with automatic refresh when dropdown is accessed
- 📝 **Comprehensive Logging** - Operation logs with optional debug mode
- 🎨 **Microsoft Design** - Clean UI following Microsoft design principles
- 💾 **Token Caching** - Intelligent token management with automatic refresh
- 🔄 **Auto-refresh** - Capacity status automatically updates when dropdown is selected

## Installation

### 1. Prepare the Extension

1. Download or clone this repository to your local machine
2. Create an icon.png file (see `icon_instructions.txt` for guidance)
3. Ensure all files are in the same directory:
   - `manifest.json`
   - `popup.html`
   - `popup.js`
   - `icon.png`

### 2. Load in Microsoft Edge

1. Open Microsoft Edge
2. Go to `edge://extensions/`
3. Enable "Developer mode" (toggle in the left sidebar)
4. Click "Load unpacked"
5. Select the folder containing the extension files
6. The extension should now appear in your extensions list

### 3. Pin the Extension (Optional)

1. Click the extensions icon (puzzle piece) in the Edge toolbar
2. Find "Fabric Capacity Extension"
3. Click the pin icon to add it to your toolbar

## Usage

### First Time Setup

1. Click the extension icon in your Edge toolbar
2. The extension will prompt you to sign in to Azure AD
3. Grant the requested permissions for Azure Management API access
4. Wait for the extension to discover your Fabric capacities

### Managing Capacities

1. **Select a Capacity**: Use the dropdown to choose from your available capacities
   - **Auto-refresh**: Status is automatically updated each time you click on the dropdown
   - Running capacities show as "(Running)" in green
   - Stopped capacities show as "(Stopped)" in red
   - Selection is preserved after status refreshes

2. **Start a Capacity**: 
   - Select a stopped capacity
   - Click the "Start Capacity" button
   - Monitor the log area for operation status
   - Status will auto-refresh after 2 seconds

3. **Stop a Capacity**:
   - Select a running capacity
   - Click the "Stop Capacity" button
   - Monitor the log area for operation status
   - Status will auto-refresh after 2 seconds

### Debug Mode

- Enable "Enable Debug Logging" for detailed API call information
- Useful for troubleshooting authentication or API issues
- Debug preference is saved between sessions

## Architecture

```
Edge Extension ←→ Azure AD (OAuth2) ←→ Azure Management API ←→ Fabric Capacities
     (UI)           (Auth)              (Discovery/Control)      (Resources)
```

### Key Components

- **popup.html**: User interface with dropdown, buttons, and logging area
- **popup.js**: Core functionality including OAuth2 flow, API calls, and capacity management with token caching
- **manifest.json**: Extension configuration and permissions (no OAuth2 client configuration needed)

### API Integration

- **Authentication**: Uses `chrome.identity.launchWebAuthFlow` with Azure AD OAuth2 v2.0 endpoints
- **Token Management**: Intelligent caching with automatic expiry detection and refresh
- **Subscription Discovery**: Lists all accessible Azure subscriptions
- **Capacity Discovery**: Queries Microsoft.Fabric/capacities across subscriptions
- **Capacity Control**: Uses suspend/resume endpoints for start/stop operations

## Permissions

The extension requires the following permissions:

- `identity`: For Azure AD authentication
- `storage`: To save user preferences (debug mode)
- `activeTab`: For extension popup functionality
- `scripting`: For extension operations
- `https://management.azure.com/*`: For Azure API access

## API Versions

- **Subscriptions**: `2022-12-01`
- **Fabric Capacities**: `2023-11-01`

## Error Handling

The extension includes comprehensive error handling for:

- Authentication failures
- Network connectivity issues
- API rate limiting
- Missing permissions
- Subscription access issues
- Capacity operation failures

## Troubleshooting

### No Capacities Found

- Ensure you have Fabric capacities in your Azure subscriptions
- Verify the Microsoft.Fabric provider is registered in your subscriptions
- Check that you have appropriate permissions to view and manage Fabric resources

### Authentication Issues

- Clear browser cache and cookies for Azure AD
- Double-click the "Microsoft Fabric" title to clear cached authentication
- Try signing out and back in to the extension
- Verify your Azure AD account has access to the required subscriptions
- Check that the extension has permission to access login.microsoftonline.com

### API Errors

- Enable debug logging to see detailed error information
- Check Azure service health for any outages
- Verify your Azure permissions include Fabric resource management

## Development

### Local Testing

1. Make changes to the extension files
2. Go to `edge://extensions/`
3. Click the refresh button on the extension card
4. Test your changes

### Building for Distribution

1. Ensure all files are included and icon.png is created
2. Zip the extension folder
3. Submit to Microsoft Edge Add-ons store (if desired)

## Security

- No credentials are stored locally
- Uses secure OAuth2 flow with Azure AD
- API calls use bearer token authentication
- Follows Microsoft security best practices

## License

This project is provided as-is for educational and development purposes.

## Support

For issues or questions:

1. Check the debug logs for error details
2. Verify Azure permissions and subscriptions
3. Test with a simple capacity operation
4. Review the browser console for JavaScript errors
