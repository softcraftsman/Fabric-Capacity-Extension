/**
 * Microsoft Fabric Capacity Extension
 * Main JavaScript functionality for managing Fabric capacities
 */

class FabricCapacityManager {
    constructor() {
        this.accessToken = null;
        this.capacities = [];
        this.debugMode = false;
        this.logArea = null;
        this.capacitySelect = null;
        this.startButton = null;
        this.stopButton = null;
        this.loadingIndicator = null;
        this.debugToggle = null;
        this.tenantInfo = null;
        this.initialLoadComplete = false;
        
        // API endpoints and configuration
        this.baseUrl = 'https://management.azure.com';
        this.subscriptionApiVersion = '2022-12-01';
        this.fabricApiVersion = '2023-11-01';
        this.scope = 'https://management.azure.com/user_impersonation';
    }

    /**
     * Initialize the extension
     */
    async init() {
        try {
            // Initialize DOM elements first
            const elementsInitialized = this.initializeElements();
            if (!elementsInitialized) {
                this.logError('Failed to initialize DOM elements');
                return;
            }

            this.setupEventListeners();
            this.log('Extension initialized');
            
            await this.authenticate();
            await this.loadCapacities();
        } catch (error) {
            this.logError('Failed to initialize extension', error);
            console.error('Extension initialization failed:', error);
        }
    }

    /**
     * Initialize DOM elements
     */
    initializeElements() {
        this.logArea = document.getElementById('logArea');
        this.capacitySelect = document.getElementById('capacitySelect');
        this.startButton = document.getElementById('startButton');
        this.stopButton = document.getElementById('stopButton');
        this.loadingIndicator = document.getElementById('loadingIndicator');
        this.debugToggle = document.getElementById('debugToggle');
        this.refreshButton = document.getElementById('refreshButton');
        this.tenantInfo = document.getElementById('tenantInfo');

        // Verify all elements were found
        const elements = {
            logArea: this.logArea,
            capacitySelect: this.capacitySelect,
            startButton: this.startButton,
            stopButton: this.stopButton,
            loadingIndicator: this.loadingIndicator,
            debugToggle: this.debugToggle,
            refreshButton: this.refreshButton,
            tenantInfo: this.tenantInfo
        };

        for (const [name, element] of Object.entries(elements)) {
            if (!element) {
                console.error(`Failed to find element: ${name}`);
                alert(`Failed to initialize: ${name} element not found`);
                return false;
            }
        }

        // Load debug mode preference
        chrome.storage.local.get(['debugMode'], (result) => {
            this.debugMode = result.debugMode || false;
            this.debugToggle.checked = this.debugMode;
        });

        return true;
    }

    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Add error handling for event listeners
        try {
            this.capacitySelect.addEventListener('change', () => {
                this.onCapacitySelectionChange();
            });

            this.refreshButton.addEventListener('click', async () => {
                this.log('Manual refresh requested');
                await this.refreshCapacities();
            });

            this.startButton.addEventListener('click', () => {
                this.startCapacity();
            });

            this.stopButton.addEventListener('click', () => {
                this.stopCapacity();
            });

            this.debugToggle.addEventListener('change', (e) => {
                this.debugMode = e.target.checked;
                chrome.storage.local.set({ debugMode: this.debugMode });
                this.log(`Debug logging ${this.debugMode ? 'enabled' : 'disabled'}`);
            });

            // Add double-click on title to clear authentication (for testing/troubleshooting)
            document.querySelector('h2').addEventListener('dblclick', async () => {
                if (confirm('Clear cached authentication? This will require re-login.')) {
                    await this.clearCachedAuth();
                    this.accessToken = null;
                    this.log('Authentication cache cleared');
                    this.capacities = [];
                    this.initialLoadComplete = false;
                    this.populateCapacityDropdown();
                }
            });

            this.log('Event listeners setup complete');
        } catch (error) {
            this.logError('Failed to setup event listeners', error);
            console.error('Event listener setup failed:', error);
        }
    }

    /**
     * Authenticate with Azure AD
     */
    async authenticate() {
        try {
            this.log('Authenticating with Azure AD...');
            this.showLoading(true);

            // Check if we have a cached token first
            const cachedToken = await this.getCachedToken();
            if (cachedToken && await this.validateToken(cachedToken)) {
                this.accessToken = cachedToken;
                this.updateTenantDisplay(cachedToken);
                this.log('Using cached authentication token');
                this.debugLog('Cached token is valid');
                return cachedToken;
            }

            // Perform OAuth2 flow using launchWebAuthFlow
            const token = await this.performOAuth2Flow();
            
            this.accessToken = token;
            this.updateTenantDisplay(token);
            await this.cacheToken(token);
            this.log('Authentication successful');
            this.debugLog('New access token obtained and cached');
            
            return token;
        } catch (error) {
            this.logError('Authentication error', error);
            throw error;
        } finally {
            this.showLoading(false);
        }
    }

    /**
     * Perform OAuth2 authentication flow
     */
    async performOAuth2Flow() {
        return new Promise((resolve, reject) => {
            // Azure AD OAuth2 endpoints
            const tenantId = 'common'; // Use 'common' for multi-tenant apps
            const clientId = 'b2f9922d-47b3-45de-be16-72911e143fa4';
            const redirectUri = chrome.identity.getRedirectURL();
            const scope = encodeURIComponent(this.scope);
            const responseType = 'token';
            const state = this.generateState();

            // Construct authorization URL
            const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
                `?client_id=${clientId}` +
                `&response_type=${responseType}` +
                `&redirect_uri=${encodeURIComponent(redirectUri)}` +
                `&scope=${scope}` +
                `&state=${state}` +
                `&response_mode=fragment` +
                `&prompt=select_account`;

            this.debugLog(`Starting OAuth2 flow with URL: ${authUrl}`);
            this.debugLog(`Redirect URI: ${redirectUri}`);

            chrome.identity.launchWebAuthFlow({
                url: authUrl,
                interactive: true
            }, (responseUrl) => {
                if (chrome.runtime.lastError) {
                    this.logError('OAuth2 flow failed', chrome.runtime.lastError);
                    reject(chrome.runtime.lastError);
                    return;
                }

                if (!responseUrl) {
                    const error = new Error('No response URL received from OAuth2 flow');
                    this.logError('OAuth2 flow failed', error);
                    reject(error);
                    return;
                }

                this.debugLog(`OAuth2 response URL: ${responseUrl}`);

                try {
                    const token = this.extractTokenFromUrl(responseUrl, state);
                    resolve(token);
                } catch (error) {
                    this.logError('Failed to extract token from response', error);
                    reject(error);
                }
            });
        });
    }

    /**
     * Extract access token from OAuth2 response URL
     */
    extractTokenFromUrl(responseUrl, expectedState) {
        const url = new URL(responseUrl);
        const fragment = url.hash.substring(1); // Remove the # character
        const params = new URLSearchParams(fragment);

        // Verify state parameter
        const state = params.get('state');
        if (state !== expectedState) {
            throw new Error('State parameter mismatch - possible CSRF attack');
        }

        // Check for errors
        const error = params.get('error');
        if (error) {
            const errorDescription = params.get('error_description') || 'Unknown error';
            throw new Error(`OAuth2 error: ${error} - ${decodeURIComponent(errorDescription)}`);
        }

        // Extract access token
        const accessToken = params.get('access_token');
        if (!accessToken) {
            throw new Error('No access token found in OAuth2 response');
        }

        // Log token info for debugging (without exposing the actual token)
        const expiresIn = params.get('expires_in');
        const tokenType = params.get('token_type');
        this.debugLog(`Token type: ${tokenType}, expires in: ${expiresIn} seconds`);

        return accessToken;
    }

    /**
     * Generate a random state parameter for CSRF protection
     */
    generateState() {
        const array = new Uint32Array(4);
        crypto.getRandomValues(array);
        return Array.from(array, dec => ('0' + dec.toString(16)).substr(-2)).join('');
    }

    /**
     * Get cached authentication token
     */
    async getCachedToken() {
        return new Promise((resolve) => {
            chrome.storage.local.get(['accessToken', 'tokenExpiry'], (result) => {
                if (chrome.runtime.lastError) {
                    this.debugLog('Failed to get cached token: ' + chrome.runtime.lastError.message);
                    resolve(null);
                    return;
                }

                const { accessToken, tokenExpiry } = result;
                
                if (!accessToken || !tokenExpiry) {
                    this.debugLog('No cached token found');
                    resolve(null);
                    return;
                }

                // Check if token is expired (with 5-minute buffer)
                const now = Date.now();
                const expiryTime = parseInt(tokenExpiry);
                const bufferTime = 5 * 60 * 1000; // 5 minutes in milliseconds

                if (now >= (expiryTime - bufferTime)) {
                    this.debugLog('Cached token is expired');
                    resolve(null);
                    return;
                }

                this.debugLog('Found valid cached token');
                resolve(accessToken);
            });
        });
    }

    /**
     * Cache authentication token with expiry
     */
    async cacheToken(token) {
        // Azure AD tokens typically expire in 1 hour
        const expiryTime = Date.now() + (60 * 60 * 1000); // 1 hour from now
        
        return new Promise((resolve) => {
            chrome.storage.local.set({
                accessToken: token,
                tokenExpiry: expiryTime.toString()
            }, () => {
                if (chrome.runtime.lastError) {
                    this.debugLog('Failed to cache token: ' + chrome.runtime.lastError.message);
                } else {
                    this.debugLog('Token cached successfully');
                }
                resolve();
            });
        });
    }

    /**
     * Validate if a token is still valid by making a test API call
     */
    async validateToken(token) {
        try {
            const testUrl = `${this.baseUrl}/subscriptions?api-version=${this.subscriptionApiVersion}&$top=1`;
            const response = await fetch(testUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': 'application/json'
                }
            });

            return response.ok;
        } catch (error) {
            this.debugLog('Token validation failed: ' + error.message);
            return false;
        }
    }

    /**
     * Clear cached authentication data
     */
    async clearCachedAuth() {
        return new Promise((resolve) => {
            chrome.storage.local.remove(['accessToken', 'tokenExpiry'], () => {
                this.debugLog('Cached authentication data cleared');
                // Clear tenant display
                if (this.tenantInfo) {
                    this.tenantInfo.style.display = 'none';
                    this.tenantInfo.textContent = '';
                }
                resolve();
            });
        });
    }

    /**
     * Load all Fabric capacities across subscriptions
     */
    async loadCapacities() {
        try {
            this.log('Loading Fabric capacities...');
            this.showLoading(true);
            this.capacities = [];

            // Get all subscriptions
            const subscriptions = await this.getSubscriptions();
            this.debugLog(`Found ${subscriptions.length} subscriptions`);

            // Get capacities from all subscriptions
            for (const subscription of subscriptions) {
                try {
                    const capacities = await this.getCapacitiesForSubscription(subscription.subscriptionId);
                    this.capacities.push(...capacities);
                    this.debugLog(`Found ${capacities.length} capacities in subscription ${subscription.displayName}`);
                } catch (error) {
                    this.debugLog(`Failed to get capacities for subscription ${subscription.subscriptionId}: ${error.message}`);
                    // Continue with other subscriptions
                }
            }

            this.populateCapacityDropdown();
            this.log(`Loaded ${this.capacities.length} Fabric capacities`);
            this.initialLoadComplete = true;

        } catch (error) {
            this.logError('Failed to load capacities', error);
        } finally {
            this.showLoading(false);
        }
    }

    /**
     * Refresh capacities list (manual refresh only)
     */
    async refreshCapacities() {
        // Don't refresh if we're already loading or don't have authentication
        if (!this.accessToken || this.loadingIndicator.style.display === 'block') {
            this.debugLog('Refresh skipped - not ready or already loading');
            return;
        }

        try {
            this.debugLog('Refreshing capacity status...');
            
            // Store the currently selected capacity to restore selection after refresh
            const selectedIndex = this.capacitySelect.value;
            const selectedCapacityId = selectedIndex ? this.capacities[parseInt(selectedIndex)]?.id : null;

            // Show subtle loading without full loading message
            this.capacitySelect.disabled = true;
            this.refreshButton.disabled = true;
            
            // Get fresh capacity data
            const refreshedCapacities = [];
            
            // Get unique subscription IDs from current capacities
            const subscriptionIds = [...new Set(this.capacities.map(c => c.subscriptionId))];
            
            if (subscriptionIds.length === 0) {
                // If no capacities loaded yet, fall back to full load
                this.debugLog('No subscription IDs found, performing full load');
                await this.loadCapacities();
                return;
            }

            // Refresh capacities from known subscriptions
            for (const subscriptionId of subscriptionIds) {
                try {
                    const capacities = await this.getCapacitiesForSubscription(subscriptionId);
                    refreshedCapacities.push(...capacities);
                } catch (error) {
                    this.debugLog(`Failed to refresh capacities for subscription ${subscriptionId}: ${error.message}`);
                    // Continue with other subscriptions
                }
            }

            this.capacities = refreshedCapacities;
            this.populateCapacityDropdown();

            // Restore selection if the capacity still exists
            if (selectedCapacityId) {
                const newIndex = this.capacities.findIndex(c => c.id === selectedCapacityId);
                if (newIndex !== -1) {
                    this.capacitySelect.value = newIndex.toString();
                    this.onCapacitySelectionChange();
                }
            }

            this.debugLog(`Refreshed ${this.capacities.length} capacities`);

        } catch (error) {
            this.logError('Failed to refresh capacities', error);
        } finally {
            this.capacitySelect.disabled = false;
            this.refreshButton.disabled = false;
        }
    }

    /**
     * Get all subscriptions
     */
    async getSubscriptions() {
        const url = `${this.baseUrl}/subscriptions?api-version=${this.subscriptionApiVersion}`;
        const response = await this.makeApiCall(url);
        return response.value || [];
    }

    /**
     * Get Fabric capacities for a specific subscription
     */
    async getCapacitiesForSubscription(subscriptionId) {
        const url = `${this.baseUrl}/subscriptions/${subscriptionId}/providers/Microsoft.Fabric/capacities?api-version=${this.fabricApiVersion}`;
        
        try {
            const response = await this.makeApiCall(url);
            const capacities = response.value || [];
            
            // Add subscription context to each capacity
            return capacities.map(capacity => ({
                ...capacity,
                subscriptionId: subscriptionId,
                displayName: `${capacity.name} (${capacity.properties?.state || 'Unknown'})`
            }));
        } catch (error) {
            if (error.message.includes('404') || error.message.includes('ResourceProviderNotRegistered')) {
                // Subscription doesn't have Fabric provider registered
                this.debugLog(`Fabric provider not registered in subscription ${subscriptionId}`);
                return [];
            }
            throw error;
        }
    }

    /**
     * Populate the capacity dropdown
     */
    populateCapacityDropdown() {
        try {
            this.debugLog(`Populating dropdown with ${this.capacities.length} capacities`);
            
            if (!this.capacitySelect) {
                this.logError('Cannot populate dropdown - capacitySelect element not found');
                return;
            }

            // Clear existing options except the first one
            while (this.capacitySelect.children.length > 1) {
                this.capacitySelect.removeChild(this.capacitySelect.lastChild);
            }

            if (this.capacities.length === 0) {
                const option = document.createElement('option');
                option.value = '';
                option.textContent = 'No capacities found';
                option.disabled = true;
                this.capacitySelect.appendChild(option);
                this.debugLog('Added "No capacities found" option');
                return;
            }

            // Add capacity options
            this.capacities.forEach((capacity, index) => {
                const option = document.createElement('option');
                option.value = index.toString();
                
                const state = capacity.properties?.state || 'Unknown';
                const stateDisplay = state === 'Active' ? '(Running)' : 
                                   state === 'Paused' ? '(Stopped)' : 
                                   `(${state})`;
                
                option.textContent = `${capacity.name} ${stateDisplay}`;
                
                // Add visual styling based on state
                if (state === 'Active') {
                    option.className = 'status-running';
                } else if (state === 'Paused') {
                    option.className = 'status-stopped';
                }
                
                this.capacitySelect.appendChild(option);
                this.debugLog(`Added capacity option: ${option.textContent}`);
            });

            this.debugLog(`Dropdown population complete. Total options: ${this.capacitySelect.children.length}`);
        } catch (error) {
            this.logError('Failed to populate dropdown', error);
            console.error('Dropdown population failed:', error);
        }
    }

    /**
     * Handle capacity selection change
     */
    onCapacitySelectionChange() {
        const selectedIndex = this.capacitySelect.value;
        
        if (!selectedIndex || selectedIndex === '') {
            this.startButton.disabled = true;
            this.stopButton.disabled = true;
            return;
        }

        const capacity = this.capacities[parseInt(selectedIndex)];
        const state = capacity.properties?.state || 'Unknown';

        // Enable/disable buttons based on current state
        this.startButton.disabled = (state === 'Active');
        this.stopButton.disabled = (state === 'Paused');

        this.log(`Selected capacity: ${capacity.name} (${state})`);
        this.debugLog(`Capacity details: ${JSON.stringify(capacity, null, 2)}`);
    }

    /**
     * Start the selected capacity
     */
    async startCapacity() {
        await this.performCapacityOperation('resume', 'Starting');
    }

    /**
     * Stop the selected capacity
     */
    async stopCapacity() {
        await this.performCapacityOperation('suspend', 'Stopping');
    }

    /**
     * Perform capacity operation (start/stop)
     */
    async performCapacityOperation(operation, operationName) {
        const selectedIndex = this.capacitySelect.value;
        if (!selectedIndex || selectedIndex === '') return;

        const capacity = this.capacities[parseInt(selectedIndex)];
        
        try {
            this.log(`${operationName} capacity ${capacity.name}...`);
            this.setButtonsEnabled(false);

            const url = `${this.baseUrl}${capacity.id}/${operation}?api-version=${this.fabricApiVersion}`;
            
            await this.makeApiCall(url, 'POST');
            
            this.logSuccess(`${operationName} operation initiated for ${capacity.name}`);
            
            // Refresh the capacity list to update status after a short delay
            setTimeout(async () => {
                await this.refreshCapacities();
            }, 2000);

        } catch (error) {
            this.logError(`Failed to ${operation} capacity ${capacity.name}`, error);
        } finally {
            this.setButtonsEnabled(true);
        }
    }

    /**
     * Make API call to Azure
     */
    async makeApiCall(url, method = 'GET', body = null) {
        if (!this.accessToken) {
            throw new Error('No access token available');
        }

        this.debugLog(`Making ${method} request to: ${url}`);

        const options = {
            method: method,
            headers: {
                'Authorization': `Bearer ${this.accessToken}`,
                'Content-Type': 'application/json'
            }
        };

        if (body) {
            options.body = JSON.stringify(body);
        }

        const response = await fetch(url, options);

        if (!response.ok) {
            const errorText = await response.text();
            this.debugLog(`API call failed: ${response.status} ${response.statusText} - ${errorText}`);
            
            if (response.status === 401) {
                // Token might be expired, clear cache and try to re-authenticate
                this.debugLog('Token expired, attempting to re-authenticate');
                await this.clearCachedAuth();
                this.accessToken = null;
                await this.authenticate();
                
                // Retry the original request with new token
                options.headers['Authorization'] = `Bearer ${this.accessToken}`;
                const retryResponse = await fetch(url, options);
                
                if (!retryResponse.ok) {
                    const retryErrorText = await retryResponse.text();
                    this.debugLog(`Retry API call failed: ${retryResponse.status} ${retryResponse.statusText} - ${retryErrorText}`);
                    throw new Error(`API call failed after retry: ${retryResponse.status} ${retryResponse.statusText}`);
                }
                
                // Handle successful retry response
                if (method === 'POST' && retryResponse.status === 202) {
                    return { success: true };
                }
                
                const retryResponseText = await retryResponse.text();
                return retryResponseText ? JSON.parse(retryResponseText) : {};
            }
            
            throw new Error(`API call failed: ${response.status} ${response.statusText}`);
        }

        // Handle empty responses for POST operations
        if (method === 'POST' && response.status === 202) {
            return { success: true };
        }

        const responseText = await response.text();
        return responseText ? JSON.parse(responseText) : {};
    }

    /**
     * Enable/disable operation buttons
     */
    setButtonsEnabled(enabled) {
        if (enabled) {
            this.onCapacitySelectionChange(); // Re-evaluate button states
            this.refreshButton.disabled = false;
        } else {
            this.startButton.disabled = true;
            this.stopButton.disabled = true;
            this.refreshButton.disabled = true;
        }
    }

    /**
     * Show/hide loading indicator
     */
    showLoading(show) {
        this.loadingIndicator.style.display = show ? 'block' : 'none';
    }

    /**
     * Log message to the log area
     */
    log(message) {
        try {
            const timestamp = new Date().toLocaleTimeString();
            const logMessage = `${message}\n`;
            
            // Fallback if logArea is not initialized
            if (!this.logArea) {
                this.logArea = document.getElementById('logArea');
            }
            
            if (this.logArea) {
                this.logArea.value += logMessage;
                this.logArea.scrollTop = this.logArea.scrollHeight;
            } else {
                console.log('Extension Log:', logMessage);
            }
        } catch (error) {
            console.error('Logging failed:', error, 'Message was:', message);
        }
    }

    /**
     * Log debug message (only if debug mode is enabled)
     */
    debugLog(message) {
        if (this.debugMode) {
            const timestamp = new Date().toLocaleTimeString();
            const logMessage = `[${timestamp}] ${message}\n`;
            this.logArea.value += logMessage;
            this.logArea.scrollTop = this.logArea.scrollHeight;
        }
    }

    /**
     * Log error message
     */
    logError(message, error = null) {
        try {
            const timestamp = new Date().toLocaleTimeString();
            let logMessage = `[${timestamp}] ERROR: ${message}`;
            
            if (error) {
                logMessage += ` - ${error.message || error}`;
            }
            
            logMessage += '\n';
            
            // Fallback if logArea is not initialized
            if (!this.logArea) {
                this.logArea = document.getElementById('logArea');
            }
            
            if (this.logArea) {
                this.logArea.value += logMessage;
                this.logArea.scrollTop = this.logArea.scrollHeight;
            } else {
                console.error('Extension Error:', logMessage);
            }
        } catch (err) {
            console.error('Error logging failed:', err, 'Original message was:', message, 'Original error was:', error);
        }
    }

    /**
     * Log success message
     */
    logSuccess(message) {
        const timestamp = new Date().toLocaleTimeString();
        const logMessage = `[${timestamp}] SUCCESS: ${message}\n`;
        this.logArea.value += logMessage;
        this.logArea.scrollTop = this.logArea.scrollHeight;
    }

    /**
     * Decode JWT token to extract tenant information
     */
    decodeJwtToken(token) {
        try {
            // JWT tokens have 3 parts separated by dots: header.payload.signature
            const parts = token.split('.');
            if (parts.length !== 3) {
                this.debugLog('Invalid JWT token format');
                return null;
            }

            // Decode the payload (second part)
            const payload = parts[1];
            // Add padding if needed for base64 decoding
            const paddedPayload = payload + '='.repeat((4 - payload.length % 4) % 4);
            const decodedPayload = atob(paddedPayload.replace(/-/g, '+').replace(/_/g, '/'));
            
            return JSON.parse(decodedPayload);
        } catch (error) {
            this.debugLog('Failed to decode JWT token: ' + error.message);
            return null;
        }
    }

    /**
     * Update tenant display with information from the access token
     */
    updateTenantDisplay(token) {
        try {
            const tokenData = this.decodeJwtToken(token);
            if (!tokenData) {
                this.debugLog('Could not decode token for tenant information');
                return;
            }

            // Extract tenant information from the token
            const tenantId = tokenData.tid;
            const tenantName = tokenData.tenant_display_name || tokenData.tenant_name;
            const userPrincipalName = tokenData.upn || tokenData.unique_name;

            this.debugLog('Token data:', JSON.stringify({
                tenantId: tenantId,
                tenantName: tenantName,
                userPrincipalName: userPrincipalName
            }));

            if (tenantName) {
                this.tenantInfo.textContent = `Tenant: ${tenantName}`;
                this.tenantInfo.style.display = 'block';
                this.debugLog(`Tenant display updated: ${tenantName}`);
            } else if (userPrincipalName) {
                // Fallback to domain from UPN if tenant name not available
                const domain = userPrincipalName.split('@')[1];
                if (domain) {
                    this.tenantInfo.textContent = `Tenant: ${domain}`;
                    this.tenantInfo.style.display = 'block';
                    this.debugLog(`Tenant display updated with domain: ${domain}`);
                }
            } else if (tenantId) {
                // Last fallback to tenant ID
                this.tenantInfo.textContent = `Tenant ID: ${tenantId}`;
                this.tenantInfo.style.display = 'block';
                this.debugLog(`Tenant display updated with ID: ${tenantId}`);
            } else {
                this.debugLog('No tenant information available in token');
            }
        } catch (error) {
            this.debugLog('Failed to update tenant display: ' + error.message);
        }
    }
}

// Initialize the extension when the popup loads
document.addEventListener('DOMContentLoaded', () => {
    const manager = new FabricCapacityManager();
    manager.init();
});
