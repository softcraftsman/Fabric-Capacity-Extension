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
        this.skuContainer = null;
        this.skuSelect = null;
        this.updateSkuButton = null;
        this.availableSkus = [];
        
        // API endpoints and configuration
        this.baseUrl = 'https://management.azure.com';
        this.graphUrl = 'https://graph.microsoft.com';
        this.subscriptionApiVersion = '2022-12-01';
        this.fabricApiVersion = '2023-11-01';
        
        // Scopes matching Entra API permissions exactly
        this.scopes = [
            'https://management.core.windows.net/user_impersonation', // Azure Service Management
            'https://graph.microsoft.com/User.Read',                  // Microsoft Graph
            'offline_access'                                          // Refresh token capability
        ];
        this.scope = this.scopes.join(' ');
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
            
            // Start proactive token refresh mechanism
            this.startTokenRefreshTimer();
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
        this.skuContainer = document.getElementById('skuContainer');
        this.skuSelect = document.getElementById('skuSelect');
        this.updateSkuButton = document.getElementById('updateSkuButton');
        this.logoutButton = document.getElementById('logoutButton');

        // Verify all elements were found
        const elements = {
            logArea: this.logArea,
            capacitySelect: this.capacitySelect,
            startButton: this.startButton,
            stopButton: this.stopButton,
            loadingIndicator: this.loadingIndicator,
            debugToggle: this.debugToggle,
            refreshButton: this.refreshButton,
            tenantInfo: this.tenantInfo,
            skuContainer: this.skuContainer,
            skuSelect: this.skuSelect,
            updateSkuButton: this.updateSkuButton,
            logoutButton: this.logoutButton
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
            this.capacitySelect.addEventListener('change', async () => {
                await this.onCapacitySelectionChange();
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

            this.updateSkuButton.addEventListener('click', () => {
                this.updateCapacitySku();
            });

            this.skuSelect.addEventListener('change', () => {
                this.onSkuSelectionChange();
            });

            this.logoutButton.addEventListener('click', async () => {
                await this.handleLogout();
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
                this.updateTenantAndUserDisplay(cachedToken);
                this.log('Using cached authentication token');
                this.debugLog('Cached token is valid');
                
                // Try to extend token lifetime silently in background
                setTimeout(async () => {
                    try {
                        const refreshedTokenInfo = await this.performSilentAuth();
                        if (refreshedTokenInfo && refreshedTokenInfo.accessToken !== cachedToken) {
                            this.accessToken = refreshedTokenInfo.accessToken;
                            await this.cacheToken(refreshedTokenInfo.accessToken, refreshedTokenInfo.expiresIn);
                            this.debugLog('Background token refresh successful');
                        }
                    } catch (error) {
                        this.debugLog('Background token refresh failed: ' + error.message);
                    }
                }, 1000); // Delay to avoid blocking initial load
                
                return cachedToken;
            }

            // Perform OAuth2 flow using launchWebAuthFlow
            const tokenInfo = await this.performOAuth2Flow();
            
            this.accessToken = tokenInfo.accessToken;
            this.updateTenantAndUserDisplay(tokenInfo.accessToken);
            await this.cacheToken(tokenInfo.accessToken, tokenInfo.expiresIn);
            this.log('Authentication successful');
            this.debugLog('New access token obtained and cached');
            
            return tokenInfo.accessToken;
        } catch (error) {
            this.logError('Authentication error', error);
            throw error;
        } finally {
            this.showLoading(false);
        }
    }

    /**
     * Perform OAuth2 authentication flow with fallback for no admin consent
     */
    async performOAuth2Flow() {
        return new Promise(async (resolve, reject) => {
            // Try primary scope first (legacy ARM scope)
            try {
                const tokenInfo = await this.attemptOAuth2WithScope(this.scopes[0]);
                resolve(tokenInfo);
                return;
            } catch (primaryError) {
                this.debugLog(`Authentication failed: ${primaryError.message}`);
                this.logError('Authentication failed. Please ensure you have the required permissions.');
                reject(primaryError);
            }
        });
    }

    /**
     * Attempt OAuth2 authentication with specific scope
     */
    async attemptOAuth2WithScope(scopeString) {
        return new Promise((resolve, reject) => {
            // Azure AD OAuth2 endpoints
            const tenantId = 'common'; // Use 'common' for multi-tenant apps
            const clientId = 'b2f9922d-47b3-45de-be16-72911e143fa4';
            const redirectUri = chrome.identity.getRedirectURL();
            const scope = encodeURIComponent(scopeString);
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

            this.debugLog(`Attempting OAuth2 flow with scope: ${scopeString}`);
            this.debugLog(`Auth URL: ${authUrl}`);
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
                    const tokenInfo = this.extractTokenFromUrl(responseUrl, state);
                    resolve(tokenInfo);
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

        // Return both token and expiry information
        return {
            accessToken: accessToken,
            expiresIn: expiresIn ? parseInt(expiresIn) : 3600, // Default to 1 hour if not provided
            tokenType: tokenType
        };
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
     * Generate a hash of the token for validation purposes (without storing the actual token)
     */
    generateTokenHash(token) {
        // Simple hash function for token validation
        let hash = 0;
        if (token.length === 0) return hash.toString();
        for (let i = 0; i < Math.min(token.length, 100); i++) { // Only hash first 100 chars for performance
            const char = token.charCodeAt(i);
            hash = ((hash << 5) - hash) + char;
            hash = hash & hash; // Convert to 32-bit integer
        }
        return Math.abs(hash).toString();
    }

    /**
     * Get cached authentication token
     */
    async getCachedToken() {
        return new Promise((resolve) => {
            chrome.storage.local.get(['accessToken', 'tokenExpiry', 'sessionInfo'], (result) => {
                if (chrome.runtime.lastError) {
                    this.debugLog('Failed to get cached token: ' + chrome.runtime.lastError.message);
                    resolve(null);
                    return;
                }

                const { accessToken, tokenExpiry, sessionInfo } = result;
                
                if (!accessToken || !tokenExpiry) {
                    this.debugLog('No cached token found');
                    resolve(null);
                    return;
                }

                // Validate session integrity if session info exists
                if (sessionInfo && accessToken) {
                    const currentTokenHash = this.generateTokenHash(accessToken);
                    if (sessionInfo.tokenHash !== currentTokenHash) {
                        this.debugLog('Cached token integrity check failed');
                        resolve(null);
                        return;
                    }
                }

                // Check if token is expired (with 2-minute buffer)
                const now = Date.now();
                const expiryTime = parseInt(tokenExpiry);
                const bufferTime = 2 * 60 * 1000; // 2 minutes in milliseconds
                const timeUntilExpiry = expiryTime - now;
                const timeUntilExpiryMinutes = Math.floor(timeUntilExpiry / (60 * 1000));

                if (now >= (expiryTime - bufferTime)) {
                    this.debugLog(`Cached token is expired or will expire soon (${timeUntilExpiryMinutes} minutes remaining)`);
                    resolve(null);
                    return;
                }

                if (sessionInfo) {
                    const sessionAge = Math.floor((now - sessionInfo.cachedAt) / (60 * 1000));
                    this.debugLog(`Found valid cached token (expires in ${timeUntilExpiryMinutes} minutes, session age: ${sessionAge} minutes)`);
                } else {
                    this.debugLog(`Found valid cached token (expires in ${timeUntilExpiryMinutes} minutes)`);
                }
                
                resolve(accessToken);
            });
        });
    }

    /**
     * Cache authentication token with expiry
     */
    async cacheToken(token, expiresInSeconds) {
        // Use the actual expiry time from Azure AD response
        // Default to 1 hour if not provided
        const expirySeconds = expiresInSeconds || 3600;
        const expiryTime = Date.now() + (expirySeconds * 1000);
        
        // Extract additional session info for better persistence
        const tokenData = this.decodeJwtToken(token);
        const sessionInfo = {
            cachedAt: Date.now(),
            tenantId: tokenData?.tid,
            userPrincipalName: tokenData?.upn || tokenData?.unique_name,
            tokenHash: this.generateTokenHash(token) // For validation without storing full token
        };
        
        this.debugLog(`Caching token with expiry in ${expirySeconds} seconds`);
        
        return new Promise((resolve) => {
            chrome.storage.local.set({
                accessToken: token,
                tokenExpiry: expiryTime.toString(),
                sessionInfo: sessionInfo
            }, () => {
                if (chrome.runtime.lastError) {
                    this.debugLog('Failed to cache token: ' + chrome.runtime.lastError.message);
                } else {
                    this.debugLog('Token cached successfully with session info');
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
            chrome.storage.local.remove(['accessToken', 'tokenExpiry', 'sessionInfo'], () => {
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
     * Handle logout button click - clear authentication and reset UI
     */
    async handleLogout() {
        try {
            this.log('Logging out...');
            
            // Clear authentication data
            await this.clearCachedAuth();
            this.accessToken = null;
            
            // Clear capacities and reset UI
            this.capacities = [];
            this.initialLoadComplete = false;
            
            // Reset UI elements
            this.populateCapacityDropdown();
            this.skuContainer.style.display = 'none';
            this.startButton.disabled = true;
            this.stopButton.disabled = true;
            this.updateSkuButton.disabled = true;
            
            // Remove any auth warning
            const existingWarning = document.getElementById('auth-warning');
            if (existingWarning) {
                existingWarning.remove();
            }
            
            this.log('Logged out successfully. Click refresh or select a capacity to re-authenticate.');
            
        } catch (error) {
            this.logError('Failed to logout', error);
        }
    }

    /**
     * Handle token expiry with automatic silent refresh attempt
     */
    async handleTokenExpiry() {
        try {
            this.debugLog('Attempting silent token refresh...');
            
            // Try to perform a silent authentication
            const tokenInfo = await this.performSilentAuth();
            if (tokenInfo) {
                this.accessToken = tokenInfo.accessToken;
                await this.cacheToken(tokenInfo.accessToken, tokenInfo.expiresIn);
                this.debugLog('Silent token refresh successful');
                return tokenInfo.accessToken;
            }
        } catch (error) {
            this.debugLog('Silent token refresh failed: ' + error.message);
        }
        
        // If silent refresh fails, clear the token and require re-authentication
        this.accessToken = null;
        await this.clearCachedAuth();
        return null;
    }

    /**
     * Attempt silent authentication (non-interactive) with fallback scopes
     */
    async performSilentAuth() {
        // Try primary scope first (legacy ARM scope)
        try {
            const tokenInfo = await this.attemptSilentAuthWithScope(this.scopes[0]);
            return tokenInfo;
        } catch (primaryError) {
            this.debugLog(`Primary silent auth failed: ${primaryError.message}`);
            
            // If primary scope fails, try with basic scopes only
            try {
                this.debugLog('Attempting fallback silent authentication with basic scopes...');
                const fallbackScopes = ['openid', 'profile', 'offline_access'];
                const tokenInfo = await this.attemptSilentAuthWithScope(fallbackScopes.join(' '));
                return tokenInfo;
            } catch (fallbackError) {
                this.debugLog(`Fallback silent auth failed: ${fallbackError.message}`);
                return null;
            }
        }
    }

    /**
     * Attempt silent authentication with specific scope
     */
    async attemptSilentAuthWithScope(scopeString) {
        return new Promise((resolve, reject) => {
            // Azure AD OAuth2 endpoints for silent auth
            const tenantId = 'common';
            const clientId = 'b2f9922d-47b3-45de-be16-72911e143fa4';
            const redirectUri = chrome.identity.getRedirectURL();
            const scope = encodeURIComponent(scopeString);
            const responseType = 'token';
            const state = this.generateState();

            // Construct authorization URL with prompt=none for silent auth
            const authUrl = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/authorize` +
                `?client_id=${clientId}` +
                `&response_type=${responseType}` +
                `&redirect_uri=${encodeURIComponent(redirectUri)}` +
                `&scope=${scope}` +
                `&state=${state}` +
                `&response_mode=fragment` +
                `&prompt=none`;

            this.debugLog(`Attempting silent auth with scope: ${scopeString}`);

            chrome.identity.launchWebAuthFlow({
                url: authUrl,
                interactive: false  // Non-interactive for silent auth
            }, (responseUrl) => {
                if (chrome.runtime.lastError) {
                    this.debugLog('Silent auth failed: ' + chrome.runtime.lastError.message);
                    reject(new Error(chrome.runtime.lastError.message));
                    return;
                }

                if (!responseUrl) {
                    this.debugLog('No response URL from silent auth');
                    reject(new Error('No response URL from silent auth'));
                    return;
                }

                try {
                    const tokenInfo = this.extractTokenFromUrl(responseUrl, state);
                    resolve(tokenInfo);
                } catch (error) {
                    this.debugLog('Failed to extract token from silent auth response: ' + error.message);
                    reject(error);
                }
            });
        });
    }

    /**
     * Start proactive token refresh timer
     */
    startTokenRefreshTimer() {
        // Check token status every 10 minutes
        const refreshInterval = setInterval(async () => {
            try {
                const cachedToken = await this.getCachedToken();
                if (!cachedToken) {
                    // No valid token, stop the timer
                    this.debugLog('No valid token found, stopping proactive refresh timer');
                    clearInterval(refreshInterval);
                    return;
                }
                
                // Get current token expiry info
                const result = await new Promise((resolve) => {
                    chrome.storage.local.get(['tokenExpiry', 'sessionInfo'], resolve);
                });
                
                if (result.tokenExpiry) {
                    const expiryTime = parseInt(result.tokenExpiry);
                    const now = Date.now();
                    const timeUntilExpiry = expiryTime - now;
                    const minutesUntilExpiry = Math.floor(timeUntilExpiry / (60 * 1000));
                    
                    // If token expires in less than 15 minutes, try to refresh
                    if (minutesUntilExpiry < 15 && minutesUntilExpiry > 0) {
                        this.debugLog(`Token expires in ${minutesUntilExpiry} minutes, attempting proactive refresh`);
                        
                        const refreshedTokenInfo = await this.performSilentAuth();
                        if (refreshedTokenInfo) {
                            this.accessToken = refreshedTokenInfo.accessToken;
                            await this.cacheToken(refreshedTokenInfo.accessToken, refreshedTokenInfo.expiresIn);
                            this.debugLog('Proactive token refresh successful');
                            
                            // Update tenant display with new token
                            this.updateTenantAndUserDisplay(refreshedTokenInfo.accessToken);
                        } else {
                            this.debugLog('Proactive token refresh failed, user will need to re-authenticate');
                        }
                    } else if (minutesUntilExpiry > 0) {
                        this.debugLog(`Token still valid for ${minutesUntilExpiry} minutes`);
                    }
                }
            } catch (error) {
                this.debugLog('Proactive token refresh check failed: ' + error.message);
            }
        }, 10 * 60 * 1000); // Every 10 minutes
        
        this.debugLog('Started proactive token refresh timer (checks every 10 minutes)');
    }

    /**
     * Load all Fabric capacities across subscriptions
     */
    async loadCapacities() {
        try {
            this.log('Loading Fabric capacities...');
            this.showLoading(true);
            this.capacities = [];

            // Check if we need to authenticate first
            if (!this.accessToken) {
                this.log('No authentication token available. Please authenticate first.');
                await this.authenticate();
            }

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
        // Check if we need to authenticate first (e.g., after logout)
        if (!this.accessToken) {
            this.debugLog('No authentication token available, triggering full load with authentication');
            await this.loadCapacities();
            return;
        }

        // Don't refresh if we're already loading
        if (this.loadingIndicator.style.display === 'block') {
            this.debugLog('Refresh skipped - already loading');
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
            this.skuContainer.style.display = 'none';
            
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
                    await this.onCapacitySelectionChange();
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
                
                const sku = capacity.sku?.name || 'Unknown SKU';
                option.textContent = `${capacity.name} - ${sku} ${stateDisplay}`;
                
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
    async onCapacitySelectionChange() {
        const selectedIndex = this.capacitySelect.value;
        
        if (!selectedIndex || selectedIndex === '') {
            this.startButton.disabled = true;
            this.stopButton.disabled = true;
            this.skuContainer.style.display = 'none';
            return;
        }

        const capacity = this.capacities[parseInt(selectedIndex)];
        const state = capacity.properties?.state || 'Unknown';

        // Enable/disable buttons based on current state
        this.startButton.disabled = (state === 'Active');
        this.stopButton.disabled = (state === 'Paused');

        // Show SKU container and load available SKUs
        this.skuContainer.style.display = 'block';
        await this.loadAvailableSkus(capacity);
        await this.populateSkuDropdown(capacity);

        this.log(`Selected capacity: ${capacity.name} (${state}) - SKU: ${capacity.sku?.name || 'Unknown'}`);
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
            await this.setButtonsEnabled(false);

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
            await this.setButtonsEnabled(true);
        }
    }

    /**
     * Load available SKUs for Fabric capacities
     */
    async loadAvailableSkus(capacity) {
        try {
            this.debugLog('Loading available SKUs...');
            
            // Common Fabric capacity SKUs based on Azure documentation
            // These are the standard SKUs available for Microsoft Fabric
            this.availableSkus = [
                { name: 'F2', displayName: 'F2 (2 vCores, 4 GB RAM)', description: 'Small workloads' },
                { name: 'F4', displayName: 'F4 (4 vCores, 8 GB RAM)', description: 'Development and testing' },
                { name: 'F8', displayName: 'F8 (8 vCores, 16 GB RAM)', description: 'Light production workloads' },
                { name: 'F16', displayName: 'F16 (16 vCores, 32 GB RAM)', description: 'Medium production workloads' },
                { name: 'F32', displayName: 'F32 (32 vCores, 64 GB RAM)', description: 'Heavy production workloads' },
                { name: 'F64', displayName: 'F64 (64 vCores, 128 GB RAM)', description: 'Enterprise workloads' },
                { name: 'F128', displayName: 'F128 (128 vCores, 256 GB RAM)', description: 'Large enterprise workloads' },
                { name: 'F256', displayName: 'F256 (256 vCores, 512 GB RAM)', description: 'Very large enterprise workloads' },
                { name: 'F512', displayName: 'F512 (512 vCores, 1024 GB RAM)', description: 'Maximum capacity workloads' }
            ];

            // Alternatively, we could fetch available SKUs from Azure API
            // const subscriptionId = capacity.subscriptionId;
            // const location = capacity.location;
            // const url = `${this.baseUrl}/subscriptions/${subscriptionId}/providers/Microsoft.Fabric/locations/${location}/skus?api-version=${this.fabricApiVersion}`;
            // const response = await this.makeApiCall(url);
            // this.availableSkus = response.value || [];

            this.debugLog(`Loaded ${this.availableSkus.length} available SKUs`);
        } catch (error) {
            this.logError('Failed to load available SKUs', error);
            // Use default SKUs if API call fails
            this.availableSkus = [
                { name: 'F2', displayName: 'F2', description: 'Small' },
                { name: 'F4', displayName: 'F4', description: 'Medium' },
                { name: 'F8', displayName: 'F8', description: 'Large' }
            ];
        }
    }

    /**
     * Populate the SKU dropdown with available options
     */
    async populateSkuDropdown(capacity) {
        try {
            const currentSku = capacity.sku?.name || 'Unknown';
            
            // Clear existing options
            this.skuSelect.innerHTML = '';
            
            // Add current SKU as selected option
            const currentOption = document.createElement('option');
            currentOption.value = currentSku;
            currentOption.textContent = `Current: ${currentSku}`;
            currentOption.selected = true;
            this.skuSelect.appendChild(currentOption);
            
            // Add separator
            const separator = document.createElement('option');
            separator.disabled = true;
            separator.textContent = '─────────────────';
            this.skuSelect.appendChild(separator);
            
            // Add available SKUs
            this.availableSkus.forEach(sku => {
                if (sku.name !== currentSku) {
                    const option = document.createElement('option');
                    option.value = sku.name;
                    option.textContent = sku.displayName || sku.name;
                    option.title = sku.description || '';
                    this.skuSelect.appendChild(option);
                }
            });
            
            // Reset update button state
            this.updateSkuButton.disabled = true;
            
            this.debugLog(`Populated SKU dropdown with current SKU: ${currentSku}`);
        } catch (error) {
            this.logError('Failed to populate SKU dropdown', error);
        }
    }

    /**
     * Handle SKU selection change
     */
    onSkuSelectionChange() {
        const selectedSku = this.skuSelect.value;
        const selectedIndex = this.capacitySelect.value;
        
        if (!selectedIndex || selectedIndex === '') return;
        
        const capacity = this.capacities[parseInt(selectedIndex)];
        const currentSku = capacity.sku?.name || 'Unknown';
        
        // Enable update button only if a different SKU is selected
        this.updateSkuButton.disabled = (selectedSku === currentSku || selectedSku.startsWith('Current:'));
        
        if (selectedSku !== currentSku && !selectedSku.startsWith('Current:')) {
            this.debugLog(`SKU change selected: ${currentSku} → ${selectedSku}`);
        }
    }

    /**
     * Update the capacity SKU
     */
    async updateCapacitySku() {
        const selectedIndex = this.capacitySelect.value;
        const selectedSku = this.skuSelect.value;
        
        if (!selectedIndex || selectedIndex === '' || !selectedSku) return;
        
        const capacity = this.capacities[parseInt(selectedIndex)];
        const currentSku = capacity.sku?.name || 'Unknown';
        
        if (selectedSku === currentSku || selectedSku.startsWith('Current:')) {
            this.log('No SKU change required');
            return;
        }
        
        // Check if capacity is running - SKU changes typically require the capacity to be stopped
        const state = capacity.properties?.state || 'Unknown';
        if (state === 'Active') {
            const proceed = confirm(
                `Changing the SKU typically requires stopping the capacity first.\n\n` +
                `Current: ${currentSku}\nNew: ${selectedSku}\n\n` +
                `Do you want to proceed? The capacity may be temporarily unavailable.`
            );
            if (!proceed) return;
        }
        
        try {
            this.log(`Updating capacity ${capacity.name} SKU from ${currentSku} to ${selectedSku}...`);
            await this.setButtonsEnabled(false);
            this.updateSkuButton.disabled = true;
            this.skuSelect.disabled = true;

            // Prepare the update payload
            const updatePayload = {
                sku: {
                    name: selectedSku
                }
            };

            const url = `${this.baseUrl}${capacity.id}?api-version=${this.fabricApiVersion}`;
            
            // Use PATCH method to update the capacity
            await this.makeApiCall(url, 'PATCH', updatePayload);
            
            this.logSuccess(`SKU update initiated for ${capacity.name}: ${currentSku} → ${selectedSku}`);
            
            // Refresh the capacity list to update SKU information after a delay
            setTimeout(async () => {
                await this.refreshCapacities();
                // Re-select the same capacity to update the SKU dropdown
                if (this.capacitySelect.value) {
                    await this.onCapacitySelectionChange();
                }
            }, 3000);

        } catch (error) {
            this.logError(`Failed to update SKU for capacity ${capacity.name}`, error);
        } finally {
            await this.setButtonsEnabled(true);
            this.skuSelect.disabled = false;
            // Update button will be re-enabled by onSkuSelectionChange if needed
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
                // Token might be expired, try silent refresh first
                this.debugLog('Token expired, attempting silent refresh');
                
                const refreshedToken = await this.handleTokenExpiry();
                
                if (refreshedToken) {
                    // Retry with refreshed token
                    options.headers['Authorization'] = `Bearer ${refreshedToken}`;
                    const retryResponse = await fetch(url, options);
                    
                    if (!retryResponse.ok) {
                        const retryErrorText = await retryResponse.text();
                        this.debugLog(`Retry failed: ${retryResponse.status} ${retryResponse.statusText} - ${retryErrorText}`);
                        
                        if (retryResponse.status === 403) {
                            throw new Error(`Insufficient permissions. Admin consent may be required for full functionality.`);
                        }
                        
                        throw new Error(`API call failed after retry: ${retryResponse.status} ${retryResponse.statusText}`);
                    }
                    
                    // Handle successful retry response
                    if (method === 'POST' && retryResponse.status === 202) {
                        return { success: true };
                    }
                    if (method === 'PATCH' && (retryResponse.status === 200 || retryResponse.status === 202)) {
                        return { success: true };
                    }
                    
                    const retryResponseText = await retryResponse.text();
                    return retryResponseText ? JSON.parse(retryResponseText) : {};
                } else {
                    // Silent refresh failed, require interactive authentication
                    this.debugLog('Silent refresh failed, requiring interactive authentication');
                    await this.authenticate();
                    
                    // Retry the original request with new token
                    options.headers['Authorization'] = `Bearer ${this.accessToken}`;
                    const retryResponse = await fetch(url, options);
                    
                    if (!retryResponse.ok) {
                        const retryErrorText = await retryResponse.text();
                        this.debugLog(`Final retry failed: ${retryResponse.status} ${retryResponse.statusText} - ${retryErrorText}`);
                        
                        if (retryResponse.status === 403) {
                            throw new Error(`Insufficient permissions. Admin consent may be required for full functionality.`);
                        }
                        
                        throw new Error(`API call failed after authentication retry: ${retryResponse.status} ${retryResponse.statusText}`);
                    }
                    
                    // Handle successful retry response
                    if (method === 'POST' && retryResponse.status === 202) {
                        return { success: true };
                    }
                    if (method === 'PATCH' && (retryResponse.status === 200 || retryResponse.status === 202)) {
                        return { success: true };
                    }
                    
                    const retryResponseText = await retryResponse.text();
                    return retryResponseText ? JSON.parse(retryResponseText) : {};
                }
            }
            
            if (response.status === 403) {
                // Check if this is a permission issue vs admin consent issue
                let errorMessage = 'Insufficient permissions to perform this operation.';
                
                if (errorText.includes('admin consent') || errorText.includes('consent')) {
                    errorMessage += ' Admin consent may be required. Please contact your Azure administrator.';
                } else if (errorText.includes('RBAC') || errorText.includes('role')) {
                    errorMessage += ' You may need additional Azure RBAC permissions.';
                }
                
                throw new Error(errorMessage);
            }
            
            throw new Error(`API call failed: ${response.status} ${response.statusText}`);
        }

        // Handle empty responses for POST and PATCH operations
        if ((method === 'POST' || method === 'PATCH') && (response.status === 200 || response.status === 202)) {
            return { success: true };
        }

        const responseText = await response.text();
        return responseText ? JSON.parse(responseText) : {};
    }

    /**
     * Enable/disable operation buttons
     */
    async setButtonsEnabled(enabled) {
        if (enabled) {
            await this.onCapacitySelectionChange(); // Re-evaluate button states
            this.refreshButton.disabled = false;
        } else {
            this.startButton.disabled = true;
            this.stopButton.disabled = true;
            this.refreshButton.disabled = true;
            // Note: updateSkuButton is handled separately in updateCapacitySku method
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
     * Check token capabilities and provide user guidance
     */
    async checkTokenCapabilities() {
        if (!this.accessToken) {
            return { hasManagementAccess: false, message: 'No authentication token available' };
        }

        try {
            // Test basic subscription access (this usually works with basic scopes)
            const testUrl = `${this.baseUrl}/subscriptions?api-version=${this.subscriptionApiVersion}&$top=1`;
            const response = await fetch(testUrl, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${this.accessToken}`,
                    'Content-Type': 'application/json'
                }
            });

            if (response.ok) {
                this.debugLog('Token has management API access');
                return { 
                    hasManagementAccess: true, 
                    message: 'Full Azure Management API access available' 
                };
            } else if (response.status === 403) {
                this.debugLog('Token has limited permissions');
                return { 
                    hasManagementAccess: false, 
                    message: 'Limited permissions - Admin consent may be required for full functionality' 
                };
            } else {
                this.debugLog(`Token validation failed: ${response.status}`);
                return { 
                    hasManagementAccess: false, 
                    message: `Authentication issue: ${response.status} ${response.statusText}` 
                };
            }
        } catch (error) {
            this.debugLog('Token capability check failed: ' + error.message);
            return { 
                hasManagementAccess: false, 
                message: 'Unable to verify token capabilities' 
            };
        }
    }

    /**
     * Get user information from Microsoft Graph
     */
    async getUserInfo() {
        try {
            if (!this.accessToken) {
                this.debugLog('No access token available for Graph API call');
                return null;
            }

            const response = await fetch(`${this.graphUrl}/v1.0/me`, {
                method: 'GET',
                headers: {
                    'Authorization': `Bearer ${this.accessToken}`,
                    'Content-Type': 'application/json'
                }
            });

            if (!response.ok) {
                this.debugLog(`Graph API call failed: ${response.status} ${response.statusText}`);
                return null;
            }

            const userInfo = await response.json();
            this.debugLog('User info retrieved from Graph:', JSON.stringify({
                displayName: userInfo.displayName,
                userPrincipalName: userInfo.userPrincipalName,
                givenName: userInfo.givenName,
                surname: userInfo.surname
            }));

            return userInfo;
        } catch (error) {
            this.debugLog(`Failed to get user info from Graph: ${error.message}`);
            return null;
        }
    }

    /**
     * Update tenant and user display with Graph API data
     */
    async updateTenantAndUserDisplay(token) {
        try {
            // First update with token data (synchronous)
            this.updateTenantDisplayFromToken(token);
            
            // Then enhance with Graph API data (asynchronous)
            const userInfo = await this.getUserInfo();
            if (userInfo) {
                const tokenData = this.decodeJwtToken(token);
                const tenantName = tokenData?.tenant_display_name || tokenData?.tenant_name;
                const userPrincipalName = tokenData?.upn || tokenData?.unique_name;
                
                // Create display with user name and tenant
                let displayText = '';
                if (userInfo.displayName) {
                    displayText = `${userInfo.displayName}`;
                    if (tenantName) {
                        displayText += ` (${tenantName})`;
                    } else if (userPrincipalName) {
                        const domain = userPrincipalName.split('@')[1];
                        if (domain) {
                            displayText += ` (${domain})`;
                        }
                    }
                } else if (tenantName) {
                    displayText = `Tenant: ${tenantName}`;
                } else if (userPrincipalName) {
                    const domain = userPrincipalName.split('@')[1];
                    displayText = domain ? `Tenant: ${domain}` : userPrincipalName;
                }

                if (displayText) {
                    this.tenantInfo.textContent = displayText;
                    this.tenantInfo.style.display = 'block';
                    this.debugLog(`User and tenant display updated: ${displayText}`);
                }
            }

            // Check token capabilities and update UI accordingly
            this.checkTokenCapabilities().then(result => {
                if (!result.hasManagementAccess) {
                    this.log(`⚠️  ${result.message}`);
                    
                    // Add a subtle warning to the UI
                    const existingWarning = document.getElementById('auth-warning');
                    if (!existingWarning) {
                        const warning = document.createElement('div');
                        warning.id = 'auth-warning';
                        warning.style.cssText = `
                            background: #fff4ce;
                            border: 1px solid #fed100;
                            border-radius: 4px;
                            padding: 8px;
                            margin-bottom: 12px;
                            font-size: 12px;
                            color: #323130;
                        `;
                        warning.innerHTML = `
                            <strong>Limited Access:</strong> Some features may require admin consent.<br>
                            <small>Contact your administrator if you experience permission errors.</small>
                        `;
                        
                        const container = document.querySelector('.container');
                        if (container) {
                            container.insertBefore(warning, container.firstChild);
                        }
                    }
                } else {
                    // Remove warning if it exists and we have full access
                    const existingWarning = document.getElementById('auth-warning');
                    if (existingWarning) {
                        existingWarning.remove();
                    }
                }
            });
        } catch (error) {
            this.debugLog('Failed to update tenant and user display: ' + error.message);
            // Fallback to token-only display
            this.updateTenantDisplayFromToken(token);
        }
    }

    /**
     * Update tenant display from token data only (fallback method)
     */
    updateTenantDisplayFromToken(token) {
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
            this.debugLog('Failed to update tenant display from token: ' + error.message);
        }
    }
}

// Initialize the extension when the popup loads
document.addEventListener('DOMContentLoaded', () => {
    const manager = new FabricCapacityManager();
    manager.init();
});
