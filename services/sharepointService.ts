import * as AuthSession from 'expo-auth-session';
import * as WebBrowser from 'expo-web-browser';

// Complete the auth session
WebBrowser.maybeCompleteAuthSession();

interface SharePointConfig {
  siteUrl: string;
  clientId: string;
  tenantId: string;
}

interface ListItem {
  [key: string]: any;
}

/**
 * Custom error class for session timeout/expired token scenarios
 */
export class SessionTimeoutError extends Error {
  constructor(message: string = 'Your session has expired. Please log in again.') {
    super(message);
    this.name = 'SessionTimeoutError';
    Object.setPrototypeOf(this, SessionTimeoutError.prototype);
  }
}

class SharePointService {
  private siteUrl: string;
  private clientId: string;
  private tenantId: string;
  private accessToken: string | null = null;
  private sharePointRoot: string = '';
  private siteId: string | null = null;
  private listIdCache: Map<string, string> = new Map();
  private fieldMappingCache: Map<string, Map<string, string>> = new Map(); // listName -> (internalName -> displayName)
  private userDisplayNameCache: Map<string | number, string> = new Map(); // SharePoint user ID -> display name
  private userInformationListId: string | null = null; // Cache for User Information List ID
  private onSessionTimeoutCallback: (() => void) | null = null;

  constructor(config: SharePointConfig) {
    this.siteUrl = config.siteUrl;
    this.clientId = config.clientId;
    this.tenantId = config.tenantId;
  }

  /**
   * Authenticate with SharePoint using OAuth 2.0
   */
  async authenticate(): Promise<string> {
    try {
      if (!this.clientId || this.clientId === 'YOUR_CLIENT_ID_HERE') {
        throw new Error(
          'Client ID not configured. Please update config/sharepointConfig.ts with your Azure AD Client ID.'
        );
      }

      if (!this.tenantId || this.tenantId === 'YOUR_TENANT_ID_HERE') {
        throw new Error(
          'Tenant ID not configured. Please update config/sharepointConfig.ts with your Azure AD Tenant ID.'
        );
      }

      // Use a fixed custom scheme URI that doesn't depend on IP address
      // Note: Custom schemes work in development/production builds, but Expo Go will fall back to exp://
      let redirectUri = AuthSession.makeRedirectUri({
        scheme: 'employee-assets',
        path: 'auth',
      });
      
      // Fallback: If still using exp:// scheme (Expo Go limitation), use localhost which is more stable
      if (redirectUri.startsWith('exp://')) {
        // Try to use localhost instead of IP for stability
        const localhostUri = AuthSession.makeRedirectUri({
          preferLocalhost: true,
        });
        if (localhostUri && localhostUri.includes('localhost')) {
          redirectUri = localhostUri.replace(/\/--\/.*$/, '/'); // Remove path, keep base URI
        } else {
          // Last resort: use fixed custom scheme format
          redirectUri = 'employee-assets://auth';
        }
      }
      
      // Ensure the redirect URI is in the correct format
      if (!redirectUri.includes('://')) {
        redirectUri = 'employee-assets://auth';
      }
      
      const sharePointRoot = this.siteUrl.split('/sites/')[0];
      this.sharePointRoot = sharePointRoot;
      
      const scopes = [
        'https://graph.microsoft.com/Sites.ReadWrite.All',
        'https://graph.microsoft.com/User.Read',
      ];
      
      const discovery = {
        authorizationEndpoint: `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/authorize`,
        tokenEndpoint: `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`,
      };

      const request = new AuthSession.AuthRequest({
        clientId: this.clientId,
        scopes: scopes,
        responseType: AuthSession.ResponseType.Code,
        redirectUri,
        usePKCE: true,
      });

      const result = await request.promptAsync(discovery);

      if (result.type === 'success') {
        if (!result.params.code) {
          throw new Error('Authorization code not received');
        }
        
        const codeVerifier = (request as any).codeVerifier || 
                            (request as any)._codeVerifier ||
                            (request as any).code_verifier;
        
        const tokenExchangeConfig: any = {
          clientId: this.clientId,
          redirectUri: redirectUri,
          code: result.params.code,
          extraParams: {},
        };
        
        if (codeVerifier) {
          tokenExchangeConfig.extraParams.code_verifier = codeVerifier;
        }
        
        const tokenResult = await AuthSession.exchangeCodeAsync(
          tokenExchangeConfig,
          discovery
        );
        
        if (!tokenResult.accessToken) {
          throw new Error('Access token not received from token exchange');
        }
        
        this.accessToken = tokenResult.accessToken;
        this.siteId = null;
        this.listIdCache.clear();
        
        return this.accessToken;
      } else if (result.type === 'error') {
        const errorMessage = result.error?.message || result.error?.code || 'Unknown error';
        const errorDescription = result.error?.description || '';
        
        // Check if it's a redirect URI mismatch error
        if (errorMessage.includes('redirect_uri') || errorDescription.includes('redirect_uri') || 
            errorMessage.includes('redirect') || errorDescription.includes('redirect')) {
          throw new Error(
            `Redirect URI mismatch!\n\n` +
            `The redirect URI used: ${redirectUri}\n\n` +
            `Please add this exact URI to Azure AD:\n` +
            `1. Go to Azure Portal → Your App → Authentication\n` +
            `2. Under "Platform configurations", add "Mobile and desktop applications"\n` +
            `3. Add this exact redirect URI: ${redirectUri}\n` +
            `4. Save and try again.\n\n` +
            `Error details: ${errorMessage}${errorDescription ? ' - ' + errorDescription : ''}`
          );
        }
        
        throw new Error(`Authentication error: ${errorMessage}${errorDescription ? ' - ' + errorDescription : ''}`);
      } else if (result.type === 'cancel') {
        throw new Error('Authentication cancelled by user');
      } else {
        throw new Error(`Authentication failed: ${result.type}`);
      }
    } catch (error: any) {
      if (error.message) {
        throw error;
      }
      throw new Error(`Authentication failed: ${error.toString()}`);
    }
  }

  setAccessToken(token: string) {
    this.accessToken = token;
  }

  getAccessToken(): string | null {
    return this.accessToken;
  }

  /**
   * Set callback to be called when session timeout is detected
   */
  setOnSessionTimeout(callback: () => void) {
    this.onSessionTimeoutCallback = callback;
  }

  /**
   * Clear session timeout callback
   */
  clearOnSessionTimeout() {
    this.onSessionTimeoutCallback = null;
  }

  /**
   * Handle session timeout - clear token and notify callback
   */
  private handleSessionTimeout() {
    this.accessToken = null;
    this.siteId = null;
    this.listIdCache.clear();
    this.fieldMappingCache.clear();
    
    if (this.onSessionTimeoutCallback) {
      this.onSessionTimeoutCallback();
    }
  }

  /**
   * Make authenticated request to Microsoft Graph API
   */
  private async makeGraphRequest(
    endpoint: string,
    options: RequestInit = {}
  ): Promise<any> {
    if (!this.accessToken) {
      throw new Error('Not authenticated. Call authenticate() first.');
    }

    const url = `https://graph.microsoft.com/v1.0/${endpoint}`;
    const response = await fetch(url, {
      ...options,
      headers: {
        ...options.headers,
        Authorization: `Bearer ${this.accessToken}`,
        Accept: 'application/json',
        'Content-Type': 'application/json',
      },
    });

    // Handle successful DELETE requests (204 No Content) - return early
    // Also handle 200 status with empty body for DELETE operations
    const isDeleteRequest = options.method === 'DELETE';
    if (response.status === 204 || (isDeleteRequest && response.status === 200)) {
      return null; // Success, but no content to return
    }

    if (!response.ok) {
      let errorText = '';
      try {
        errorText = await response.text();
      } catch (textError) {
        // If we can't read error text, use status text
        errorText = response.statusText || `HTTP ${response.status}`;
      }
      
      // Check for 401 Unauthorized (session timeout/expired token)
      if (response.status === 401) {
        // Handle session timeout
        this.handleSessionTimeout();
        throw new SessionTimeoutError('Your session has expired. Please log in again.');
      }
      
      // Check for 403 Forbidden (could also indicate expired token in some cases)
      if (response.status === 403) {
        // Try to parse error to see if it's token-related
        try {
          if (errorText) {
            const errorJson = JSON.parse(errorText);
            const errorCode = errorJson.error?.code || '';
            const errorMessage = errorJson.error?.message || '';
            
            // Check for token expiration indicators
            if (
              errorCode.includes('InvalidAuthenticationToken') ||
              errorCode.includes('AuthenticationTokenExpired') ||
              errorMessage.includes('token') ||
              errorMessage.includes('expired') ||
              errorMessage.includes('authentication')
            ) {
              this.handleSessionTimeout();
              throw new SessionTimeoutError('Your session has expired. Please log in again.');
            }
          }
        } catch (parseError) {
          // If error parsing fails, continue with generic error
        }
      }
      
      throw new Error(`Microsoft Graph API error: ${response.status} - ${errorText}`);
    }

    // For DELETE requests, even if status is 200, return null (no content expected)
    if (isDeleteRequest) {
      return null;
    }

    // Check if response has content before trying to parse JSON
    const contentType = response.headers.get('content-type');
    const contentLength = response.headers.get('content-length');
    
    // If content-length is 0, return null without trying to parse
    if (contentLength === '0') {
      return null;
    }

    // Try to parse JSON, but handle empty responses gracefully
    try {
      const text = await response.text();
      if (!text || text.trim() === '') {
        return null;
      }
      return JSON.parse(text);
    } catch (parseError: any) {
      // If JSON parsing fails but status is OK, return null (empty response)
      // This handles cases where response is successful but body is empty/invalid JSON
      if (response.ok) {
        return null;
      }
      // If status is not OK, rethrow the original error
      throw parseError;
    }
  }

  /**
   * Get SharePoint site ID from Microsoft Graph API
   */
  private async getSiteId(): Promise<string> {
    if (this.siteId) {
      return this.siteId;
    }

    const urlObj = new URL(this.siteUrl);
    const hostname = urlObj.hostname;
    const pathParts = urlObj.pathname.split('/').filter(p => p);
    
    let sitePath = hostname;
    if (pathParts.length > 0) {
      sitePath += ':/' + pathParts.join('/');
    }
    
    const encodedSitePath = encodeURIComponent(sitePath);
    const data = await this.makeGraphRequest(`sites/${encodedSitePath}`);
    
    if (!data.id) {
      throw new Error('Site ID not found in response');
    }
    
    this.siteId = String(data.id);
    return this.siteId;
  }

  /**
   * Get list ID by list name
   */
  private async getListId(listName: string): Promise<string> {
    if (this.listIdCache.has(listName)) {
      return this.listIdCache.get(listName)!;
    }

    const siteId = await this.getSiteId();
    const listsData = await this.makeGraphRequest(`sites/${siteId}/lists`);
    const lists = listsData.value || [];
    
    const list = lists.find((l: any) => 
      l.displayName?.toLowerCase() === listName.toLowerCase() || 
      l.name?.toLowerCase() === listName.toLowerCase()
    );

    if (!list) {
      const availableLists = lists.map((l: any) => l.displayName || l.name).join(', ');
      throw new Error(
        `List "${listName}" not found.\n` +
        `Available lists: ${availableLists || 'none'}`
      );
    }

    this.listIdCache.set(listName, list.id);
    return list.id;
  }

  /**
   * Get all lists in the SharePoint site
   */
  async getLists(): Promise<any[]> {
    const siteId = await this.getSiteId();
    const response = await this.makeGraphRequest(`sites/${siteId}/lists`);
    return response.value || [];
  }

  /**
   * Get list by name
   */
  async getList(listName: string): Promise<any> {
    const listId = await this.getListId(listName);
    const siteId = await this.getSiteId();
    return await this.makeGraphRequest(`sites/${siteId}/lists/${listId}`);
  }

  /**
   * Get field definitions for a list and create mapping from internal names to display names
   * Also returns lookup field names for expansion
   */
  private async getFieldMapping(listName: string): Promise<{
    mapping: Map<string, string>;
    lookupFields: string[];
  }> {
    // Check cache first - we'll need to update cache structure
    const cacheKey = `${listName}_fieldInfo`;
    
    try {
      const listId = await this.getListId(listName);
      const siteId = await this.getSiteId();
      
      // Fetch columns/fields from SharePoint
      const columnsResponse = await this.makeGraphRequest(
        `sites/${siteId}/lists/${listId}/columns`
      );
      
      const columns = columnsResponse.value || [];
      const mapping = new Map<string, string>();
      const lookupFields: string[] = [];
      
      // Create mapping: internal name -> display name
      columns.forEach((column: any) => {
        const internalName = column.name || column.internalName;
        const displayName = column.displayName || column.name;
        
        if (internalName && displayName) {
          mapping.set(internalName, displayName);
          // Also map common variations
          if (internalName.includes('LookupId')) {
            const baseName = internalName.replace('LookupId', '');
            mapping.set(baseName, displayName);
          }
          
          // Identify lookup fields (Person or Lookup column types)
          // Check various ways the column type might be indicated
          const columnType = column.text?.type || column.type || column.columnType || '';
          const isPersonField = columnType === 'person' || 
                               columnType === 'user' ||
                               column.person ||
                               column.user ||
                               (column.text && column.text.type === 'person');
          const isLookupField = columnType === 'lookup' || 
                               column.lookup ||
                               (column.text && column.text.type === 'lookup');
          
          // If it ends with LookupId, it's definitely a lookup field
          const hasLookupIdSuffix = internalName.endsWith('LookupId');
          
          if (isPersonField || isLookupField || hasLookupIdSuffix) {
            // Get the base field name (without LookupId suffix)
            const baseFieldName = internalName.replace('LookupId', '');
            if (baseFieldName && !lookupFields.includes(baseFieldName)) {
              lookupFields.push(baseFieldName);
            }
          }
        }
      });
      
      // Cache the mapping (store both mapping and lookupFields)
      this.fieldMappingCache.set(listName, mapping);
      
      return { mapping, lookupFields };
    } catch (error: any) {
      // Return empty map if fetch fails
      const emptyMap = new Map<string, string>();
      this.fieldMappingCache.set(listName, emptyMap);
      return { mapping: emptyMap, lookupFields: [] };
    }
  }

  /**
   * Normalize field names in a record: remove duplicates when both internal and display names exist
   * Only uses keys from SharePoint response, doesn't create new keys
   */
  private normalizeFieldNames(record: any, fieldMapping: Map<string, string>): any {
    const normalized: any = { ...record }; // Start with all original fields
    
    // Track which display names to remove (when internal name exists with same value)
    const displayNamesToRemove = new Set<string>();
    
    // Process all fields - only remove duplicates, don't create new keys
    Object.keys(record).forEach((key) => {
      // Check if this is an internal field name that has a display name mapping
      const displayName = fieldMapping.get(key);
      
      if (displayName && displayName !== key) {
        const value = record[key];
        const displayValue = record[displayName];
        
        // If display name already exists in original record with same value, mark it for removal
        // (prefer internal name since code uses internal names)
        if (displayValue !== undefined) {
          // Compare values (handle objects, arrays, primitives)
          const valuesMatch = JSON.stringify(displayValue) === JSON.stringify(value);
          if (valuesMatch) {
            // Both exist with same value, remove display name to avoid duplicate (keep internal name)
            displayNamesToRemove.add(displayName);
          }
          // If display name exists but values differ, keep both (don't overwrite)
        }
        // If display name doesn't exist, don't create it - only use what SharePoint returns
      }
    });
    
    // Remove duplicate display names when internal name exists with same value
    displayNamesToRemove.forEach((displayName) => {
      delete normalized[displayName];
    });
    
    return normalized;
  }

  /**
   * Extract employee name from employee item fields
   */
  private extractEmployeeName(employeeItem: any): string | null {
    const fields = employeeItem.fields || {};
    const empIdValue = fields.EmpID || fields.EmpId || fields.EmpID0 || fields.field_1;
    
    // Skip common non-name values
    const skipValues = ['Assigned', 'Available', 'Item', 'ContentType', 'Edit', 'Attachments'];
    const skipFieldNames = ['cardstatus', 'contenttype', 'accesscardno', 'assets', 'empid', 
                           'employeeid', 'lookupid', 'id', 'odata', 'author', 'editor'];
    
    // Try Title field first
    if (fields.Title) {
      const titleValue = String(fields.Title);
      if (!titleValue.match(/^HPH\s?\d+/) && titleValue !== empIdValue && 
          !skipValues.includes(titleValue) && titleValue.trim().length > 0) {
        return titleValue;
      }
    }
    
    // Try Employee field (Person or Group)
    if (fields.Employee) {
      if (typeof fields.Employee === 'object' && fields.Employee.displayName) {
        return fields.Employee.displayName;
      } else if (typeof fields.Employee === 'string' && 
                 !fields.Employee.match(/^HPH\s?\d+/) && 
                 fields.Employee !== empIdValue) {
        return fields.Employee;
      }
    }
    
    // Search all string fields for name-like values
    for (const [fieldName, fieldValue] of Object.entries(fields)) {
      if (typeof fieldValue !== 'string' || !fieldValue.trim()) continue;
      
      const value = String(fieldValue);
      const fieldNameLower = fieldName.toLowerCase();
      
      if (value.match(/^HPH\s?\d+/) || value === empIdValue) continue;
      if (fieldNameLower.includes('empid') || fieldNameLower.includes('emp_id')) continue;
      if (skipValues.includes(value)) continue;
      if (skipFieldNames.some(skip => fieldNameLower.includes(skip))) continue;
      if (value.match(/^\d+$/) || value.length < 3) continue;
      if (value.match(/^\d{4}-\d{2}-\d{2}/)) continue;
      
      // Check if it looks like a person's name
      if ((value.match(/^[A-Za-z\s]+$/) && value.length > 5) || 
          (value.includes(' ') && value.length > 8 && value.match(/^[A-Za-z\s.]+$/))) {
        return value;
      }
    }
    
    return null;
  }

  /**
   * Insert a record into a SharePoint list
   */
  async insertRecord(listName: string, fields: ListItem): Promise<any> {
    const listId = await this.getListId(listName);
    const siteId = await this.getSiteId();
    
    const itemData = {
      fields: fields,
    };

    const response = await this.makeGraphRequest(
      `sites/${siteId}/lists/${listId}/items`,
      {
        method: 'POST',
        body: JSON.stringify(itemData),
      }
    );

    return {
      d: {
        ...response,
        fields: response.fields || {},
      },
    };
  }

  /**
   * Get User Information List ID (cached)
   */
  private async getUserInformationListId(): Promise<string | null> {
    if (this.userInformationListId) {
      return this.userInformationListId;
    }
    
    try {
      const siteId = await this.getSiteId();
      const listsResponse = await this.makeGraphRequest(`sites/${siteId}/lists?$filter=displayName eq 'User Information List'`);
      const lists = listsResponse.value || [];
      
      if (lists.length > 0 && lists[0].id) {
        this.userInformationListId = lists[0].id;
        return this.userInformationListId;
      }
    } catch (error) {
      // Silently fail - will try other approaches
    }
    
    return null;
  }

  /**
   * Resolve Graph API user ID or email to SharePoint User Information List ID
   */
  async resolveSharePointUserId(graphUserIdOrEmail: string): Promise<number | null> {
    if (!graphUserIdOrEmail || !this.accessToken) {
      return null;
    }
    
    try {
      const userInfoListId = await this.getUserInformationListId();
      if (!userInfoListId) {
        return null;
      }
      
      const siteId = await this.getSiteId();
      
      // Try to find user by email or principal name
      // First, get user info from Graph API if it's a Graph user ID
      let userEmail: string | null = null;
      let userName: string | null = null;
      try {
        const graphUser = await this.makeGraphRequest(`users/${graphUserIdOrEmail}`);
        userEmail = graphUser.mail || graphUser.userPrincipalName;
        userName = graphUser.displayName;
      } catch (error) {
        // If it's already an email, use it directly
        if (graphUserIdOrEmail.includes('@')) {
          userEmail = graphUserIdOrEmail;
        }
      }
      
      if (!userEmail) {
        return null;
      }
      
      // Get all users from User Information List and search manually
      // This is more reliable than filter queries which may not work with all SharePoint setups
      let allUsers: any[] = [];
      let nextLink: string | null = null;
      let pageCount = 0;
      const maxPages = 10; // Limit pages to prevent infinite loops
      
      do {
        const endpoint = nextLink 
          ? nextLink.replace('https://graph.microsoft.com/v1.0/', '')
          : `sites/${siteId}/lists/${userInfoListId}/items?$expand=fields&$top=100`;
        
        const response = await this.makeGraphRequest(endpoint);
        const items = response.value || [];
        allUsers.push(...items);
        
        nextLink = response['@odata.nextLink'] || null;
        pageCount++;
        
        if (pageCount >= maxPages) {
          break;
        }
      } while (nextLink);
      
      // Search for matching user by email or name
      for (const userItem of allUsers) {
        const fields = userItem.fields || {};
        const itemEmail = fields.Email || fields.EMail || fields.Mail || '';
        const itemName = fields.Name || fields.Title || '';
        const itemPrincipalName = fields.UserName || fields.PrincipalName || '';
        
        // Match by email (case-insensitive)
        if (itemEmail && itemEmail.toLowerCase() === userEmail.toLowerCase()) {
          return parseInt(userItem.id, 10);
        }
        
        // Match by principal name (case-insensitive)
        if (itemPrincipalName && itemPrincipalName.toLowerCase() === userEmail.toLowerCase()) {
          return parseInt(userItem.id, 10);
        }
        
        // Match by name if available
        if (userName && itemName && itemName.toLowerCase() === userName.toLowerCase()) {
          return parseInt(userItem.id, 10);
        }
      }
      
      return null;
    } catch (error: any) {
      console.error('Error resolving SharePoint user ID:', error);
      return null;
    }
  }

  /**
   * Resolve SharePoint user ID to display name using User Information List
   */
  private async resolveUserDisplayName(userId: number | string): Promise<string | null> {
    if (!userId || !this.accessToken) {
      return null;
    }
    
    // Check cache first
    if (this.userDisplayNameCache.has(userId)) {
      return this.userDisplayNameCache.get(userId) || null;
    }
    
    const userIdStr = String(userId);
    let displayName: string | null = null;
    
    // Use User Information List via Graph API (direct item lookup)
    try {
      const userInfoListId = await this.getUserInformationListId();
      if (userInfoListId) {
        const siteId = await this.getSiteId();
        const userItemResponse = await this.makeGraphRequest(
          `sites/${siteId}/lists/${userInfoListId}/items/${userIdStr}?$expand=fields`
        );
        const fields = userItemResponse.fields || {};
        if (fields.Title) {
          displayName = fields.Title;
        } else if (fields.Name) {
          displayName = fields.Name;
        }
      }
    } catch (error: any) {
      // Return null if lookup fails
    }
    
    // Cache the result if found
    if (displayName) {
      this.userDisplayNameCache.set(userId, displayName);
    }
    
    return displayName;
  }

  /**
   * Get all items from a list
   */
  async getRecords(listName: string, cachedEmployees?: any[]): Promise<any[]> {
    const listId = await this.getListId(listName);
    const siteId = await this.getSiteId();
    
    const expandQuery = '$expand=fields';
    const url = `sites/${siteId}/lists/${listId}/items?${expandQuery}`;
    
    const response = await this.makeGraphRequest(url);
    
    const items = response.value || [];
    
    // Transform to flat structure - fields are already expanded
    const records = items.map((item: any) => {
      // Fields might be in item.fields or directly in item
      const fields = item.fields || {};
      const record: any = {
        Id: item.id,
        ...fields,
      };
      
      return record;
    });
    
    // Resolve user display names for LookupId fields
    try {
      const recordsWithNames = await Promise.all(
        records.map(async (record: any) => {
          // Find all fields ending with LookupId
          const lookupIdFields = Object.keys(record).filter(key => 
            key.endsWith('LookupId') && 
            (typeof record[key] === 'number' || typeof record[key] === 'string') &&
            record[key] !== null &&
            record[key] !== undefined &&
            record[key] !== ''
          );
          
          // Resolve each LookupId to display name
          for (const lookupIdField of lookupIdFields) {
            const userId = record[lookupIdField];
            if (userId) {
              try {
                const displayName = await this.resolveUserDisplayName(userId);
                if (displayName) {
                  record[`${lookupIdField}_displayName`] = displayName;
                }
              } catch (error: any) {
                // Silently continue if resolution fails for this field
              }
            }
          }
          
          return record;
        })
      );
      
      return recordsWithNames;
    } catch (error: any) {
      console.error('Error resolving user display names:', error);
      // Return records even if resolution fails
      return records;
    }
  }

  /**
   * Update a record in a SharePoint list
   */
  async updateRecord(
    listName: string,
    itemId: number | string,
    fields: ListItem
  ): Promise<any> {
    const listId = await this.getListId(listName);
    const siteId = await this.getSiteId();

    return await this.makeGraphRequest(
      `sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
      {
        method: 'PATCH',
        body: JSON.stringify(fields),
      }
    );
  }

  /**
   * Delete a record from a SharePoint list
   */
  async deleteRecord(listName: string, itemId: number | string): Promise<void> {
    const listId = await this.getListId(listName);
    const siteId = await this.getSiteId();

    await this.makeGraphRequest(
      `sites/${siteId}/lists/${listId}/items/${itemId}`,
      {
        method: 'DELETE',
      }
    );
    // DELETE requests return 204 No Content, which is handled in makeGraphRequest
  }

  /**
   * Get current user information
   */
  async getCurrentUser(): Promise<{
    id: string;
    displayName: string;
    userPrincipalName: string;
    mail: string;
    jobTitle?: string;
    officeLocation?: string;
  }> {
    const response = await this.makeGraphRequest('me');
    return {
      id: response.id,
      displayName: response.displayName,
      userPrincipalName: response.userPrincipalName,
      mail: response.mail || response.userPrincipalName,
      jobTitle: response.jobTitle,
      officeLocation: response.officeLocation,
    };
  }

  /**
   * Check if current user is an admin
   */
  async isCurrentUserAdmin(): Promise<{
    isAdmin: boolean;
    roles: string[];
  }> {
    try {
      const response = await this.makeGraphRequest('me/memberOf');
      const groups = response.value || [];
      
      const adminRoleTemplateIds = [
        '62e90394-69f5-4237-9190-012177145e10', // Global Administrator
        'f28a1f50-f6e7-4571-818b-6a12f2af6b6c', // SharePoint Administrator
        'b0f54661-2d74-4c50-afa3-1ec803f12efe', // Exchange Administrator
        '29232cdf-9323-42fd-ade2-1d097af3e4de', // User Administrator
      ];

      const adminRoles: string[] = [];
      let isAdmin = false;

      for (const group of groups) {
        if (group['@odata.type'] === '#microsoft.graph.directoryRole') {
          const roleTemplateId = group.roleTemplateId;
          if (adminRoleTemplateIds.includes(roleTemplateId)) {
            isAdmin = true;
            adminRoles.push(group.displayName || 'Unknown Role');
          }
        }
      }

      return {
        isAdmin,
        roles: adminRoles,
      };
    } catch (error: any) {
      return {
        isAdmin: false,
        roles: [],
      };
    }
  }

  /**
   * Get all users in the organization
   */
  async getAllUsers(): Promise<Array<{
    id: string;
    displayName: string;
    mail: string;
    jobTitle?: string;
  }>> {
    try {
      const allUsers: any[] = [];
      let nextLink: string | null = null;
      let pageCount = 0;
      const maxPages = 50; // Limit to prevent infinite loops
      
      do {
        const endpoint = nextLink 
          ? nextLink.replace('https://graph.microsoft.com/v1.0/', '') // Remove base URL if present
          : `users?$select=id,displayName,mail,jobTitle&$top=999`;
        
        const response = await this.makeGraphRequest(endpoint);
        
        if (response.value && response.value.length > 0) {
          const users = response.value.map((user: any) => ({
            id: user.id,
            displayName: user.displayName || '',
            mail: user.mail || '',
            jobTitle: user.jobTitle,
          }));
          
          allUsers.push(...users);
        }
        
        // Check for next page
        nextLink = response['@odata.nextLink'] || null;
        pageCount++;
        
        // Safety check to prevent infinite loops
        if (pageCount >= maxPages) {
          break;
        }
      } while (nextLink);
      
      return allUsers;
    } catch (error: any) {
      const errorMessage = error.message || '';
      if (
        errorMessage.includes('403') || 
        errorMessage.includes('Forbidden') ||
        errorMessage.includes('Insufficient') ||
        errorMessage.includes('Directory.Read') ||
        errorMessage.includes('User.Read.All')
      ) {
        return [];
      }
      return [];
    }
  }

  /**
   * Get current user info with admin status
   */
  async getCurrentUserWithAdminStatus(): Promise<{
    user: {
      id: string;
      displayName: string;
      mail: string;
      jobTitle?: string;
    };
    isAdmin: boolean;
    roles: string[];
  }> {
    const [user, adminStatus] = await Promise.all([
      this.getCurrentUser(),
      this.isCurrentUserAdmin(),
    ]);

    return {
      user,
      ...adminStatus,
    };
  }
}

export default SharePointService;
