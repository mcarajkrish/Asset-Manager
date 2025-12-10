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

    if (!response.ok) {
      const errorText = await response.text();
      
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
        } catch (parseError) {
          // If error parsing fails, continue with generic error
        }
      }
      
      throw new Error(`Microsoft Graph API error: ${response.status} - ${errorText}`);
    }

    return response.json();
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
   * Expand lookup fields by fetching full details for fields that only have LookupId
   */
  private async expandLookupFields(
    records: any[],
    listName: string,
    siteId: string,
    listId: string,
    lookupFields: string[],
    cachedEmployees?: any[]
  ): Promise<void> {
    // For each lookup field, check if we need to expand it
    for (const lookupField of lookupFields) {
      const lookupIdField = `${lookupField}LookupId`;
      
      // Collect unique lookup IDs that need expansion
      const lookupIdsToExpand = new Set<string | number>();
      const recordsNeedingExpansion: any[] = [];
      
      records.forEach((record) => {
        const lookupId = record[lookupIdField];
        const lookupValue = record[lookupField];
        
        // If we have LookupId but not an expanded object, we need to fetch it
        if (lookupId != null && lookupId !== '' && 
            (lookupValue == null || typeof lookupValue !== 'object' || Object.keys(lookupValue).length === 0)) {
          lookupIdsToExpand.add(lookupId);
          recordsNeedingExpansion.push(record);
        }
      });
      
      if (lookupIdsToExpand.size === 0) continue;
      
      // Determine if this is a Person field or a Lookup field
      // For Person fields, we'll try to fetch from Microsoft Graph users
      // For Lookup fields, we need to determine the target list
      
      // Try to fetch lookup details
      // First, check if it's a Person field by trying to fetch from Graph users
      const lookupIdsArray = Array.from(lookupIdsToExpand);
      
      // Check if IDs are GUIDs (Microsoft Graph user IDs)
      const guidIds = lookupIdsArray.filter(id => {
        const idStr = String(id);
        return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(idStr);
      });
      
      // Fetch Person field details from Microsoft Graph
      if (guidIds.length > 0) {
        const personDetailsMap = new Map<string, any>();
        
        // Fetch user details in parallel
        const fetchPromises = guidIds.map(async (guidId) => {
          try {
            const userResponse = await this.makeGraphRequest(`users/${guidId}?$select=id,displayName,userPrincipalName,mail,jobTitle`);
            return { id: guidId, details: userResponse };
          } catch (error) {
            // User not found or access denied
            return null;
          }
        });
        
        const results = await Promise.all(fetchPromises);
        results.forEach((result) => {
          if (result && result.details) {
            personDetailsMap.set(String(result.id), {
              id: result.details.id,
              displayName: result.details.displayName,
              userPrincipalName: result.details.userPrincipalName,
              email: result.details.mail || result.details.userPrincipalName,
              mail: result.details.mail || result.details.userPrincipalName,
              jobTitle: result.details.jobTitle,
            });
          }
        });
        
        // Populate records with Person field details
        recordsNeedingExpansion.forEach((record) => {
          const lookupId = record[lookupIdField];
          if (lookupId != null) {
            const personDetails = personDetailsMap.get(String(lookupId));
            if (personDetails) {
              record[lookupField] = personDetails;
            }
          }
        });
      }
      
      // For non-GUID IDs (SharePoint list item IDs), try to resolve from cached employees
      // This is useful for Assets list where assignee might point to Employees list
      const integerIds = lookupIdsArray.filter(id => {
        const idStr = String(id);
        const isGuid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(idStr);
        if (isGuid) return false;
        const intId = parseInt(idStr, 10);
        return !isNaN(intId) && intId > 0;
      });
      
      if (integerIds.length > 0 && cachedEmployees && cachedEmployees.length > 0) {
        const integerDetailsMap = new Map<string, any>();
        
        integerIds.forEach((intId) => {
          const intIdStr = String(intId);
          const employee = this.findEmployeeById(cachedEmployees, intId);
          
          if (employee) {
            // Create a person-like object from employee record
            integerDetailsMap.set(intIdStr, {
              id: employee.Id,
              displayName: employee.Employee || employee.EmployeeName || employee.Title || employee.displayName,
              userPrincipalName: employee.Email || employee.Mail || employee.userPrincipalName,
              email: employee.Email || employee.Mail || employee.email || employee.mail,
              mail: employee.Email || employee.Mail || employee.email || employee.mail,
              jobTitle: employee.Designation || employee.JobTitle,
            });
          }
        });
        
        // Populate records with integer lookup details
        recordsNeedingExpansion.forEach((record) => {
          const lookupId = record[lookupIdField];
          if (lookupId != null && !record[lookupField]) {
            const lookupDetails = integerDetailsMap.get(String(lookupId));
            if (lookupDetails) {
              record[lookupField] = lookupDetails;
            }
          }
        });
      }
    }
  }

  /**
   * Resolve Employee names for Access Cards - fetches from SharePoint only, no mapping to Employees records
   */
  private async resolveEmployeeNames(records: any[], siteId: string, cachedEmployees?: any[]): Promise<void> {
    // Check if Employee field is already expanded
    const firstRecord = records[0];
    if (firstRecord?.Employee && typeof firstRecord.Employee === 'object') {
      records.forEach((record: any) => {
        if (record.Employee && typeof record.Employee === 'object') {
          const employeeObj = record.Employee;
          const employeeName = employeeObj.displayName || 
                             employeeObj.LookupValue || 
                             employeeObj.Title ||
                             employeeObj.email ||
                             employeeObj.mail ||
                             employeeObj.userPrincipalName;
          
          if (employeeName && !employeeName.match(/^HPH\s?\d+/)) {
            record['Employee'] = employeeName;
            record['EmployeeName'] = employeeName;
          }
        }
      });
      // Continue to check EmployeeLookupId even if Employee field is expanded
    }
    
    // Collect unique EmployeeLookupId values
    const employeeIdsToFetch = new Set<string>();
    records.forEach((record: any) => {
      if (record.EmployeeLookupId != null) {
        employeeIdsToFetch.add(String(record.EmployeeLookupId));
      }
    });
    
    if (employeeIdsToFetch.size === 0) return;
    
    const employeeNameCache = new Map<string, string>();
    
    // Get Employees list ID for integer IDs (SharePoint list item IDs)
    let employeesListId: string | null = null;
    const integerIds = Array.from(employeeIdsToFetch).filter(id => {
      const intId = parseInt(id, 10);
      return !isNaN(intId) && intId > 0;
    });
    
    if (integerIds.length > 0) {
      try {
        employeesListId = await this.getListId('Employees');
      } catch (error) {
        // Employees list not found
      }
      
      // Get Access Cards list ID once (needed as fallback even if Employees list exists)
      let accessCardsListId: string | null = null;
      try {
        accessCardsListId = await this.getListId('Access Cards');
      } catch (error) {
        // Access Cards list not found
      }
      
      // Fetch employee names from SharePoint API in parallel
      const fetchPromises = integerIds.map(async (intIdStr) => {
        const intId = parseInt(intIdStr, 10);
        try {
          let employeeItem: any = null;
          
          // Try Employees list first
          if (employeesListId) {
            try {
              employeeItem = await this.makeGraphRequest(
                `sites/${siteId}/lists/${employeesListId}/items/${intId}?$expand=fields`
              );
            } catch (error) {
              // Try Access Cards list if Employees list fails
            }
          }
          
          // Try Access Cards list if Employees list failed or doesn't exist
          if (!employeeItem && accessCardsListId) {
            try {
              employeeItem = await this.makeGraphRequest(
                `sites/${siteId}/lists/${accessCardsListId}/items/${intId}?$expand=fields`
              );
              
              // If it's an Access Card, extract Employee field from it
              if (employeeItem?.fields?.AccessCardNo && employeeItem.fields?.Employee) {
                const empField = employeeItem.fields.Employee;
                if (typeof empField === 'object' && empField.displayName) {
                  return { intIdStr, employeeName: empField.displayName };
                } else if (typeof empField === 'string' && !empField.match(/^HPH\s?\d+/)) {
                  return { intIdStr, employeeName: empField };
                }
              }
            } catch (error) {
              // Continue to try extracting from employeeItem if available
            }
          }
          
          // Extract employee name from the item
          if (employeeItem) {
            const employeeName = this.extractEmployeeName(employeeItem);
            if (employeeName) {
              return { intIdStr, employeeName };
            }
          }
        } catch (error: any) {
          // Failed to fetch employee name
        }
        return null;
      });
      
      // Wait for all requests to complete in parallel
      const results = await Promise.all(fetchPromises);
      
      // Populate cache from results
      results.forEach((result) => {
        if (result && result.employeeName) {
          employeeNameCache.set(result.intIdStr, result.employeeName);
        }
      });
    }
    
    // Populate Employee field in all records
    records.forEach((record: any) => {
      if (record.EmployeeLookupId != null) {
        const lookupId = String(record.EmployeeLookupId);
        const employeeName = employeeNameCache.get(lookupId);
        if (employeeName) {
          record['Employee'] = employeeName;
          record['EmployeeName'] = employeeName;
        } else {
          record['Employee'] = `[ID: ${lookupId}]`;
        }
      }
    });
  }

  /**
   * Find employee in cached employees by multiple criteria
   */
  private findEmployeeById(cachedEmployees: any[], lookupId: string | number): any | null {
    if (!cachedEmployees || cachedEmployees.length === 0) return null;
    
    const lookupIdStr = String(lookupId);
    
    // Check if it's a GUID (Microsoft Graph user ID format)
    const isGuidFormat = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(lookupIdStr);
    
    if (isGuidFormat) {
      // Match by Graph user ID (GUID)
      const byGuid = cachedEmployees.find((emp: any) => {
        const empId = String(emp.Id || '');
        return empId.toLowerCase() === lookupIdStr.toLowerCase();
      });
      if (byGuid) return byGuid;
    }
    
    // Try to parse as integer (SharePoint list item ID)
    const lookupIdInt = typeof lookupId === 'string' ? parseInt(lookupId, 10) : lookupId;
    if (!isNaN(lookupIdInt) && lookupIdInt > 0) {
      // First try matching by integer ID (in case it's a SharePoint list item ID)
      const byIntId = cachedEmployees.find((emp: any) => {
        const empId = typeof emp.Id === 'string' ? parseInt(emp.Id, 10) : emp.Id;
        return !isNaN(empId) && empId === lookupIdInt;
      });
      if (byIntId) return byIntId;
      
      // Also try matching by EmpID field (in case lookupId is an EmpID)
      const byEmpId = cachedEmployees.find((emp: any) => {
        const empId = String(emp.EmpID || emp.EmpId || '');
        return empId === lookupIdStr || empId === String(lookupIdInt);
      });
      if (byEmpId) return byEmpId;
    }
    
    // Try matching by email or userPrincipalName (in case lookupId is an email)
    if (lookupIdStr.includes('@')) {
      const byEmail = cachedEmployees.find((emp: any) => {
        const empEmail = String(emp.Email || emp.Mail || emp.email || emp.mail || emp.userPrincipalName || emp.UserPrincipalName || '');
        return empEmail.toLowerCase() === lookupIdStr.toLowerCase();
      });
      if (byEmail) return byEmail;
    }
    
    // Try matching by userPrincipalName prefix (before @)
    if (lookupIdStr.includes('@')) {
      const emailPrefix = lookupIdStr.split('@')[0];
      const byUpnPrefix = cachedEmployees.find((emp: any) => {
        const empUpn = String(emp.userPrincipalName || emp.UserPrincipalName || '');
        return empUpn.split('@')[0] === emailPrefix;
      });
      if (byUpnPrefix) return byUpnPrefix;
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
   * Get all items from a list
   */
  async getRecords(listName: string, cachedEmployees?: any[]): Promise<any[]> {
    const listId = await this.getListId(listName);
    const siteId = await this.getSiteId();
    
    // Get field mapping and lookup fields to convert internal names to display names
    const { mapping: fieldMapping, lookupFields } = await this.getFieldMapping(listName);
    
    // Build expand query to include lookup fields
    // Microsoft Graph API supports expanding lookup fields within fields
    let expandQuery = '$expand=fields';
    if (lookupFields.length > 0) {
      // For each lookup field, we need to expand it
      // Microsoft Graph API syntax: $expand=fields(field_2)
      // However, we can only expand one field at a time in nested expands
      // So we'll expand the first lookup field, or use a different approach
      // Actually, we can use: $expand=fields($expand=field_2) but this might not work
      // Let's try expanding lookup fields by fetching them separately if needed
      // For now, just expand fields and we'll handle lookup expansion below
    }
    
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
    
    // Expand lookup fields that only have LookupId but not the full object
    // Do this BEFORE normalizing field names so expanded objects are in original field names
    if (lookupFields.length > 0 && records.length > 0) {
      await this.expandLookupFields(records, listName, siteId, listId, lookupFields, cachedEmployees);
    }
    
    // Normalize field names: convert internal names (field_0, field_2, etc.) to display names
    // This will copy expanded lookup objects to display name fields too
    const normalizedRecords = records.map((record: any) => {
      return this.normalizeFieldNames(record, fieldMapping);
    });
    
    // Resolve lookup fields (still use internal names for resolution logic)
    if (listName === 'Access Cards' && normalizedRecords.length > 0) {
      await this.resolveEmployeeNames(normalizedRecords, siteId, cachedEmployees);
    }
    
    return normalizedRecords;
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
    userPrincipalName: string;
    mail: string;
    jobTitle?: string;
    officeLocation?: string;
    department?: string;
  }>> {
    try {
      // Fetch all users from Microsoft Graph API with pagination
      // Using $select to get only needed fields for better performance
      const allUsers: any[] = [];
      let nextLink: string | null = null;
      let pageCount = 0;
      const maxPages = 50; // Limit to prevent infinite loops
      
      do {
        const endpoint = nextLink 
          ? nextLink.replace('https://graph.microsoft.com/v1.0/', '') // Remove base URL if present
          : `users?$select=id,displayName,userPrincipalName,mail,jobTitle,officeLocation,department&$top=999`;
        
        const response = await this.makeGraphRequest(endpoint);
        
        if (response.value && response.value.length > 0) {
          const users = response.value.map((user: any) => ({
            id: user.id,
            displayName: user.displayName || '',
            userPrincipalName: user.userPrincipalName || '',
            mail: user.mail || user.userPrincipalName || '',
            jobTitle: user.jobTitle,
            officeLocation: user.officeLocation,
            department: user.department,
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
   * Get list of admin users
   */
  async getAdminUsers(): Promise<Array<{
    id: string;
    displayName: string;
    userPrincipalName: string;
    mail: string;
    roles: string[];
  }>> {
    try {
      const rolesResponse = await this.makeGraphRequest('directoryRoles');
      const globalAdminRole = rolesResponse.value?.find(
        (role: any) => role.roleTemplateId === '62e90394-69f5-4237-9190-012177145e10'
      );

      if (!globalAdminRole) {
        return [];
      }

      const membersResponse = await this.makeGraphRequest(
        `directoryRoles/${globalAdminRole.id}/members`
      );

      const admins = [];
      for (const member of membersResponse.value || []) {
        if (member['@odata.type'] === '#microsoft.graph.user') {
          admins.push({
            id: member.id,
            displayName: member.displayName,
            userPrincipalName: member.userPrincipalName,
            mail: member.mail || member.userPrincipalName,
            roles: ['Global Administrator'],
          });
        }
      }

      return admins;
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
      userPrincipalName: string;
      mail: string;
      jobTitle?: string;
      officeLocation?: string;
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
