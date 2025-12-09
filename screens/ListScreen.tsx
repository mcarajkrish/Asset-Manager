import React, { useState, useEffect, useCallback } from 'react';
import { useFocusEffect } from '@react-navigation/native';
import {
  View,
  Text,
  TouchableOpacity,
  StyleSheet,
  ScrollView,
  ActivityIndicator,
  RefreshControl,
  Alert,
  TextInput,
} from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import SharePointService, { SessionTimeoutError } from '../services/sharepointService';

interface ListScreenProps {
  sharePointService: SharePointService;
  listName: string;
  employees?: any[];
  onRefreshEmployees?: () => Promise<void>;
  onRecordPress: (record: any) => void;
  onCreatePress?: () => void;
  onBack: () => void;
}

interface Record {
  Id: number | string;
  Title?: string;
  [key: string]: any;
}

const ListScreen: React.FC<ListScreenProps> = ({
  sharePointService,
  listName,
  employees = [],
  onRefreshEmployees,
  onRecordPress,
  onCreatePress,
  onBack,
}) => {
  const [records, setRecords] = useState<Record[]>([]);
  const [loading, setLoading] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [searchQuery, setSearchQuery] = useState<string>('');

  const loadRecords = async () => {
    try {
      setLoading(true);
      setError(null);
      
      // For Employees list, use cached employees from organization instead of fetching from SharePoint
      if (listName === 'Employees' && employees && employees.length > 0) {
        setRecords(employees);
      } else {
        const items = await sharePointService.getRecords(listName, employees);
        setRecords(items);
      }
    } catch (error: any) {
      const errorMessage = error.message || 'Failed to load records';
      setError(errorMessage);
      console.error('Error loading records:', error);
      
      // Handle session timeout - don't show alert as App.tsx will handle it
      if (error instanceof SessionTimeoutError) {
        // Session timeout is handled by App.tsx callback, just set error state
        setError('Session expired. Please log in again.');
        return;
      }
      
      Alert.alert('Error', errorMessage);
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  };

  // Load records when screen comes into focus (e.g., when returning from detail screen)
  useFocusEffect(
    useCallback(() => {
      loadRecords();
    }, [listName, employees])
  );

  const handleRefresh = async () => {
    setRefreshing(true);
    // If this is Employees list and we have refresh callback, refresh employees cache first
    if (listName === 'Employees' && onRefreshEmployees) {
      await onRefreshEmployees();
    }
    loadRecords();
  };

  const handleDelete = async (recordId: number | string) => {
    Alert.alert(
      'Delete Record',
      'Are you sure you want to delete this record? This action cannot be undone.',
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Delete',
          style: 'destructive',
          onPress: async () => {
            try {
              await sharePointService.deleteRecord(listName, recordId);
              Alert.alert('Success', 'Record deleted successfully!');
              loadRecords();
            } catch (error: any) {
              // Handle session timeout - don't show alert as App.tsx will handle it
              if (error instanceof SessionTimeoutError) {
                return;
              }
              Alert.alert('Error', error.message || 'Failed to delete record');
            }
          },
        },
      ]
    );
  };

  const getDisplayValue = (record: Record): string => {
    if (record.Title) return record.Title;
    
    const displayFields = ['Name', 'AssetName', 'EmployeeName', 'CardNumber', 'AccessCardNumber'];
    for (const field of displayFields) {
      if (record[field]) {
        return String(record[field]);
      }
    }
    
    const keys = Object.keys(record).filter(
      (key) => key !== 'Id' && key !== '__metadata' && !key.startsWith('_')
    );
    if (keys.length > 0) {
      const firstValue = record[keys[0]];
      if (firstValue !== null && firstValue !== undefined) {
        return String(firstValue);
      }
    }
    
    return `Record #${record.Id}`;
  };

  // Search function to check if record matches search query
  const matchesSearch = (record: Record, query: string): boolean => {
    if (!query.trim()) return true;
    
    const searchTerm = query.toLowerCase().trim();
    
    // Search through all record fields
    for (const [key, value] of Object.entries(record)) {
      // Skip metadata fields
      if (key === 'Id' || key.startsWith('_') || key === '__metadata') continue;
      
      // Convert value to string for searching
      let searchableValue = '';
      if (value === null || value === undefined) continue;
      
      if (typeof value === 'object' && !Array.isArray(value)) {
        // Handle object values (lookup fields)
        searchableValue = value.Title || value.displayName || value.name || value.LookupValue || JSON.stringify(value);
      } else if (Array.isArray(value)) {
        searchableValue = value.join(' ');
      } else {
        searchableValue = String(value);
      }
      
      // Check if search term matches
      if (searchableValue.toLowerCase().includes(searchTerm)) {
        return true;
      }
    }
    
    return false;
  };

  // Filter records based on search query
  const filteredRecords = records.filter(record => matchesSearch(record, searchQuery));

  const getFieldValue = (record: Record, fieldNames: string[]): string => {
    // First, try exact field names (case-insensitive)
    for (const fieldName of fieldNames) {
      // Try exact match first
      let value = record[fieldName];
      
      // Try case-insensitive match
      if (value === null || value === undefined || value === '') {
        const recordKeys = Object.keys(record);
        const exactMatch = recordKeys.find(key => key.toLowerCase() === fieldName.toLowerCase());
        if (exactMatch) {
          value = record[exactMatch];
        }
      }
      
      // Handle empty strings as not found
      if (value === null || value === undefined || value === '') {
        continue;
      }
      
      // Handle lookup fields (objects with Title property)
      if (typeof value === 'object' && !Array.isArray(value)) {
        // Check for Title property (common in lookup fields)
        if (value.Title) {
          return String(value.Title);
        }
        // Check for displayName or name
        if (value.displayName) {
          return String(value.displayName);
        }
        if (value.name) {
          return String(value.name);
        }
        // Check for LookupValue (common in SharePoint lookup fields)
        if (value.LookupValue) {
          return String(value.LookupValue);
        }
        // Check for email property
        if (value.email) {
          return String(value.email);
        }
        // If it's an object but no readable property, try to stringify
        const stringified = JSON.stringify(value);
        if (stringified !== '{}' && stringified !== 'null') {
          return stringified;
        }
        continue;
      }
      
      // Return string value if not empty
      const stringValue = String(value).trim();
      if (stringValue !== '' && stringValue !== 'null' && stringValue !== 'undefined') {
        return stringValue;
      }
    }
    
    // If not found, search through all record keys for partial matches
    const recordKeys = Object.keys(record);
    for (const fieldName of fieldNames) {
      const searchTerm = fieldName.toLowerCase();
      // Find keys that contain the search term, but exclude ID fields
      const matchingKey = recordKeys.find(key => {
        const keyLower = key.toLowerCase();
        // Exclude fields that end with 'id' or 'lookupid' when searching for display values
        const isIdField = keyLower.endsWith('id') || keyLower.endsWith('lookupid');
        const matches = (keyLower.includes(searchTerm) || searchTerm.includes(keyLower));
        return matches && !isIdField;
      });
      
      if (matchingKey) {
        const value = record[matchingKey];
        if (value !== null && value !== undefined && value !== '') {
          // Handle lookup fields (objects with Title property)
          if (typeof value === 'object' && !Array.isArray(value)) {
            if (value.Title) {
              return String(value.Title);
            }
            if (value.displayName) {
              return String(value.displayName);
            }
            if (value.name) {
              return String(value.name);
            }
            if (value.LookupValue) {
              return String(value.LookupValue);
            }
            const stringified = JSON.stringify(value);
            if (stringified !== '{}' && stringified !== 'null') {
              return stringified;
            }
            continue;
          }
          const stringValue = String(value).trim();
          if (stringValue !== '' && stringValue !== 'null' && stringValue !== 'undefined') {
            return stringValue;
          }
        }
      }
    }
    
    return '-';
  };

  const renderAccessCardRecord = (record: Record) => {
    const accessCardNo = getFieldValue(record, ['AccessCardNo', 'AccessCardNumber', 'CardNumber', 'CardNo', 'field_1', 'field1', 'Title']);
    const cardStatus = getFieldValue(record, ['CardStatus', 'Status', 'field_2', 'field2']);
    
    // For Employee field, explicitly exclude ID fields and look for display values
    // First, find all non-ID employee fields
    const allKeys = Object.keys(record);
    const employeeFieldKeys = allKeys.filter(key => {
      const keyLower = key.toLowerCase();
      const isIdField = keyLower.endsWith('id') || keyLower.endsWith('lookupid');
      return (keyLower.includes('employee') || keyLower.includes('emp')) && !isIdField;
    });
    
    // Try multiple variations of Employee field names (excluding ID fields)
    const employee = getFieldValue(record, [
      'Employee', 
      'EmployeeName', 
      'EmployeeName_x003a_', 
      'Employee/Title',
      'Employee_x003a_Title',
      'Employee_x003a_Name',
      ...employeeFieldKeys
    ]);
    
    const empId = getFieldValue(record, ['EmpId', 'EmployeeId', 'EmpID', 'EmployeeID', 'EmployeeLookupId']);

    return (
      <View style={styles.accessCardContent}>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Access Card No:</Text>
          <Text style={styles.accessCardValue}>{accessCardNo}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Card Status:</Text>
          <Text style={styles.accessCardValue}>{cardStatus}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Assignee:</Text>
          <Text style={styles.accessCardValue}>{employee}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Emp Id:</Text>
          <Text style={styles.accessCardValue}>{empId}</Text>
        </View>
      </View>
    );
  };

  const renderEmployeeRecord = (record: Record) => {
    // Get Emp ID - try multiple field name variations
    const empId = getFieldValue(record, [
      'EmpID',
      'EmpId', 
      'EmpID0',
      'EmployeeID',
      'EmployeeId',
      'EmployeeID0',
      'ID',
      'Id'
    ]);
    
    // Get Emp Name - try multiple field name variations and search more thoroughly
    // First, try common field names
    let empName = getFieldValue(record, [
      'Employee',
      'EmployeeName',
      'Title',
      'Name',
      'EmployeeName_x003a_',
      'Employee/Title',
      'Employee_x003a_Title',
      'Employee_x003a_Name',
      'FullName',
      'DisplayName'
    ]);
    
    // If not found, search through all fields for name-like values
    if (empName === '-' || !empName || empName.trim() === '') {
      const allKeys = Object.keys(record);
      const empIdValue = empId;
      const skipFieldPatterns = [
        'lookupid', 'lookup', 'odata', 'contenttype', 'modified', 'created',
        'author', 'editor', 'attachments', 'edit', 'folder', 'item', 'compliance',
        'version', 'status', 'cardstatus', 'accesscardno', 'assets'
      ];
      
      // Look for fields that might contain the name
      for (const key of allKeys) {
        const keyLower = key.toLowerCase();
        
        // Skip system/metadata fields
        if (keyLower.startsWith('_') || 
            skipFieldPatterns.some(pattern => keyLower.includes(pattern))) {
          continue;
        }
        
        // Skip if it's clearly an ID field (ends with 'id' or 'lookupid', but allow fields like 'EmployeeId' if they contain 'name' or 'employee')
        if ((keyLower.endsWith('id') || keyLower.endsWith('lookupid')) && 
            !keyLower.includes('name') && !keyLower.includes('employee')) {
          continue;
        }
        
        // Skip email/mail fields (we'll get those separately)
        if (keyLower.includes('email') || keyLower.includes('mail')) {
          continue;
        }
        
        const value = record[key];
        if (value && typeof value === 'string' && value.trim() !== '') {
          const valueStr = String(value).trim();
          // Skip if it's the same as Emp ID
          if (valueStr === empIdValue || valueStr === empId) continue;
          // Skip if it looks like an ID (starts with HPH and numbers)
          if (valueStr.match(/^HPH\s?\d+/)) continue;
          // Skip short values or numbers only
          if (valueStr.length < 3 || valueStr.match(/^\d+$/)) continue;
          // Skip common non-name values
          if (['Assigned', 'Available', 'Item', 'ContentType', 'Edit'].includes(valueStr)) continue;
          // Skip date strings
          if (valueStr.match(/^\d{4}-\d{2}-\d{2}/)) continue;
          
          // If it looks like a name (has letters and spaces, reasonable length)
          if ((valueStr.match(/^[A-Za-z\s.]+$/) && valueStr.length > 5) || 
              (valueStr.includes(' ') && valueStr.length > 8)) {
            empName = valueStr;
            break;
          }
        } else if (value && typeof value === 'object' && !Array.isArray(value)) {
          // Check if it's an object with displayName or Title
          if (value.displayName) {
            empName = String(value.displayName);
            break;
          } else if (value.Title) {
            empName = String(value.Title);
            break;
          } else if (value.name) {
            empName = String(value.name);
            break;
          }
        }
      }
      
    }
    
    // Get Mail/Email - try multiple field name variations
    const mailId = getFieldValue(record, [
      'Email',
      'Mail',
      'email',
      'mail',
      'EmailAddress',
      'EmailAddress0',
      'MailAddress',
      'UserPrincipalName',
      'userPrincipalName',
      'UPN'
    ]);
    
    // Get Designation/JobTitle - try multiple field name variations
    const designation = getFieldValue(record, [
      'Designation',
      'JobTitle',
      'jobTitle',
      'Title',
      'Position',
      'Role',
      'Job Title',
      'Designation0'
    ]);

    return (
      <View style={styles.accessCardContent}>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Emp Name:</Text>
          <Text style={styles.accessCardValue}>{empName || '-'}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Mail Id:</Text>
          <Text style={styles.accessCardValue}>{mailId}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Emp Id:</Text>
          <Text style={styles.accessCardValue}>{empId}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Designation:</Text>
          <Text style={styles.accessCardValue}>{designation || '-'}</Text>
        </View>
      </View>
    );
  };

  const renderAssetRecord = (record: Record) => {
    // Debug: Log all available fields for Assets
    const recordIndex = records.indexOf(record);
    if (recordIndex === 0) {

      // Log all non-empty field values
      const nonEmptyFields = Object.entries(record)
        .filter(([key, value]) => {
          if (key === 'Id' || key.startsWith('_')) return false;
          if (value === null || value === undefined || value === '') return false;
          if (typeof value === 'object' && Object.keys(value).length === 0) return false;
          return true;
        })
        .map(([key, value]) => `${key}: ${JSON.stringify(value)}`);
    }
    
    console.log('[Assets] Available fields:', Object.keys(record));
    console.log('[Assets] Available values:', Object.values(record));

    // Get Asset Id - try multiple field name variations
    const assetId = getFieldValue(record, [
      'AssetId', 
      'AssetID', 
      'Asset_Id', 
      'Asset Id',
      'AssetID0',
      'field_1',
      'field1',
      'ID',
      'Id',
      'Title'
    ]);
    
    // Get Brand and Company separately - try more variations including encoded field names
    // Note: field_2 is the assignee field, not company, so exclude it from company lookup
    const company = getFieldValue(record, ['Company', 'Company0', 'company', 'field2', 'Brand', 'brand']);
    const model = getFieldValue(record, ['Model', 'Model0', 'model', 'field_3', 'field3']);
    
    // Combine Brand + Company for Asset field
    const asset = [company, model].filter(Boolean).join(' ').trim() || '-';
    
    // Get Serial Number - try field_4 and other variations
    const serialNumber = getFieldValue(record, ['field_4', 'field4', 'SerialNumber', 'Serial_Number', 'SerialNo', 'Serial_No', 'Serial Number']);
    
    // For Assignee field, prioritize the resolved Assignee/AssigneeName fields set by resolveAssigneeNames
    // Only use fallbacks if the service hasn't populated these fields
    let assignee = '-';
    
    // First, check if Assignee or AssigneeName fields were populated by the service (highest priority)
    const resolvedAssignee = record['Assignee'] || record['AssigneeName'];
    if (resolvedAssignee && resolvedAssignee !== '-' && !String(resolvedAssignee).startsWith('[ID:')) {
      assignee = String(resolvedAssignee);
    } else {
      // If service didn't resolve it, try other field name variations
      assignee = getFieldValue(record, [
        'Assignee',
        'AssigneeName',
        'Assignee_x003a_Title',
        'Assignee/Title',
        'AssignedTo',
        'AssignedToName',
        'Assigned To',
        'Assign',
      ]);
      
      // If still not found and we have field_2LookupId, try to resolve from cached employees
      if (assignee === '-' && record['field_2LookupId'] && employees && employees.length > 0) {
        const assigneeId = String(record['field_2LookupId']);
        
        // Check if it's a GUID (Microsoft Graph user ID)
        const isGuidFormat = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(assigneeId);
        
        let cachedEmployee: any = null;
        
        if (isGuidFormat) {
          // Match by GUID (Graph user ID)
          cachedEmployee = employees.find((emp: any) => {
            const empId = String(emp.Id || '');
            return empId.toLowerCase() === assigneeId.toLowerCase();
          });
        } else {
          // Try integer matching (SharePoint list item ID)
          const assigneeIdInt = parseInt(assigneeId, 10);
          if (!isNaN(assigneeIdInt) && assigneeIdInt > 0) {
            cachedEmployee = employees.find((emp: any) => {
              const cachedId = typeof emp.Id === 'string' ? parseInt(emp.Id, 10) : emp.Id;
              return cachedId === assigneeIdInt;
            });
            
            // If not found by integer ID, try matching by EmpID field
            if (!cachedEmployee) {
              cachedEmployee = employees.find((emp: any) => {
                const empId = String(emp.EmpID || emp.EmpId || '');
                return empId === assigneeId || empId === String(assigneeIdInt);
              });
            }
          }
        }
        
        if (cachedEmployee) {
          assignee = cachedEmployee.Employee || 
                    cachedEmployee.EmployeeName || 
                    cachedEmployee.Title ||
                    cachedEmployee.displayName ||
                    '-';
        }
      }
      
      // Final fallback: check field_2 object only if we still haven't found anything
      // This should rarely be needed if resolveAssigneeNames is working correctly
      if (assignee === '-' && record['field_2'] && typeof record['field_2'] === 'object' && !Array.isArray(record['field_2'])) {
        const field2Value = record['field_2'].displayName || 
                           record['field_2'].Title || 
                           record['field_2'].userPrincipalName ||
                           record['field_2'].email ||
                           record['field_2'].name;
        if (field2Value && field2Value !== 'null' && field2Value !== 'undefined' && String(field2Value).trim() !== '') {
          assignee = String(field2Value).trim();
        }
      }
      
      // Last resort: show the ID if we have field_2LookupId but couldn't resolve the name
      if (assignee === '-' && record['field_2LookupId']) {
        assignee = `[ID: ${record['field_2LookupId']}]`;
      }
    }
    

    return (
      <View style={styles.accessCardContent}>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Asset Id:</Text>
          <Text style={styles.accessCardValue}>{assetId}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Asset:</Text>
          <Text style={styles.accessCardValue}>{asset}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Serial Number:</Text>
          <Text style={styles.accessCardValue}>{serialNumber}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Assignee:</Text>
          <Text style={styles.accessCardValue}>{assignee}</Text>
        </View>
      </View>
    );
  };

  return (
    <SafeAreaView style={styles.container} edges={['top', 'bottom']}>
      {/* Header */}
      <View style={styles.header}>
        <TouchableOpacity onPress={onBack} style={styles.backButton}>
          <Text style={styles.backButtonText}>←</Text>
        </TouchableOpacity>
        <Text style={styles.headerTitle}>{listName}</Text>
        {/* {onCreatePress && (
          <TouchableOpacity onPress={onCreatePress} style={styles.addButton}>
            <Text style={styles.addButtonText}>+ Add</Text>
          </TouchableOpacity>
        )} */}
        {!onCreatePress && <View style={styles.placeholder} />}
      </View>

      {/* Search Bar */}
      {!loading && records.length > 0 && (
        <View style={styles.searchContainer}>
          <TextInput
            style={styles.searchInput}
            placeholder={`Search ${listName.toLowerCase()}...`}
            placeholderTextColor="#999"
            value={searchQuery}
            onChangeText={setSearchQuery}
            autoCapitalize="none"
            autoCorrect={false}
          />
          {searchQuery.length > 0 && (
            <TouchableOpacity
              style={styles.clearButton}
              onPress={() => setSearchQuery('')}
            >
              <Text style={styles.clearButtonText}>✕</Text>
            </TouchableOpacity>
          )}
        </View>
      )}

      {/* Count */}
      {!loading && (
        <View style={styles.countContainer}>
          <Text style={styles.countText}>
            {searchQuery.trim() 
              ? `${filteredRecords.length} of ${records.length} record(s)`
              : `${records.length} record(s)`
            }
          </Text>
        </View>
      )}

      {/* Content */}
      {loading && records.length === 0 ? (
        <View style={styles.loadingContainer}>
          <ActivityIndicator size="large" color="#0078d4" />
          <Text style={styles.loadingText}>Loading records...</Text>
        </View>
      ) : error ? (
        <View style={styles.errorContainer}>
          <Text style={styles.errorText}>{error}</Text>
          <TouchableOpacity style={styles.retryButton} onPress={loadRecords}>
            <Text style={styles.retryButtonText}>Retry</Text>
          </TouchableOpacity>
        </View>
      ) : records.length === 0 ? (
        <View style={styles.emptyContainer}>
          <Text style={styles.emptyText}>No records found</Text>
          <Text style={styles.emptySubtext}>
            Pull down to refresh or go back to create a new record
          </Text>
        </View>
      ) : filteredRecords.length === 0 && searchQuery.trim() ? (
        <View style={styles.emptyContainer}>
          <Text style={styles.emptyText}>No records match your search</Text>
          <Text style={styles.emptySubtext}>
            Try a different search term or clear the search
          </Text>
        </View>
      ) : (
        <ScrollView
          style={styles.scrollView}
          refreshControl={
            <RefreshControl refreshing={refreshing} onRefresh={handleRefresh} />
          }
        >
          {filteredRecords.map((record) => (
            <TouchableOpacity
              key={String(record.Id)}
              style={styles.recordCard}
              onPress={() => onRecordPress(record)}
            >
              <View style={styles.recordContent}>
                {listName === 'Access Cards' ? (
                  renderAccessCardRecord(record)
                ) : listName === 'Assets' ? (
                  renderAssetRecord(record)
                ) : listName === 'Employees' ? (
                  renderEmployeeRecord(record)
                ) : (
                  <>
                    <Text style={styles.recordTitle}>{getDisplayValue(record)}</Text>
                    <Text style={styles.recordId}>ID: {record.Id}</Text>
                  </>
                )}
              </View>
              {/* {listName !== 'Employees' && (
                <TouchableOpacity
                  style={styles.deleteButton}
                  onPress={(e) => {
                    e.stopPropagation();
                    handleDelete(record.Id);
                  }}
                >
                  <Text style={styles.deleteButtonText}>🗑️</Text>
                </TouchableOpacity>
              )} */}
            </TouchableOpacity>
          ))}
        </ScrollView>
      )}
    </SafeAreaView>
  );
};

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#f5f5f5',
  },
  header: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: 15,
    backgroundColor: '#fff',
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
  },
  backButton: {
    padding: 5,
  },
  backButtonText: {
    fontSize: 16,
    color: '#0078d4',
    fontWeight: '600',
  },
  headerTitle: {
    fontSize: 18,
    fontWeight: 'bold',
    color: '#333',
    flex: 1,
    textAlign: 'center',
  },
  placeholder: {
    width: 60,
  },
  addButton: {
    padding: 5,
  },
  addButtonText: {
    fontSize: 16,
    color: '#0078d4',
    fontWeight: '600',
  },
  searchContainer: {
    flexDirection: 'row',
    padding: 15,
    backgroundColor: '#fff',
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
    alignItems: 'center',
  },
  searchInput: {
    flex: 1,
    height: 40,
    backgroundColor: '#f5f5f5',
    borderRadius: 8,
    paddingHorizontal: 15,
    fontSize: 14,
    color: '#333',
    borderWidth: 1,
    borderColor: '#e0e0e0',
  },
  clearButton: {
    marginLeft: 10,
    padding: 8,
    justifyContent: 'center',
    alignItems: 'center',
  },
  clearButtonText: {
    fontSize: 16,
    color: '#666',
    fontWeight: 'bold',
  },
  countContainer: {
    padding: 15,
    backgroundColor: '#fff',
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
  },
  countText: {
    fontSize: 14,
    color: '#666',
  },
  loadingContainer: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 40,
  },
  loadingText: {
    marginTop: 10,
    fontSize: 16,
    color: '#666',
  },
  scrollView: {
    flex: 1,
  },
  emptyContainer: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 40,
  },
  emptyText: {
    fontSize: 18,
    color: '#999',
    marginBottom: 8,
  },
  emptySubtext: {
    fontSize: 14,
    color: '#bbb',
    textAlign: 'center',
  },
  recordCard: {
    flexDirection: 'row',
    backgroundColor: '#fff',
    marginHorizontal: 15,
    marginVertical: 8,
    padding: 15,
    borderRadius: 8,
    borderWidth: 1,
    borderColor: '#e0e0e0',
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 1 },
    shadowOpacity: 0.1,
    shadowRadius: 2,
    elevation: 2,
  },
  recordContent: {
    flex: 1,
  },
  recordTitle: {
    fontSize: 16,
    fontWeight: '600',
    color: '#333',
    marginBottom: 4,
  },
  recordId: {
    fontSize: 12,
    color: '#999',
  },
  deleteButton: {
    justifyContent: 'center',
    paddingLeft: 15,
  },
  deleteButtonText: {
    fontSize: 20,
  },
  accessCardContent: {
    flex: 1,
  },
  accessCardRow: {
    flexDirection: 'row',
    marginBottom: 8,
    alignItems: 'flex-start',
  },
  accessCardLabel: {
    fontSize: 14,
    fontWeight: '600',
    color: '#666',
    width: 120,
    marginRight: 8,
  },
  accessCardValue: {
    fontSize: 14,
    color: '#333',
    flex: 1,
  },
  errorContainer: {
    flex: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 20,
  },
  errorText: {
    color: '#c62828',
    fontSize: 14,
    marginBottom: 20,
    textAlign: 'center',
  },
  retryButton: {
    backgroundColor: '#f44336',
    padding: 12,
    borderRadius: 6,
  },
  retryButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
});

export default ListScreen;
