import React, { useState, useCallback } from 'react';
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
      
      // For Assets list, refresh employees cache first to ensure correct assignee resolution
      if (listName === 'Assets' && onRefreshEmployees) {
        await onRefreshEmployees();
      }
      
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
    // Refresh employees cache first for Employees list or Assets list (to ensure correct assignee resolution)
    if (onRefreshEmployees && (listName === 'Employees' || listName === 'Assets')) {
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

  // Sort records based on list type
  const sortedRecords = React.useMemo(() => {
    if (listName === 'Assets') {
      return [...filteredRecords].sort((a, b) => {
        const assetIdA = getFieldValue(a, ['AssetID']);
        const assetIdB = getFieldValue(b, ['AssetID']);
        // Try numeric comparison first, then string comparison
        const numA = parseInt(assetIdA, 10);
        const numB = parseInt(assetIdB, 10);
        if (!isNaN(numA) && !isNaN(numB)) {
          return numA - numB;
        }
        return assetIdA.localeCompare(assetIdB, undefined, { numeric: true, sensitivity: 'base' });
      });
    } else if (listName === 'Access Cards') {
      return [...filteredRecords].sort((a, b) => {
        const cardNoA = getFieldValue(a, ['AccessCardNo']);
        const cardNoB = getFieldValue(b, ['AccessCardNo']);
        // Try numeric comparison first, then string comparison
        const numA = parseInt(cardNoA, 10);
        const numB = parseInt(cardNoB, 10);
        if (!isNaN(numA) && !isNaN(numB)) {
          return numA - numB;
        }
        return cardNoA.localeCompare(cardNoB, undefined, { numeric: true, sensitivity: 'base' });
      });
    }
    return filteredRecords;
  }, [filteredRecords, listName]);

  const getCardStatus = (record: Record): string => {
    return getFieldValue(record, ['CardStatus']).toLowerCase();
  };

  const getDeviceStatus = (record: Record): string => {
    return getFieldValue(record, ['DeviceStatus']).toLowerCase();
  };

  const getCardBackgroundColor = (status: string): string => {
    const normalizedStatus = status.toLowerCase().trim();
    if (normalizedStatus === 'available') {
      return '#e8f5e9'; // Light green
    } else if (normalizedStatus === 'assigned') {
      return '#e3f2fd'; // Light blue
    }
    return '#fff'; // Default white
  };

  const renderAccessCardRecord = (record: Record) => {

    return (
      <View style={styles.accessCardContent}>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Access Card No:</Text>
          <Text style={styles.accessCardValue}>{record['AccessCardNo'] || '-'}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Assignee:</Text>
          <Text style={styles.accessCardValue}>{record['Employee'] || '-'}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Emp Id:</Text>
          <Text style={styles.accessCardValue}>{record['EmpId'] || '-'}</Text>
        </View>
      </View>
    );
  };

  const renderEmployeeRecord = (record: Record) => {
    
    return (
      <View style={styles.accessCardContent}>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Name:</Text>
          <Text style={styles.accessCardValue}>{record['displayName'] || '-'}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Mail Id:</Text>
          <Text style={styles.accessCardValue}>{record['mail'] || '-'}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Designation:</Text>
          <Text style={styles.accessCardValue}>{record['jobTitle'] || '-'}</Text>
        </View>
      </View>
    );
  };

  const renderAssetRecord = (record: Record) => {

    return (
      <View style={styles.accessCardContent}>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Asset Id:</Text>
          <Text style={styles.accessCardValue}>{record['AssetID'] || '-'}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Asset:</Text>
          <Text style={styles.accessCardValue}>{record['Company'] || ''} {record['Model'] || ''}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Serial Number:</Text>
          <Text style={styles.accessCardValue}>{record['Serial Number'] || '-'}</Text>
        </View>
        <View style={styles.accessCardRow}>
          <Text style={styles.accessCardLabel}>Assignee:</Text>
          <Text style={styles.accessCardValue}>{record['Id'] || '-'}</Text>
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
          {sortedRecords.map((record) => {
            let cardBackgroundColor = '#fff';
            if (listName === 'Access Cards') {
              cardBackgroundColor = getCardBackgroundColor(getCardStatus(record));
            } else if (listName === 'Assets') {
              cardBackgroundColor = getCardBackgroundColor(getDeviceStatus(record));
            }
            
            return (
              <TouchableOpacity
                key={String(record.Id || record.id)}
                style={[styles.recordCard, { backgroundColor: cardBackgroundColor }]}
                onPress={() => onRecordPress(record)}
              >
                <View style={styles.recordContent}>
                  {listName === 'Access Cards' ? (
                    renderAccessCardRecord(record)
                  ) : listName === 'Assets' ? (
                    renderAssetRecord(record)
                  ) : listName === 'Employees' ? (
                    renderEmployeeRecord(record)
                  ) : null}
                </View>
              </TouchableOpacity>
            );
          })}
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
