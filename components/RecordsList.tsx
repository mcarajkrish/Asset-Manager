import React, { useState, useEffect } from 'react';
import {
  View,
  Text,
  TouchableOpacity,
  StyleSheet,
  ScrollView,
  ActivityIndicator,
  Alert,
  RefreshControl,
} from 'react-native';
import SharePointService from '../services/sharepointService';
import RecordView from './RecordView';

interface RecordsListProps {
  sharePointService: SharePointService;
  listName: string;
  onRecordUpdated?: () => void;
}

interface Record {
  Id: number;
  Title?: string;
  [key: string]: any;
}

const RecordsList: React.FC<RecordsListProps> = ({
  sharePointService,
  listName,
  onRecordUpdated,
}) => {
  const [records, setRecords] = useState<Record[]>([]);
  const [loading, setLoading] = useState(false);
  const [refreshing, setRefreshing] = useState(false);
  const [selectedRecord, setSelectedRecord] = useState<Record | null>(null);
  const [error, setError] = useState<string | null>(null);

  const loadRecords = async () => {
    try {
      setLoading(true);
      setError(null);
      const items = await sharePointService.getRecords(listName);
      setRecords(items);
    } catch (error: any) {
      const errorMessage = error.message || 'Failed to load records';
      setError(errorMessage);
      console.error('Error loading records:', error);
      Alert.alert('Error', errorMessage);
    } finally {
      setLoading(false);
      setRefreshing(false);
    }
  };

  useEffect(() => {
    if (listName) {
      loadRecords();
    }
  }, [listName]);

  const handleRefresh = () => {
    setRefreshing(true);
    loadRecords();
  };

  const handleRecordDeleted = async () => {
    setSelectedRecord(null);
    await loadRecords();
    if (onRecordUpdated) {
      onRecordUpdated();
    }
  };

  const handleRecordUpdated = async () => {
    setSelectedRecord(null);
    await loadRecords();
    if (onRecordUpdated) {
      onRecordUpdated();
    }
  };

  if (selectedRecord) {
    return (
      <RecordView
        sharePointService={sharePointService}
        listName={listName}
        record={selectedRecord}
        onClose={() => setSelectedRecord(null)}
        onRecordUpdated={handleRecordUpdated}
        onRecordDeleted={handleRecordDeleted}
      />
    );
  }

  const getDisplayValue = (record: Record): string => {
    // Try to find a meaningful display value
    if (record.Title) return record.Title;
    
    // Try common field names
    const displayFields = ['Name', 'AssetName', 'EmployeeName', 'CardNumber', 'AccessCardNumber'];
    for (const field of displayFields) {
      if (record[field]) {
        return String(record[field]);
      }
    }
    
    // Fall back to first non-ID field
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

  if (loading && records.length === 0) {
    return (
      <View style={styles.container}>
        <View style={styles.loadingContainer}>
          <ActivityIndicator size="large" color="#0078d4" />
          <Text style={styles.loadingText}>Loading records...</Text>
        </View>
      </View>
    );
  }

  return (
    <View style={styles.container}>
      <View style={styles.header}>
        <Text style={styles.title}>Records in: {listName}</Text>
        <Text style={styles.count}>{records.length} record(s)</Text>
      </View>

      {error && (
        <View style={styles.errorContainer}>
          <Text style={styles.errorText}>{error}</Text>
          <TouchableOpacity
            style={styles.retryButton}
            onPress={loadRecords}
          >
            <Text style={styles.retryButtonText}>Retry</Text>
          </TouchableOpacity>
        </View>
      )}

      <ScrollView
        style={styles.scrollView}
        refreshControl={
          <RefreshControl refreshing={refreshing} onRefresh={handleRefresh} />
        }
      >
        {records.length === 0 && !loading ? (
          <View style={styles.emptyContainer}>
            <Text style={styles.emptyText}>No records found</Text>
            <Text style={styles.emptySubtext}>
              Use the form below to add a new record
            </Text>
          </View>
        ) : (
          records.map((record) => (
            <TouchableOpacity
              key={record.Id}
              style={styles.recordCard}
              onPress={() => setSelectedRecord(record)}
            >
              <View style={styles.recordContent}>
                <Text style={styles.recordTitle}>{getDisplayValue(record)}</Text>
                <Text style={styles.recordId}>ID: {record.Id}</Text>
              </View>
            </TouchableOpacity>
          ))
        )}
      </ScrollView>
    </View>
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
  title: {
    fontSize: 18,
    fontWeight: 'bold',
    color: '#333',
  },
  count: {
    fontSize: 14,
    color: '#666',
  },
  scrollView: {
    flex: 1,
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
  emptyContainer: {
    padding: 40,
    alignItems: 'center',
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
  errorContainer: {
    backgroundColor: '#ffebee',
    margin: 15,
    padding: 15,
    borderRadius: 8,
    borderLeftWidth: 4,
    borderLeftColor: '#f44336',
  },
  errorText: {
    color: '#c62828',
    fontSize: 14,
    marginBottom: 10,
  },
  retryButton: {
    backgroundColor: '#f44336',
    padding: 8,
    borderRadius: 4,
    alignSelf: 'flex-start',
  },
  retryButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
});

export default RecordsList;
