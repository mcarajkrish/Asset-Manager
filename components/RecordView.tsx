import React, { useState } from 'react';
import {
  View,
  Text,
  TextInput,
  TouchableOpacity,
  StyleSheet,
  ScrollView,
  Alert,
  ActivityIndicator,
} from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import SharePointService, { SessionTimeoutError } from '../services/sharepointService';

interface RecordViewProps {
  sharePointService: SharePointService;
  listName: string;
  record: any;
  onClose: () => void;
  onRecordUpdated?: () => void;
  onRecordDeleted?: () => void;
}

const RecordView: React.FC<RecordViewProps> = ({
  sharePointService,
  listName,
  record,
  onClose,
  onRecordUpdated,
  onRecordDeleted,
}) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editedFields, setEditedFields] = useState<{ [key: string]: any }>({});
  const [loading, setLoading] = useState(false);
  
  // Check if this is an organization user (not a SharePoint list item)
  // Organization users have UserPrincipalName or Mail but no SharePoint list item structure
  const isOrganizationUser = listName === 'Employees' && (
    record.UserPrincipalName || 
    (record.Mail && !record.hasOwnProperty('AccessCardNo') && !record.hasOwnProperty('AssetId'))
  );

  // Filter out metadata and system fields
  const getDisplayFields = () => {
    return Object.keys(record).filter(
      (key) =>
        !key.startsWith('_') &&
        !key.startsWith('__') &&
        key !== '__metadata' &&
        key !== 'odata' &&
        !key.includes('@odata') &&
        record[key] !== null &&
        record[key] !== undefined
    );
  };

  const handleFieldChange = (fieldName: string, value: any) => {
    setEditedFields((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const handleSave = async () => {
    // Prevent updates for organization users (not SharePoint list items)
    if (isOrganizationUser) {
      Alert.alert(
        'Info', 
        'Organization users cannot be edited from this app. User information is managed in Azure AD.'
      );
      setIsEditing(false);
      setEditedFields({});
      return;
    }
    
    try {
      setLoading(true);
      
      // Only send changed fields
      const fieldsToUpdate: { [key: string]: any } = {};
      Object.keys(editedFields).forEach((key) => {
        fieldsToUpdate[key] = editedFields[key];
      });

      if (Object.keys(fieldsToUpdate).length === 0) {
        Alert.alert('Info', 'No changes to save');
        setIsEditing(false);
        setLoading(false);
        return;
      }

      await sharePointService.updateRecord(listName, record.Id, fieldsToUpdate);
      Alert.alert('Success', 'Record updated successfully!');
      setIsEditing(false);
      setEditedFields({});
      
      if (onRecordUpdated) {
        onRecordUpdated();
      }
    } catch (error: any) {
      // Handle session timeout - don't show alert as App.tsx will handle it
      if (error instanceof SessionTimeoutError) {
        return;
      }
      Alert.alert('Error', error.message || 'Failed to update record');
      console.error('Update error:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleDelete = () => {
    // Prevent deletes for organization users (not SharePoint list items)
    if (isOrganizationUser) {
      Alert.alert(
        'Info', 
        'Organization users cannot be deleted from this app. User management is done in Azure AD.'
      );
      return;
    }
    
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
              setLoading(true);
              await sharePointService.deleteRecord(listName, record.Id);
              Alert.alert('Success', 'Record deleted successfully!');
              if (onRecordDeleted) {
                onRecordDeleted();
              }
            } catch (error: any) {
              // Handle session timeout - don't show alert as App.tsx will handle it
              if (error instanceof SessionTimeoutError) {
                return;
              }
              Alert.alert('Error', error.message || 'Failed to delete record');
              console.error('Delete error:', error);
            } finally {
              setLoading(false);
            }
          },
        },
      ]
    );
  };

  const getFieldValue = (fieldName: string): string => {
    if (isEditing && editedFields.hasOwnProperty(fieldName)) {
      return String(editedFields[fieldName] ?? '');
    }
    const value = record[fieldName];
    if (value === null || value === undefined) {
      return '';
    }
    if (typeof value === 'object') {
      return JSON.stringify(value);
    }
    return String(value);
  };

  const displayFields = getDisplayFields();

  // Get context-specific title based on list name
  const getTitle = (): string => {
    switch (listName) {
      case 'Employees':
        return 'Employee Details';
      case 'Assets':
        return 'Asset Details';
      case 'Access Cards':
        return 'Access Card Details';
      default:
        return `${listName} Details`;
    }
  };

  return (
    <SafeAreaView style={styles.container} edges={['top', 'bottom']}>
      <View style={styles.header}>
        <TouchableOpacity onPress={onClose} style={styles.backButton}>
          <Text style={styles.backButtonText}>←</Text>
        </TouchableOpacity>
        <Text style={styles.headerTitle}>{getTitle()}</Text>
        {!isEditing && isOrganizationUser && <View style={styles.placeholder} />}
      </View>

      <ScrollView style={styles.scrollView}>
        <View style={styles.content}>
          {displayFields.map((fieldName) => {
            const isEditable = fieldName !== 'Id' && fieldName !== 'Created' && fieldName !== 'Modified';
            const value = getFieldValue(fieldName);

            return (
              <View key={fieldName} style={styles.fieldContainer}>
                <Text style={styles.fieldLabel}>
                  {fieldName}
                  {fieldName === 'Id' && <Text style={styles.readOnly}> (Read-only)</Text>}
                </Text>
                {isEditing && isEditable ? (
                  <TextInput
                    style={styles.fieldInput}
                    value={value}
                    onChangeText={(text) => handleFieldChange(fieldName, text)}
                    multiline={value.length > 50}
                    numberOfLines={value.length > 50 ? 3 : 1}
                  />
                ) : (
                  <Text style={styles.fieldValue}>
                    {value || <Text style={styles.emptyValue}>—</Text>}
                  </Text>
                )}
              </View>
            );
          })}

          {isEditing && (
            <View style={styles.editActions}>
              <TouchableOpacity
                style={[styles.actionButton, styles.cancelButton]}
                onPress={() => {
                  setIsEditing(false);
                  setEditedFields({});
                }}
                disabled={loading}
              >
                <Text style={styles.cancelButtonText}>Cancel</Text>
              </TouchableOpacity>
              <TouchableOpacity
                style={[styles.actionButton, styles.saveButton, loading && styles.buttonDisabled]}
                onPress={handleSave}
                disabled={loading}
              >
                {loading ? (
                  <ActivityIndicator color="#fff" />
                ) : (
                  <Text style={styles.saveButtonText}>Save Changes</Text>
                )}
              </TouchableOpacity>
            </View>
          )}
        </View>
      </ScrollView>
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
  scrollView: {
    flex: 1,
  },
  content: {
    padding: 20,
  },
  fieldContainer: {
    marginBottom: 20,
    backgroundColor: '#fff',
    padding: 15,
    borderRadius: 8,
    borderWidth: 1,
    borderColor: '#e0e0e0',
  },
  fieldLabel: {
    fontSize: 12,
    fontWeight: '600',
    color: '#666',
    marginBottom: 8,
    textTransform: 'uppercase',
  },
  readOnly: {
    fontSize: 11,
    color: '#999',
    fontWeight: 'normal',
    textTransform: 'none',
  },
  fieldValue: {
    fontSize: 16,
    color: '#333',
    lineHeight: 22,
  },
  emptyValue: {
    color: '#999',
    fontStyle: 'italic',
  },
  fieldInput: {
    fontSize: 16,
    color: '#333',
    borderWidth: 1,
    borderColor: '#0078d4',
    borderRadius: 6,
    padding: 10,
    backgroundColor: '#fff',
  },
  editActions: {
    flexDirection: 'row',
    gap: 10,
    marginTop: 10,
  },
  actionButton: {
    flex: 1,
    padding: 15,
    borderRadius: 8,
    alignItems: 'center',
  },
  cancelButton: {
    backgroundColor: '#f5f5f5',
    borderWidth: 1,
    borderColor: '#ddd',
  },
  cancelButtonText: {
    color: '#333',
    fontSize: 16,
    fontWeight: '600',
  },
  saveButton: {
    backgroundColor: '#0078d4',
  },
  saveButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: '600',
  },
  buttonDisabled: {
    opacity: 0.6,
  },
  placeholder: {
    width: 40,
  },
});

export default RecordView;
