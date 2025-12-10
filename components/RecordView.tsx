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
  Modal,
  FlatList,
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
  employees?: any[];
}

const RecordView: React.FC<RecordViewProps> = ({
  sharePointService,
  listName,
  record,
  onClose,
  onRecordUpdated,
  onRecordDeleted,
  employees = [],
}) => {
  const [isEditing, setIsEditing] = useState(false);
  const [editedFields, setEditedFields] = useState<{ [key: string]: any }>({});
  const [loading, setLoading] = useState(false);
  const [showEmployeeModal, setShowEmployeeModal] = useState(false);
  const [assigning, setAssigning] = useState(false);
  
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
    
    // Get context-specific delete message
    const getDeleteTitle = () => {
      if (listName === 'Assets') return 'Delete Asset';
      if (listName === 'Access Cards') return 'Delete Access Card';
      return 'Delete Record';
    };
    
    const getDeleteMessage = () => {
      if (listName === 'Assets') {
        return 'Are you sure you want to delete this Asset? This action cannot be undone.';
      }
      if (listName === 'Access Cards') {
        return 'Are you sure you want to delete this Access Card? This action cannot be undone.';
      }
      return 'Are you sure you want to delete this record? This action cannot be undone.';
    };
    
    const getSuccessMessage = () => {
      if (listName === 'Assets') return 'Asset deleted successfully!';
      if (listName === 'Access Cards') return 'Access Card deleted successfully!';
      return 'Record deleted successfully!';
    };
    
    Alert.alert(
      getDeleteTitle(),
      getDeleteMessage(),
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Delete',
          style: 'destructive',
          onPress: async () => {
            try {
              setLoading(true);
              await sharePointService.deleteRecord(listName, record.Id);
              // Navigate back immediately
              if (onRecordDeleted) {
                onRecordDeleted();
              }
              // Show success message after a brief delay to allow navigation
              setTimeout(() => {
                Alert.alert('Success', getSuccessMessage());
              }, 300);
            } catch (error: any) {
              // Handle session timeout - don't show alert as App.tsx will handle it
              if (error instanceof SessionTimeoutError) {
                setLoading(false);
                return;
              }
              setLoading(false);
              Alert.alert('Error', error.message || 'Failed to delete record');
              console.error('Delete error:', error);
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
    
    // Check if this is a LookupId field and try to get display name
    if (fieldName.endsWith('LookupId')) {
      const displayNameField = `${fieldName}_displayName`;
      if (record[displayNameField]) {
        return String(record[displayNameField]);
      }
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

  // Check if Access Card is assigned
  const isAccessCardAssigned = () => {
    if (listName !== 'Access Cards') return false;
    const employeeLookupId = record['EmployeeLookupId'];
    return employeeLookupId && employeeLookupId !== null && employeeLookupId !== undefined && employeeLookupId !== '';
  };

  // Check if Asset is assigned
  const isAssetAssigned = () => {
    if (listName !== 'Assets') return false;
    const assigneeLookupId = record['field_2LookupId'];
    return assigneeLookupId && assigneeLookupId !== null && assigneeLookupId !== undefined && assigneeLookupId !== '';
  };

  // Handle assign Access Card
  const handleAssignAccessCard = async (employeeId: string) => {
    try {
      setAssigning(true);
      setShowEmployeeModal(false);
      
      // Find employee in the employees list
      const employee = employees.find(emp => emp.id === employeeId);
      if (!employee) {
        Alert.alert('Error', 'Employee not found');
        return;
      }

      // Resolve Graph API user ID to SharePoint User Information List ID
      // Check if method exists (in case of caching issues)
      if (typeof sharePointService.resolveSharePointUserId !== 'function') {
        Alert.alert('Error', 'Assignment feature not available. Please restart the app.');
        setAssigning(false);
        return;
      }
      
      const sharePointUserId = await sharePointService.resolveSharePointUserId(employeeId);
      if (!sharePointUserId) {
        // Fallback: try using email
        if (employee.mail) {
          const sharePointUserIdByEmail = await sharePointService.resolveSharePointUserId(employee.mail);
          if (sharePointUserIdByEmail) {
            await sharePointService.updateRecord(listName, record.Id, {
              EmployeeLookupId: sharePointUserIdByEmail,
              CardStatus: 'Assigned',
            });
            Alert.alert('Success', 'Access Card assigned successfully!');
            if (onRecordUpdated) {
              onRecordUpdated();
            }
            return;
          }
        }
        Alert.alert('Error', 'Could not resolve employee in SharePoint. Please try again.');
        return;
      }

      await sharePointService.updateRecord(listName, record.Id, {
        EmployeeLookupId: sharePointUserId,
        CardStatus: 'Assigned',
      });
      
      Alert.alert('Success', 'Access Card assigned successfully!');
      if (onRecordUpdated) {
        onRecordUpdated();
      }
    } catch (error: any) {
      if (error instanceof SessionTimeoutError) {
        return;
      }
      Alert.alert('Error', error.message || 'Failed to assign Access Card');
      console.error('Assign error:', error);
    } finally {
      setAssigning(false);
    }
  };

  // Handle unassign Access Card
  const handleUnassignAccessCard = async () => {
    Alert.alert(
      'Unassign Access Card',
      'Are you sure you want to unassign this Access Card?',
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Unassign',
          style: 'destructive',
          onPress: async () => {
            try {
              setAssigning(true);
              await sharePointService.updateRecord(listName, record.Id, {
                EmployeeLookupId: null,
                CardStatus: 'Available',
              });
              Alert.alert('Success', 'Access Card unassigned successfully!');
              if (onRecordUpdated) {
                onRecordUpdated();
              }
            } catch (error: any) {
              if (error instanceof SessionTimeoutError) {
                return;
              }
              Alert.alert('Error', error.message || 'Failed to unassign Access Card');
              console.error('Unassign error:', error);
            } finally {
              setAssigning(false);
            }
          },
        },
      ]
    );
  };

  // Handle assign Asset
  const handleAssignAsset = async (employeeId: string) => {
    try {
      setAssigning(true);
      setShowEmployeeModal(false);
      
      const employee = employees.find(emp => emp.id === employeeId);
      if (!employee) {
        Alert.alert('Error', 'Employee not found');
        return;
      }

      // Resolve Graph API user ID to SharePoint User Information List ID
      // Check if method exists (in case of caching issues)
      if (typeof sharePointService.resolveSharePointUserId !== 'function') {
        Alert.alert('Error', 'Assignment feature not available. Please restart the app.');
        setAssigning(false);
        return;
      }
      
      const sharePointUserId = await sharePointService.resolveSharePointUserId(employeeId);
      if (!sharePointUserId) {
        // Fallback: try using email
        if (employee.mail) {
          const sharePointUserIdByEmail = await sharePointService.resolveSharePointUserId(employee.mail);
          if (sharePointUserIdByEmail) {
            await sharePointService.updateRecord(listName, record.Id, {
              field_2LookupId: sharePointUserIdByEmail,
              DeviceStatus: 'Assigned',
            });
            Alert.alert('Success', 'Asset assigned successfully!');
            if (onRecordUpdated) {
              onRecordUpdated();
            }
            return;
          }
        }
        Alert.alert('Error', 'Could not resolve employee in SharePoint. Please try again.');
        return;
      }

      await sharePointService.updateRecord(listName, record.Id, {
        field_2LookupId: sharePointUserId,
        DeviceStatus: 'Assigned',
      });
      
      Alert.alert('Success', 'Asset assigned successfully!');
      if (onRecordUpdated) {
        onRecordUpdated();
      }
    } catch (error: any) {
      if (error instanceof SessionTimeoutError) {
        return;
      }
      Alert.alert('Error', error.message || 'Failed to assign Asset');
      console.error('Assign error:', error);
    } finally {
      setAssigning(false);
    }
  };

  // Handle unassign Asset
  const handleUnassignAsset = async () => {
    Alert.alert(
      'Unassign Asset',
      'Are you sure you want to unassign this Asset?',
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Unassign',
          style: 'destructive',
          onPress: async () => {
            try {
              setAssigning(true);
              await sharePointService.updateRecord(listName, record.Id, {
                field_2LookupId: null,
                DeviceStatus: 'Available',
              });
              Alert.alert('Success', 'Asset unassigned successfully!');
              if (onRecordUpdated) {
                onRecordUpdated();
              }
            } catch (error: any) {
              if (error instanceof SessionTimeoutError) {
                return;
              }
              Alert.alert('Error', error.message || 'Failed to unassign Asset');
              console.error('Unassign error:', error);
            } finally {
              setAssigning(false);
            }
          },
        },
      ]
    );
  };

  // Render employee selection modal
  const renderEmployeeModal = () => {
    return (
      <Modal
        visible={showEmployeeModal}
        animationType="slide"
        transparent={true}
        onRequestClose={() => setShowEmployeeModal(false)}
      >
        <View style={styles.modalOverlay}>
          <View style={styles.modalContent}>
            <View style={styles.modalHeader}>
              <Text style={styles.modalTitle}>Select Employee</Text>
              <TouchableOpacity
                onPress={() => setShowEmployeeModal(false)}
                style={styles.modalCloseButton}
              >
                <Text style={styles.modalCloseText}>✕</Text>
              </TouchableOpacity>
            </View>
            <FlatList
              data={employees}
              keyExtractor={(item) => item.id}
              renderItem={({ item }) => (
                <TouchableOpacity
                  style={styles.employeeItem}
                  onPress={() => {
                    if (listName === 'Access Cards') {
                      handleAssignAccessCard(item.id);
                    } else if (listName === 'Assets') {
                      handleAssignAsset(item.id);
                    }
                  }}
                >
                  <Text style={styles.employeeName}>{item.displayName}</Text>
                  <Text style={styles.employeeEmail}>{item.mail || ''}</Text>
                </TouchableOpacity>
              )}
              ListEmptyComponent={
                <View style={styles.emptyEmployeeList}>
                  <Text style={styles.emptyEmployeeText}>No employees available</Text>
                </View>
              }
            />
          </View>
        </View>
      </Modal>
    );
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

          {/* Assign/Unassign Actions for Access Cards */}
          {!isEditing && listName === 'Access Cards' && (
            <View style={styles.assignActions}>
              {isAccessCardAssigned() ? (
                <TouchableOpacity
                  style={[styles.assignButton, styles.unassignButton, assigning && styles.buttonDisabled]}
                  onPress={handleUnassignAccessCard}
                  disabled={assigning}
                >
                  {assigning ? (
                    <ActivityIndicator color="#fff" />
                  ) : (
                    <Text style={styles.assignButtonText}>Unassign Access Card</Text>
                  )}
                </TouchableOpacity>
              ) : (
                <TouchableOpacity
                  style={[styles.assignButton, styles.assignButtonStyle, assigning && styles.buttonDisabled]}
                  onPress={() => setShowEmployeeModal(true)}
                  disabled={assigning || employees.length === 0}
                >
                  {assigning ? (
                    <ActivityIndicator color="#fff" />
                  ) : (
                    <Text style={styles.assignButtonText}>Assign Access Card</Text>
                  )}
                </TouchableOpacity>
              )}
            </View>
          )}

          {/* Assign/Unassign Actions for Assets */}
          {!isEditing && listName === 'Assets' && (
            <View style={styles.assignActions}>
              {isAssetAssigned() ? (
                <TouchableOpacity
                  style={[styles.assignButton, styles.unassignButton, assigning && styles.buttonDisabled]}
                  onPress={handleUnassignAsset}
                  disabled={assigning}
                >
                  {assigning ? (
                    <ActivityIndicator color="#fff" />
                  ) : (
                    <Text style={styles.assignButtonText}>Unassign Asset</Text>
                  )}
                </TouchableOpacity>
              ) : (
                <TouchableOpacity
                  style={[styles.assignButton, styles.assignButtonStyle, assigning && styles.buttonDisabled]}
                  onPress={() => setShowEmployeeModal(true)}
                  disabled={assigning || employees.length === 0}
                >
                  {assigning ? (
                    <ActivityIndicator color="#fff" />
                  ) : (
                    <Text style={styles.assignButtonText}>Assign Asset</Text>
                  )}
                </TouchableOpacity>
              )}
            </View>
          )}

          {/* Delete Button for Assets */}
          {!isEditing && listName === 'Assets' && (
            <View style={styles.deleteActions}>
              <TouchableOpacity
                style={[styles.deleteButton, loading && styles.buttonDisabled]}
                onPress={handleDelete}
                disabled={loading}
              >
                {loading ? (
                  <ActivityIndicator color="#fff" />
                ) : (
                  <Text style={styles.deleteButtonText}>Delete Asset</Text>
                )}
              </TouchableOpacity>
            </View>
          )}
        </View>
      </ScrollView>
      {renderEmployeeModal()}
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
  assignActions: {
    marginTop: 20,
  },
  assignButton: {
    padding: 15,
    borderRadius: 8,
    alignItems: 'center',
  },
  assignButtonStyle: {
    backgroundColor: '#4CAF50',
  },
  unassignButton: {
    backgroundColor: '#f44336',
  },
  assignButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: '600',
  },
  deleteActions: {
    marginTop: 20,
  },
  deleteButton: {
    padding: 15,
    borderRadius: 8,
    alignItems: 'center',
    backgroundColor: '#d32f2f',
  },
  deleteButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: '600',
  },
  modalOverlay: {
    flex: 1,
    backgroundColor: 'rgba(0, 0, 0, 0.5)',
    justifyContent: 'flex-end',
  },
  modalContent: {
    backgroundColor: '#fff',
    borderTopLeftRadius: 20,
    borderTopRightRadius: 20,
    maxHeight: '80%',
    paddingBottom: 20,
  },
  modalHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: 20,
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
  },
  modalTitle: {
    fontSize: 18,
    fontWeight: 'bold',
    color: '#333',
  },
  modalCloseButton: {
    padding: 5,
  },
  modalCloseText: {
    fontSize: 24,
    color: '#666',
  },
  employeeItem: {
    padding: 15,
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
  },
  employeeName: {
    fontSize: 16,
    fontWeight: '600',
    color: '#333',
    marginBottom: 4,
  },
  employeeEmail: {
    fontSize: 14,
    color: '#666',
  },
  emptyEmployeeList: {
    padding: 40,
    alignItems: 'center',
  },
  emptyEmployeeText: {
    fontSize: 16,
    color: '#999',
  },
});

export default RecordView;
