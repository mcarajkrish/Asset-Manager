import React, { useState } from 'react';
import {
  View,
  Text,
  TextInput,
  TouchableOpacity,
  StyleSheet,
  Alert,
  ScrollView,
  ActivityIndicator,
} from 'react-native';
import SharePointService, { SessionTimeoutError } from '../services/sharepointService';

interface RecordFormProps {
  sharePointService: SharePointService;
  listName: string;
  onRecordInserted?: () => void;
}

interface FormField {
  name: string;
  label: string;
  type?: 'text' | 'number' | 'date' | 'email';
  required?: boolean;
}

const RecordForm: React.FC<RecordFormProps> = ({
  sharePointService,
  listName,
  onRecordInserted,
}) => {
  const [fields, setFields] = useState<FormField[]>([
    { name: 'Title', label: 'Title', type: 'text', required: true },
  ]);
  const [formData, setFormData] = useState<{ [key: string]: string }>({});
  const [loading, setLoading] = useState(false);
  const [customFields, setCustomFields] = useState<{ name: string; value: string }[]>([]);

  const handleInputChange = (fieldName: string, value: string) => {
    setFormData((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const addCustomField = () => {
    setCustomFields((prev) => [...prev, { name: '', value: '' }]);
  };

  const updateCustomField = (index: number, field: 'name' | 'value', value: string) => {
    setCustomFields((prev) =>
      prev.map((item, i) => (i === index ? { ...item, [field]: value } : item))
    );
  };

  const removeCustomField = (index: number) => {
    setCustomFields((prev) => prev.filter((_, i) => i !== index));
  };

  const handleSubmit = async () => {
    try {
      setLoading(true);

      // Validate required fields
      const requiredFields = fields.filter((f) => f.required);
      for (const field of requiredFields) {
        if (!formData[field.name]?.trim()) {
          Alert.alert('Validation Error', `${field.label} is required`);
          setLoading(false);
          return;
        }
      }

      // Prepare data object
      const recordData: { [key: string]: any } = { ...formData };

      // Add custom fields
      customFields.forEach((customField) => {
        if (customField.name && customField.value) {
          recordData[customField.name] = customField.value;
        }
      });

      // Insert record
      await sharePointService.insertRecord(listName, recordData);

      // Reset form
      setFormData({});
      setCustomFields([]);
      
      Alert.alert('Success', 'Record inserted successfully!', [
        {
          text: 'OK',
          onPress: () => {
            if (onRecordInserted) {
              onRecordInserted();
            }
          },
        },
      ]);
    } catch (error: any) {
      // Handle session timeout - don't show alert as App.tsx will handle it
      if (error instanceof SessionTimeoutError) {
        return;
      }
      Alert.alert('Error', error.message || 'Failed to insert record');
      console.error('Insert error:', error);
    } finally {
      setLoading(false);
    }
  };

  return (
    <ScrollView style={styles.container}>
      <Text style={styles.title}>Insert Record into: {listName}</Text>

      {fields.map((field) => (
        <View key={field.name} style={styles.fieldContainer}>
          <Text style={styles.label}>
            {field.label}
            {field.required && <Text style={styles.required}> *</Text>}
          </Text>
          <TextInput
            style={styles.input}
            value={formData[field.name] || ''}
            onChangeText={(value) => handleInputChange(field.name, value)}
            placeholder={`Enter ${field.label.toLowerCase()}`}
            keyboardType={
              field.type === 'number'
                ? 'numeric'
                : field.type === 'email'
                ? 'email-address'
                : 'default'
            }
          />
        </View>
      ))}

      <Text style={styles.sectionTitle}>Custom Fields</Text>
      {customFields.map((customField, index) => (
        <View key={index} style={styles.customFieldContainer}>
          <TextInput
            style={[styles.input, styles.customFieldName]}
            value={customField.name}
            onChangeText={(value) => updateCustomField(index, 'name', value)}
            placeholder="Field name"
          />
          <TextInput
            style={[styles.input, styles.customFieldValue]}
            value={customField.value}
            onChangeText={(value) => updateCustomField(index, 'value', value)}
            placeholder="Field value"
          />
          <TouchableOpacity
            style={styles.removeButton}
            onPress={() => removeCustomField(index)}
          >
            <Text style={styles.removeButtonText}>Remove</Text>
          </TouchableOpacity>
        </View>
      ))}

      <TouchableOpacity
        style={styles.addButton}
        onPress={addCustomField}
      >
        <Text style={styles.addButtonText}>+ Add Custom Field</Text>
      </TouchableOpacity>

      <TouchableOpacity
        style={[styles.submitButton, loading && styles.submitButtonDisabled]}
        onPress={handleSubmit}
        disabled={loading}
      >
        {loading ? (
          <ActivityIndicator color="#fff" />
        ) : (
          <Text style={styles.submitButtonText}>Insert Record</Text>
        )}
      </TouchableOpacity>
    </ScrollView>
  );
};

const styles = StyleSheet.create({
  container: {
    flex: 1,
    padding: 20,
    backgroundColor: '#f5f5f5',
  },
  title: {
    fontSize: 24,
    fontWeight: 'bold',
    marginBottom: 20,
    color: '#333',
  },
  sectionTitle: {
    fontSize: 18,
    fontWeight: '600',
    marginTop: 20,
    marginBottom: 10,
    color: '#333',
  },
  fieldContainer: {
    marginBottom: 15,
  },
  label: {
    fontSize: 16,
    fontWeight: '500',
    marginBottom: 5,
    color: '#333',
  },
  required: {
    color: 'red',
  },
  input: {
    backgroundColor: '#fff',
    borderWidth: 1,
    borderColor: '#ddd',
    borderRadius: 8,
    padding: 12,
    fontSize: 16,
  },
  customFieldContainer: {
    flexDirection: 'row',
    marginBottom: 10,
    gap: 10,
    alignItems: 'center',
  },
  customFieldName: {
    flex: 2,
  },
  customFieldValue: {
    flex: 3,
  },
  removeButton: {
    backgroundColor: '#ff4444',
    paddingHorizontal: 12,
    paddingVertical: 8,
    borderRadius: 6,
  },
  removeButtonText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  addButton: {
    backgroundColor: '#4CAF50',
    padding: 12,
    borderRadius: 8,
    alignItems: 'center',
    marginBottom: 20,
  },
  addButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: '600',
  },
  submitButton: {
    backgroundColor: '#0078d4',
    padding: 16,
    borderRadius: 8,
    alignItems: 'center',
    marginTop: 10,
    marginBottom: 30,
  },
  submitButtonDisabled: {
    opacity: 0.6,
  },
  submitButtonText: {
    color: '#fff',
    fontSize: 18,
    fontWeight: 'bold',
  },
});

export default RecordForm;

