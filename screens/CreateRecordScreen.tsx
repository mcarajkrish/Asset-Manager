import React from 'react';
import {
  View,
  Text,
  TouchableOpacity,
  StyleSheet,
} from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import RecordForm from '../components/RecordForm';
import SharePointService from '../services/sharepointService';

interface CreateRecordScreenProps {
  sharePointService: SharePointService;
  listName: string;
  onBack: () => void;
  onRecordCreated: () => void;
}

const CreateRecordScreen: React.FC<CreateRecordScreenProps> = ({
  sharePointService,
  listName,
  onBack,
  onRecordCreated,
}) => {
  return (
    <SafeAreaView style={styles.container} edges={['top', 'bottom']}>
      <View style={styles.header}>
        <TouchableOpacity onPress={onBack} style={styles.backButton}>
          <Text style={styles.backButtonText}>←</Text>
        </TouchableOpacity>
        <Text style={styles.headerTitle}>Create Record</Text>
        <View style={styles.placeholder} />
      </View>
      <RecordForm
        sharePointService={sharePointService}
        listName={listName}
        onRecordInserted={() => {
          onRecordCreated();
        }}
      />
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
});

export default CreateRecordScreen;
