import React, { useState, useEffect } from 'react';
import {
  View,
  Text,
  TouchableOpacity,
  StyleSheet,
  ScrollView,
} from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import SharePointService, { SessionTimeoutError } from '../services/sharepointService';

interface HomeScreenProps {
  sharePointService: SharePointService;
  currentUser: any;
  isAdmin: boolean;
  employees: any[];
  setEmployees: (employees: any[]) => void;
  onListPress: (listName: string) => void;
  onLogout: () => void;
  navigation?: any;
}

const HomeScreen: React.FC<HomeScreenProps> = ({
  sharePointService,
  currentUser,
  isAdmin,
  employees,
  setEmployees,
  onListPress,
  onLogout,
  navigation,
}) => {
  const [loadingEmployees, setLoadingEmployees] = useState(false);

  useEffect(() => {
    // Fetch employees on mount if not already loaded
    if (employees.length === 0) {
      fetchEmployees();
    }
  }, []);

  const fetchEmployees = async () => {
    try {
      setLoadingEmployees(true);
      // Fetch all users from the organization using Microsoft Graph API
      const allUsers = await sharePointService.getAllUsers();
      setEmployees(allUsers);
    } catch (error: any) {
      // Handle session timeout - don't show alert as App.tsx will handle it
      if (error instanceof SessionTimeoutError) {
        return;
      }
      console.error('Error fetching employees:', error);
    } finally {
      setLoadingEmployees(false);
    }
  };

  return (
    <SafeAreaView style={styles.container} edges={['top', 'bottom']}>
      <ScrollView style={styles.scrollView} contentContainerStyle={styles.scrollContent}>
        {/* Header */}
        <View style={styles.header}>
          <Text style={styles.title}>Asset Management</Text>
          <View style={styles.connectedBadge}>
            <View style={styles.connectedDot} />
            <Text style={styles.connectedText}>Connected</Text>
          </View>
        </View>

        {/* User Info Section */}
        <View style={styles.userInfoContainer}>
          <Text style={styles.sectionTitle}>Logged-In User</Text>
          {currentUser && (
            <View style={styles.userCard}>
              {isAdmin && (
                <View style={styles.adminBadge}>
                  <Text style={styles.adminBadgeText}>ADMIN</Text>
                </View>
              )}
              <Text style={styles.userLabel}>Name:</Text>
              <Text style={styles.userValue}>{currentUser.displayName}</Text>
              
              <Text style={styles.userLabel}>Email:</Text>
              <Text style={styles.userValue}>
                {currentUser.mail || currentUser.userPrincipalName}
              </Text>
              
              {currentUser.jobTitle && (
                <>
                  <Text style={styles.userLabel}>Title:</Text>
                  <Text style={styles.userValue}>{currentUser.jobTitle}</Text>
                </>
              )}
            </View>
          )}
        </View>

        {/* Employees Section */}
        <View style={styles.employeesContainer}>
          <View style={styles.sectionHeader}>
            <Text style={styles.sectionTitle}>Employees</Text>
            <View style={styles.employeeActions}>
              <TouchableOpacity
                onPress={fetchEmployees}
                style={styles.refreshButton}
                disabled={loadingEmployees}
              >
                <Text style={styles.refreshButtonText}>
                  {loadingEmployees ? 'Loading...' : '🔄'}
                </Text>
              </TouchableOpacity>
              <TouchableOpacity
                onPress={() => navigation?.navigate('List', { listName: 'Employees' })}
                style={styles.viewButton}
              >
                <Text style={styles.viewButtonText}>View All</Text>
              </TouchableOpacity>
            </View>
          </View>
          <Text style={styles.employeeCount}>
            {loadingEmployees ? 'Loading...' : `${employees.length} employees loaded`}
          </Text>
        </View>

        {/* Quick Access Section */}
        <View style={styles.quickAccessContainer}>
          <Text style={styles.sectionTitle}>Quick Access</Text>
          <Text style={styles.sectionSubtitle}>Select a list to view and manage records</Text>
          
          <View style={styles.quickAccessButtons}>
            <TouchableOpacity
              style={styles.quickAccessButton}
              onPress={() => onListPress('Assets')}
            >
              <Text style={styles.quickAccessIcon}>📦</Text>
              <Text style={styles.quickAccessText}>Assets</Text>
            </TouchableOpacity>
            
            <TouchableOpacity
              style={styles.quickAccessButton}
              onPress={() => onListPress('Access Cards')}
            >
              <Text style={styles.quickAccessIcon}>🎫</Text>
              <Text style={styles.quickAccessText}>Access Cards</Text>
            </TouchableOpacity>
          </View>
        </View>

        {/* Logout Button */}
        <TouchableOpacity
          style={styles.logoutButton}
          onPress={onLogout}
        >
          <Text style={styles.logoutButtonText}>Logout</Text>
        </TouchableOpacity>
      </ScrollView>
    </SafeAreaView>
  );
};

const styles = StyleSheet.create({
  container: {
    flex: 1,
    backgroundColor: '#f5f5f5',
  },
  scrollView: {
    flex: 1,
  },
  scrollContent: {
    paddingBottom: 30,
  },
  header: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: 20,
    backgroundColor: '#fff',
    borderBottomWidth: 1,
    borderBottomColor: '#e0e0e0',
  },
  title: {
    fontSize: 24,
    fontWeight: 'bold',
    color: '#333',
    flex: 1,
  },
  connectedBadge: {
    flexDirection: 'row',
    alignItems: 'center',
    backgroundColor: '#4CAF50',
    paddingHorizontal: 12,
    paddingVertical: 6,
    borderRadius: 20,
  },
  connectedDot: {
    width: 8,
    height: 8,
    borderRadius: 4,
    backgroundColor: '#fff',
    marginRight: 6,
  },
  connectedText: {
    color: '#fff',
    fontSize: 12,
    fontWeight: '600',
  },
  userInfoContainer: {
    padding: 20,
    backgroundColor: '#fff',
    marginTop: 10,
    marginBottom: 10,
  },
  sectionTitle: {
    fontSize: 20,
    fontWeight: 'bold',
    color: '#333',
    marginBottom: 15,
  },
  sectionSubtitle: {
    fontSize: 14,
    color: '#666',
    marginBottom: 15,
  },
  userCard: {
    backgroundColor: '#f9f9f9',
    padding: 15,
    borderRadius: 8,
    borderWidth: 1,
    borderColor: '#e0e0e0',
  },
  adminBadge: {
    backgroundColor: '#ff6b6b',
    paddingHorizontal: 10,
    paddingVertical: 4,
    borderRadius: 12,
    alignSelf: 'flex-start',
    marginBottom: 10,
  },
  adminBadgeText: {
    color: '#fff',
    fontSize: 11,
    fontWeight: 'bold',
  },
  userLabel: {
    fontSize: 12,
    fontWeight: '600',
    color: '#666',
    marginTop: 8,
    marginBottom: 4,
  },
  userValue: {
    fontSize: 16,
    color: '#333',
    marginBottom: 4,
  },
  quickAccessContainer: {
    padding: 20,
    backgroundColor: '#fff',
    marginBottom: 10,
  },
  quickAccessButtons: {
    flexDirection: 'row',
    gap: 15,
    marginTop: 10,
  },
  quickAccessButton: {
    flex: 1,
    backgroundColor: '#0078d4',
    padding: 20,
    borderRadius: 12,
    alignItems: 'center',
    justifyContent: 'center',
    minHeight: 120,
    shadowColor: '#000',
    shadowOffset: { width: 0, height: 2 },
    shadowOpacity: 0.1,
    shadowRadius: 4,
    elevation: 3,
  },
  quickAccessIcon: {
    fontSize: 48,
    marginBottom: 10,
  },
  quickAccessText: {
    color: '#fff',
    fontSize: 18,
    fontWeight: '600',
  },
  logoutButton: {
    backgroundColor: '#f44336',
    marginHorizontal: 20,
    marginTop: 20,
    padding: 15,
    borderRadius: 8,
    alignItems: 'center',
  },
  logoutButtonText: {
    color: '#fff',
    fontSize: 16,
    fontWeight: '600',
  },
  employeesContainer: {
    padding: 20,
    backgroundColor: '#fff',
    marginBottom: 10,
  },
  sectionHeader: {
    flexDirection: 'row',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: 10,
  },
  employeeActions: {
    flexDirection: 'row',
    gap: 10,
    alignItems: 'center',
  },
  refreshButton: {
    paddingHorizontal: 12,
    paddingVertical: 6,
    backgroundColor: '#0078d4',
    borderRadius: 6,
  },
  refreshButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  viewButton: {
    paddingHorizontal: 16,
    paddingVertical: 6,
    backgroundColor: '#4CAF50',
    borderRadius: 6,
  },
  viewButtonText: {
    color: '#fff',
    fontSize: 14,
    fontWeight: '600',
  },
  employeeCount: {
    fontSize: 14,
    color: '#666',
  },
});

export default HomeScreen;
