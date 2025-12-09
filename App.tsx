import React, { useState, useEffect } from 'react';
import { StatusBar } from 'expo-status-bar';
import { Alert } from 'react-native';
import { SafeAreaProvider } from 'react-native-safe-area-context';
import { NavigationContainer } from '@react-navigation/native';
import { createNativeStackNavigator } from '@react-navigation/native-stack';
import SharePointService, { SessionTimeoutError } from './services/sharepointService';
import LoginScreen from './screens/LoginScreen';
import HomeScreen from './screens/HomeScreen';
import ListScreen from './screens/ListScreen';
import DetailScreen from './screens/DetailScreen';
import CreateRecordScreen from './screens/CreateRecordScreen';
import { SharePointConfig } from './config/sharepointConfig';

export type RootStackParamList = {
  Login: undefined;
  Home: undefined;
  List: { listName: string };
  Detail: { listName: string; record: any };
  CreateRecord: { listName: string };
};

const Stack = createNativeStackNavigator<RootStackParamList>();

export default function App() {
  const [sharePointService] = useState(
    () =>
      new SharePointService({
        siteUrl: SharePointConfig.siteUrl,
        clientId: SharePointConfig.clientId,
        tenantId: SharePointConfig.tenantId,
      })
  );
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [currentUser, setCurrentUser] = useState<any>(null);
  const [isAdmin, setIsAdmin] = useState(false);
  const [loadingUserInfo, setLoadingUserInfo] = useState(false);
  const [employees, setEmployees] = useState<any[]>([]);

  useEffect(() => {
    // Set up session timeout handler
    sharePointService.setOnSessionTimeout(() => {
      handleSessionTimeout();
    });

    // Check if already authenticated
    const token = sharePointService.getAccessToken();
    if (token) {
      setIsAuthenticated(true);
      loadUserInfo();
    }

    // Cleanup on unmount
    return () => {
      sharePointService.clearOnSessionTimeout();
    };
  }, []);

  const loadUserInfo = async () => {
    try {
      setLoadingUserInfo(true);
      const userInfo = await sharePointService.getCurrentUserWithAdminStatus();
      setCurrentUser(userInfo.user);
      setIsAdmin(userInfo.isAdmin);
    } catch (error: any) {
      console.error('Error loading user info:', error);
    } finally {
      setLoadingUserInfo(false);
    }
  };

  const handleLoginSuccess = async () => {
    setIsAuthenticated(true);
    await loadUserInfo();
  };

  const handleSessionTimeout = () => {
    // Clear authentication state
    sharePointService.setAccessToken('');
    setIsAuthenticated(false);
    setCurrentUser(null);
    setIsAdmin(false);
    setEmployees([]);
    
    // Show alert to user
    Alert.alert(
      'Session Expired',
      'Your session has expired. Please log in again.',
      [{ text: 'OK' }]
    );
  };

  const handleLogout = () => {
    Alert.alert(
      'Logout',
      'Are you sure you want to logout?',
      [
        { text: 'Cancel', style: 'cancel' },
        {
          text: 'Logout',
          style: 'destructive',
          onPress: () => {
            sharePointService.setAccessToken('');
            setIsAuthenticated(false);
            setCurrentUser(null);
            setIsAdmin(false);
            setEmployees([]);
          },
        },
      ]
    );
  };

  return (
    <SafeAreaProvider>
      <NavigationContainer>
        <StatusBar style="auto" />
      <Stack.Navigator
        screenOptions={{
          headerShown: false,
        }}
      >
        {!isAuthenticated ? (
          <Stack.Screen name="Login">
            {(props) => (
              <LoginScreen
                {...props}
                sharePointService={sharePointService}
                onLoginSuccess={handleLoginSuccess}
              />
            )}
          </Stack.Screen>
        ) : (
          <>
            <Stack.Screen name="Home">
              {(props) => (
                <HomeScreen
                  {...props}
                  sharePointService={sharePointService}
                  currentUser={currentUser}
                  isAdmin={isAdmin}
                  employees={employees}
                  setEmployees={setEmployees}
                  onListPress={(listName) => {
                    props.navigation.navigate('List', { listName });
                  }}
                  onLogout={handleLogout}
                />
              )}
            </Stack.Screen>
            <Stack.Screen name="List">
              {(props) => {
                const isEmployeesList = props.route.params.listName === 'Employees';
                return (
                  <ListScreen
                    {...props}
                    sharePointService={sharePointService}
                    listName={props.route.params.listName}
                    employees={employees}
                    onRefreshEmployees={isEmployeesList ? async () => {
                      try {
                        // Fetch all users from the organization
                        const allUsers = await sharePointService.getAllUsers();
                        
                        // Transform users to match the expected employee format
                        const employeesList = allUsers.map((user: any) => ({
                          Id: user.id,
                          EmpID: user.userPrincipalName.split('@')[0] || user.id,
                          Employee: user.displayName || user.userPrincipalName,
                          EmployeeName: user.displayName || user.userPrincipalName,
                          Title: user.displayName || user.userPrincipalName,
                          Email: user.mail || user.userPrincipalName,
                          Mail: user.mail || user.userPrincipalName,
                          JobTitle: user.jobTitle,
                          Department: user.department,
                          OfficeLocation: user.officeLocation,
                          UserPrincipalName: user.userPrincipalName,
                        }));
                        
                        setEmployees(employeesList);
                      } catch (error) {
                        console.error('Error refreshing employees:', error);
                      }
                    } : undefined}
                    onRecordPress={(record) => {
                      props.navigation.navigate('Detail', {
                        listName: props.route.params.listName,
                        record,
                      });
                    }}
                    onCreatePress={!isEmployeesList ? () => {
                      props.navigation.navigate('CreateRecord', {
                        listName: props.route.params.listName,
                      });
                    } : undefined}
                    onBack={() => props.navigation.goBack()}
                  />
                );
              }}
            </Stack.Screen>
            <Stack.Screen name="CreateRecord">
              {(props) => (
                <CreateRecordScreen
                  {...props}
                  sharePointService={sharePointService}
                  listName={props.route.params.listName}
                  onBack={() => props.navigation.goBack()}
                  onRecordCreated={() => {
                    props.navigation.goBack();
                  }}
                />
              )}
            </Stack.Screen>
            <Stack.Screen name="Detail">
              {(props) => (
                <DetailScreen
                  {...props}
                  sharePointService={sharePointService}
                  listName={props.route.params.listName}
                  record={props.route.params.record}
                  onBack={() => props.navigation.goBack()}
                  onRecordUpdated={() => {
                    // Navigate back to list to refresh
                    props.navigation.goBack();
                  }}
                />
              )}
            </Stack.Screen>
          </>
        )}
      </Stack.Navigator>
      </NavigationContainer>
    </SafeAreaProvider>
  );
}
