import React, { useState, useEffect } from 'react';
import {
  View,
  Text,
  TouchableOpacity,
  StyleSheet,
  Alert,
  ActivityIndicator,
  ScrollView,
} from 'react-native';
import { SafeAreaView } from 'react-native-safe-area-context';
import * as AuthSession from 'expo-auth-session';
import SharePointService from '../services/sharepointService';

interface LoginScreenProps {
  sharePointService: SharePointService;
  onLoginSuccess: () => void;
}

const LoginScreen: React.FC<LoginScreenProps> = ({ sharePointService, onLoginSuccess }) => {
  const [isConnecting, setIsConnecting] = useState(false);
  const [redirectUri, setRedirectUri] = useState<string>('');

  useEffect(() => {
    // Use a fixed custom scheme URI that doesn't depend on IP address
    // Note: Custom schemes work in development/production builds, but Expo Go will fall back to exp://
    let uri = AuthSession.makeRedirectUri({
      scheme: 'employee-assets',
      path: 'auth',
    });
    
    // Fallback: If still using exp:// scheme (Expo Go limitation), use localhost which is more stable
    if (uri.startsWith('exp://')) {
      // Try to use localhost instead of IP for stability
      const localhostUri = AuthSession.makeRedirectUri({
        preferLocalhost: true,
      });
      if (localhostUri && localhostUri.includes('localhost')) {
        uri = localhostUri.replace(/\/--\/.*$/, '/'); // Remove path, keep base URI
      } else {
        // Last resort: use fixed custom scheme format
        uri = 'employee-assets://auth';
      }
    }
    
    // Ensure the redirect URI is in the correct format
    if (!uri.includes('://')) {
      uri = 'employee-assets://auth';
    }
    
    setRedirectUri(uri);
  }, []);

  const handleConnect = async () => {
    try {
      setIsConnecting(true);
      await sharePointService.authenticate();
      onLoginSuccess();
    } catch (error: any) {
      const errorMessage = error.message || 'Failed to connect to SharePoint';
      Alert.alert('Connection Error', errorMessage);
    } finally {
      setIsConnecting(false);
    }
  };

  return (
    <SafeAreaView style={styles.container} edges={['top', 'bottom']}>
      <ScrollView contentContainerStyle={styles.scrollContent}>
        <Text style={styles.title}>Employee & Assets Allocation</Text>
        <Text style={styles.subtitle}>
          Manage employee records and assets allocation in SharePoint
        </Text>

        {redirectUri && (
          <View style={styles.redirectUriContainer}>
            <Text style={styles.redirectUriLabel}>Redirect URI:</Text>
            <Text style={styles.redirectUriText} selectable={true}>
              {redirectUri}
            </Text>
            <Text style={styles.redirectUriNote}>
              ✅ This URI is fixed and will not change with your IP address{'\n'}
              ⚠️ Make sure this exact URI is added to Azure AD → Authentication → Redirect URIs
            </Text>
          </View>
        )}

        <TouchableOpacity
          style={[styles.connectButton, isConnecting && styles.buttonDisabled]}
          onPress={handleConnect}
          disabled={isConnecting}
        >
          {isConnecting ? (
            <ActivityIndicator color="#fff" />
          ) : (
            <Text style={styles.connectButtonText}>Connect to SharePoint</Text>
          )}
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
  scrollContent: {
    flexGrow: 1,
    justifyContent: 'center',
    alignItems: 'center',
    padding: 20,
  },
  title: {
    fontSize: 28,
    fontWeight: 'bold',
    marginBottom: 10,
    color: '#333',
    textAlign: 'center',
  },
  subtitle: {
    fontSize: 16,
    color: '#666',
    marginBottom: 30,
    textAlign: 'center',
  },
  adminNoteContainer: {
    backgroundColor: '#fff3cd',
    borderLeftWidth: 4,
    borderLeftColor: '#ffc107',
    padding: 15,
    borderRadius: 8,
    marginBottom: 20,
    width: '100%',
    maxWidth: 400,
  },
  adminNoteTitle: {
    fontSize: 16,
    fontWeight: 'bold',
    color: '#856404',
    marginBottom: 8,
  },
  adminNoteText: {
    fontSize: 13,
    color: '#856404',
    lineHeight: 20,
  },
  redirectUriContainer: {
    backgroundColor: '#f0f0f0',
    padding: 15,
    borderRadius: 8,
    marginBottom: 20,
    width: '100%',
    maxWidth: 400,
  },
  redirectUriLabel: {
    fontSize: 14,
    fontWeight: '600',
    marginBottom: 5,
    color: '#333',
  },
  redirectUriText: {
    fontSize: 12,
    fontFamily: 'monospace',
    color: '#0078d4',
    marginBottom: 10,
    backgroundColor: '#fff',
    padding: 8,
    borderRadius: 4,
  },
  redirectUriNote: {
    fontSize: 11,
    color: '#d83b01',
    fontStyle: 'italic',
  },
  connectButton: {
    backgroundColor: '#0078d4',
    paddingHorizontal: 30,
    paddingVertical: 15,
    borderRadius: 8,
    minWidth: 200,
    alignItems: 'center',
  },
  buttonDisabled: {
    opacity: 0.6,
  },
  connectButtonText: {
    color: '#fff',
    fontSize: 18,
    fontWeight: 'bold',
  },
});

export default LoginScreen;
