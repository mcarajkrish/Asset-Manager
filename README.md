# Employee & Assets Allocation

A React Native Expo mobile application for managing employee records and asset allocation in SharePoint. Connect to your SharePoint site to view, create, update, and manage employee and asset records.

## Features

- 🔐 OAuth 2.0 authentication with Azure AD (PKCE)
- 📋 Connect to SharePoint sites
- 👥 Manage employee records
- 📦 Manage asset records
- 🎫 Manage access card records
- ➕ Create new records
- ✏️ Update existing records
- 🗑️ Delete records
- 🔍 Browse and select from available lists
- 📱 Beautiful, modern mobile UI

## Quick Start

1. **Install dependencies**: `npm install`
2. **Set up Azure AD**: Follow the [Complete Setup Guide](#complete-setup-guide) below
3. **Configure app**: Update `config/sharepointConfig.ts` with your Client ID and Tenant ID
4. **Validate**: Run `npm run validate-config`
5. **Run**: `npm start` then press `i` for iOS simulator

## Table of Contents

1. [Complete Setup Guide](#complete-setup-guide)
2. [Azure AD App Registration](#azure-ad-app-registration)
3. [Configure API Permissions](#configure-api-permissions)
4. [Configure Redirect URI](#configure-redirect-uri)
5. [Admin Consent](#admin-consent)
6. [Permissions Guide](#permissions-guide)
7. [Troubleshooting](#troubleshooting)
8. [Common Errors](#common-errors)
9. [Usage](#usage)
10. [Project Structure](#project-structure)

---

## Complete Setup Guide

### Step 1: Install Dependencies

```bash
npm install
```

### Step 2: Set Up Azure AD (See detailed steps below)

1. Create app registration
2. Configure API permissions
3. Get Client ID and Tenant ID
4. Configure redirect URI
5. Grant admin consent

### Step 3: Update Configuration

Edit `config/sharepointConfig.ts`:
```typescript
export const SharePointConfig = {
  siteUrl: 'https://humanpoweredhealth.sharepoint.com/sites/hph-BHARAT-asset',
  clientId: 'YOUR_CLIENT_ID_HERE',  // ← Replace with your Client ID
  tenantId: 'YOUR_TENANT_ID_HERE',  // ← Replace with your Tenant ID
  defaultListName: 'Items',
};
```

### Step 4: Validate Configuration

```bash
npm run validate-config
```

### Step 5: Run the App

```bash
npm start
# Press 'i' for iOS simulator
```

---

## Azure AD App Registration

### Step 1: Access Azure Portal

1. Go to [Azure Portal](https://portal.azure.com)
2. Sign in with your Microsoft account

### Step 2: Create App Registration

1. Search for **"Azure Active Directory"** in the top search bar
2. Click on **Azure Active Directory** from the results
3. In the left sidebar, click on **App registrations**
4. Click the **+ New registration** button

### Step 3: Register Application

Fill in the registration form:

1. **Name**: Enter a name (e.g., "HPH-BHARAT-Asset-App")
2. **Supported account types**: 
   - Select **"Accounts in this organizational directory only"** (Single tenant)
3. **Redirect URI**: Leave blank for now (we'll add it later)
4. Click **Register**

### Step 4: Copy Your IDs

After registration, on the **Overview** page:

1. **Copy the Application (client) ID** - This is your Client ID
2. **Copy the Directory (tenant) ID** - This is your Tenant ID

**Save these values!** Both are in GUID format: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

---

## Configure API Permissions

### Step 1: Add Microsoft Graph Permission

1. In your app registration, click on **API permissions** in the left sidebar
2. Click **+ Add a permission**
3. Select **Microsoft Graph**
4. Choose **Delegated permissions**
5. Search for and select:
   - `Sites.ReadWrite.All`
   - `User.Read` (usually already added)
6. Click **Add permissions**

### Step 2: Verify Permissions

You should now see:
- ✅ Microsoft Graph: `User.Read` (Delegated)
- ✅ Microsoft Graph: `Sites.ReadWrite.All` (Delegated)

**Note**: The app uses Microsoft Graph API, so you don't need SharePoint Online API permissions.

---

## Configure Redirect URI

### Permanent Solution (Recommended)

The app is configured to use a **fixed custom scheme** that doesn't change with your IP address:

**Fixed Redirect URI: `employee-assets://auth`**

This URI will work consistently and never change, regardless of your network or IP address.

#### Step 1: Add Fixed URI to Azure AD

1. In Azure Portal, go to your app registration → **Authentication**
2. Under **Platform configurations**, click **+ Add a platform**
3. Select **"Mobile and desktop applications"**
4. Click **Configure**
5. In the **Redirect URIs** field, add: `employee-assets://auth`
6. Under **"Implicit grant and hybrid flows"**, check:
   - ✅ **Access tokens** (ID tokens)
   - ✅ **ID tokens**
7. Click **Save**

#### Step 2: Build Development Build (Required for Custom Scheme)

**Note**: Custom URL schemes like `employee-assets://auth` work in development builds and production builds, but **not in Expo Go**.

To use the permanent fixed URI, create a development build:

```bash
# For iOS
npx expo run:ios

# For Android
npx expo run:android
```

This will create a development build that supports the custom scheme.

### Temporary Solution (Expo Go Only)

If you're using **Expo Go** (not a development build), the app will fall back to an IP-based URI that changes with your network:

1. Open your app in Expo Go
2. Look at the connection screen for the **"Redirect URI:"**
3. It will show something like:
   - `exp://localhost:8081` (iOS Simulator - more stable)
   - `exp://192.168.x.x:8081` (physical device - changes with IP)

**To minimize changes**, use `exp://localhost:8081` when testing in the iOS Simulator, as it's more stable than IP addresses.

#### Add Multiple URIs to Azure AD

You can add multiple redirect URIs to Azure AD:
- `employee-assets://auth` (permanent, works with development/production builds)
- `exp://localhost:8081` (for Expo Go in simulator)
- `exp://192.168.x.x:8081` (for Expo Go on physical device - may need to update when IP changes)

### Important Notes

- **For permanent fix**: Use `employee-assets://auth` with a development build
- **For Expo Go**: The URI may change when your IP changes (add multiple URIs or use localhost)
- The redirect URI must match **exactly** (protocol, host, port, path)
- You can add multiple URIs in Azure AD if needed

---

## Admin Consent

### Why Admin Consent is Required

`Sites.ReadWrite.All` is a high-privilege permission that requires admin consent to protect organizational data.

### Option 1: Grant Admin Consent in Azure Portal (Recommended)

1. Go to Azure Portal → Your App → **API permissions**
2. Click **"Grant admin consent for [Your Organization]"** button
3. Click **Yes** to confirm
4. **Wait for the page to refresh**
5. Verify the **Status** column shows:
   - ✅ **"Granted for [Your Organization]"** with green checkmarks for all permissions

### Option 2: Admin Signs In During Authentication

If you see the "Need admin approval" screen:

1. Click **"Have an admin account? Sign in with that account"**
2. An admin signs in with their admin account
3. They grant consent when prompted
4. After this, the app will work for all users

### If You Don't Have Admin Access

Ask your Azure AD administrator to grant consent. Send them:

**Subject: Admin Consent Required for SharePoint App**

Hi,

I need admin consent for the Azure AD app registration "[Your App Name]" to access SharePoint.

**Required Action:**
1. Go to Azure Portal → Azure AD → App registrations → [Your App Name]
2. Click "API permissions"
3. Click "Grant admin consent for [Organization]"
4. Confirm

**Permissions Needed:**
- Microsoft Graph: Sites.ReadWrite.All (Delegated)
- Microsoft Graph: User.Read (Delegated)

This is a one-time action that will allow all users in the organization to use the app.

Thank you!

---

## Permissions Guide

### Required Azure AD Permissions

The app requires these **Delegated Permissions**:

#### Microsoft Graph API

| Permission | Type | Description | Admin Consent |
|------------|------|-------------|---------------|
| `Sites.ReadWrite.All` | Delegated | Read and write items in all site collections | ✅ **Required** |
| `User.Read` | Delegated | Sign in and read user profile | ❌ No |

### Authentication Scopes

The app requests these scopes during OAuth flow:

```
https://graph.microsoft.com/Sites.ReadWrite.All
https://graph.microsoft.com/User.Read
```

### SharePoint Site Permissions

**Site**: `https://humanpoweredhealth.sharepoint.com/sites/hph-BHARAT-asset`

User account needs:
- **Minimum**: Edit permission level
- **Recommended**: Contribute permission level

Includes:
- ✅ View Items
- ✅ Add Items
- ✅ Edit Items
- ✅ Delete Items
- ✅ Open Items

### SharePoint List Permissions

**Lists**: Assets and Access Cards

User account needs on both lists:
- ✅ View Items (read)
- ✅ Add Items (write)
- ✅ Edit Items (write)
- ✅ Delete Items (optional)
- ✅ Open Items

**Note**: When using the API, use:
- Assets list: `'Assets'`
- Access Cards list: `'Access Cards'` (includes the space)

### Permission Summary Checklist

#### ✅ Azure AD Configuration
- [ ] `Sites.ReadWrite.All` added to Microsoft Graph (Delegated)
- [ ] `User.Read` added to Microsoft Graph (Delegated)
- [ ] **Admin consent granted** (all show green checkmark ✅)

#### ✅ SharePoint Access
- [ ] User can access the site in browser
- [ ] User can view the **Assets** list
- [ ] User can create/edit items in the **Assets** list manually
- [ ] User can view the **Access Cards** list
- [ ] User can create/edit items in the **Access Cards** list manually

#### ✅ Authentication
- [ ] App requests `Sites.ReadWrite.All` scope
- [ ] App requests `User.Read` scope
- [ ] User authenticates successfully

---

## Troubleshooting

### Step-by-Step Verification Checklist

1. ✅ **Configuration Valid**
   ```bash
   npm run validate-config
   ```

2. ✅ **Redirect URI Added to Azure AD**
   - Check the URI shown in your app
   - Add it to Azure Portal → Authentication → Redirect URIs
   - Must match exactly

3. ✅ **API Permissions Configured**
   - Microsoft Graph: `Sites.ReadWrite.All`
   - Microsoft Graph: `User.Read`
   - Admin consent granted (Status shows "Granted for [Organization]")

4. ✅ **Correct Account**
   - Using account with SharePoint access
   - Account in same tenant as app registration

5. ✅ **App Registration Active**
   - App exists in Azure Portal
   - Not disabled or deleted

### Common Issues

#### "Need admin approval" Screen

**Problem**: Admin consent hasn't been granted.

**Solution**:
- Grant admin consent in Azure Portal (see [Admin Consent](#admin-consent) section)
- Or have an admin sign in during authentication

#### "Authentication cancelled by user"

**Problem**: Usually redirect URI mismatch or browser closed.

**Solution**:
- Verify redirect URI matches exactly in Azure AD
- Don't close the browser during authentication
- Complete the sign-in process

#### Missing Permissions

**Problem**: Only `User.Read` is configured.

**Solution**:
- Add `Sites.ReadWrite.All` for Microsoft Graph
- See [Configure API Permissions](#configure-api-permissions) section

#### Redirect URI Mismatch

**Problem**: URI in Azure AD doesn't match app.

**Solution**:
- Check exact URI shown in your app
- Compare character-by-character with Azure AD
- Update Azure AD to match exactly

#### "Access Denied" or "Forbidden" Errors

**Possible Causes:**
1. Admin consent not granted
2. User doesn't have site permissions
3. List has unique permissions restricting access
4. User account is not in the correct tenant

**Solutions:**
1. Check admin consent status in Azure Portal
2. Verify user has Edit/Contribute permission on the site
3. Check list permissions (Settings → List settings → Permissions)
4. Verify user account tenant matches app registration tenant

---

## Common Errors

### "invalid_client"
- **Cause**: Client ID is incorrect or app registration doesn't exist
- **Fix**: Verify Client ID in `config/sharepointConfig.ts` matches Azure Portal

### "invalid_tenant"
- **Cause**: Tenant ID is incorrect
- **Fix**: Verify Tenant ID in `config/sharepointConfig.ts` matches Azure Portal

### "invalid_grant"
- **Cause**: Redirect URI mismatch or code already used/expired
- **Fix**: 
  - Verify redirect URI matches exactly
  - Try connecting again (codes expire quickly)

### "AADSTS50020"
- **Cause**: User account not found in tenant
- **Fix**: Use an account that exists in the Azure AD tenant

### "AADSTS70016"
- **Cause**: Application not found in the tenant
- **Fix**: Verify the app registration exists in the correct tenant

### "AADSTS70011" / "scope is not valid"
- **Cause**: Invalid scope combination
- **Fix**: Use only Microsoft Graph scopes (already configured in the app)

### "AADSTS90014" / "required field 'request' is missing"
- **Cause**: PKCE code verifier issue
- **Fix**: Usually resolves automatically. If persists, clear app cache and retry

### "insufficient_privileges" or "Access Denied"
- **Cause**: Permissions not granted or not consented
- **Fix**: Grant admin consent in Azure Portal → API permissions

---

## Usage

After successful setup:

1. **Connect to SharePoint**: Tap "Connect to SharePoint" and authenticate
2. **Select List**: 
   - Use quick access buttons for **Assets** or **Access Cards**
   - Or enter a list name manually and use "Show Lists" to browse
3. **View Records**: Records are displayed automatically in "View Records" mode
4. **Create Record**: Switch to "Create Record" mode, fill in the form, and tap "Insert Record"
5. **Update Record**: Tap a record to view details, then tap "✏️ Edit Record"
6. **Delete Record**: Tap the 🗑️ icon on a record card or in the record view

### Supported Lists

The app supports these SharePoint lists:

1. **Assets** (`Assets`)
   - Quick access button available
   - Full CRUD operations

2. **Access Cards** (`Access Cards`)
   - Quick access button available
   - Full CRUD operations

3. **Any Other List**
   - Manual selection via list browser
   - Full CRUD operations

---

## Project Structure

```
├── App.tsx                    # Main app component with navigation
├── components/
│   ├── RecordForm.tsx         # Form component for creating records
│   ├── RecordsList.tsx       # Component for displaying list of records
│   └── RecordView.tsx        # Component for viewing/editing individual records
├── screens/
│   ├── LoginScreen.tsx       # Authentication screen
│   ├── HomeScreen.tsx        # Main home screen with quick access
│   ├── ListScreen.tsx        # Screen for viewing and managing list records
│   ├── DetailScreen.tsx      # Screen for viewing record details
│   └── CreateRecordScreen.tsx # Screen for creating new records
├── services/
│   └── sharepointService.ts  # SharePoint API service (Microsoft Graph API)
├── config/
│   └── sharepointConfig.ts   # Configuration file
├── scripts/
│   └── validateConfig.ts     # Configuration validation script
└── README.md                 # This comprehensive guide
```

---

## Technical Details

### Authentication

- **OAuth 2.0 with PKCE**: Industry-standard secure authentication
- **Token-Based**: Short-lived access tokens (stored in memory)
- **User Context**: All actions are performed as the authenticated user
- **Delegated Permissions**: App acts with user's permissions, not app-only

### API

- **Microsoft Graph API**: Modern API for accessing SharePoint
- **Endpoints**: Uses Graph API endpoints like `https://graph.microsoft.com/v1.0/sites/{siteId}/lists`
- **Data Format**: Standard JSON (no OData verbose format)

### Security

- ✅ **OAuth 2.0 with PKCE**: Secure authentication
- ✅ **Token-Based**: No stored credentials
- ✅ **User Context**: All actions tied to authenticated user
- ✅ **Delegated Permissions**: No elevation of privileges
- ✅ **Scope Limited**: Users can only access data they have permission to see

---

## Additional Notes

- The app uses OAuth 2.0 with PKCE for secure authentication
- Tokens are stored in memory (consider implementing token persistence for production)
- The form includes a default "Title" field which is required in most SharePoint lists
- You can add custom fields dynamically in the form
- Admin consent is organization-wide and only needs to be done once
- The app uses Microsoft Graph API (not SharePoint REST API)

---

## Getting Help

If you're still having issues:

1. **Check Azure AD sign-in logs**: Azure Portal → Azure AD → Sign-in logs
2. **Verify all steps**: Go through the verification checklist above
3. **Check console logs**: Look for specific error messages in your app logs
4. **Try a different account**: Rule out account-specific issues
5. **Run validation**: `npm run validate-config`

---

## License

This project is private and proprietary.

---

**Last Updated**: Current  
**App Version**: Using Microsoft Graph API  
**Status**: Production Ready
