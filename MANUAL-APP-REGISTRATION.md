# Azure AD App Registration Setup Guide

Complete guide for setting up Azure AD app registration with **certificate-based authentication** for SharePoint site template export/import operations.

## Overview

This guide walks through creating an Azure AD app registration with:
- ✅ Certificate-based authentication (recommended - most secure)
- ✅ SharePoint API permissions
- ✅ Admin consent for tenant-wide access
- ✅ Bypasses Conditional Access restrictions

## Prerequisites

- Global Administrator or Application Administrator role in Azure AD
- PowerShell 5.1 or later
- Access to Azure Portal

## Step-by-Step Setup

### Step 1: Access Azure Portal

1. Go to: https://portal.azure.com
2. Sign in with your admin account
3. Navigate to **Azure Active Directory** → **App registrations** → **New registration**

   Direct link: https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps

### Step 2: Register the Application

1. **Name**: `SharePoint Site Template Manager`
2. **Supported account types**: `Accounts in this organizational directory only (Single tenant)`
3. **Redirect URI**: Leave blank
4. Click **Register**

5. **Save these values** from the Overview page:
   - **Application (client) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
   - **Directory (tenant) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`

### Step 3: Configure API Permissions

**CRITICAL**: You must use **SharePoint** API permissions (NOT Microsoft Graph)

1. In your app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **SharePoint** (Office 365 SharePoint Online)
   - ⚠️ Do NOT select "Microsoft Graph"
4. Select **Application permissions** (not Delegated)
5. Expand **Sites** and check:
   - ✅ `Sites.FullControl.All` - Have full control of all site collections

6. Click **Add permissions**
7. Click **Grant admin consent for [Your Organization]**
8. Click **Yes** to confirm

**Verify**: You should see:
- ✅ **SharePoint** - Sites.FullControl.All - Status: **Granted for [Organization]** (green checkmark)

### Step 4: Create Certificate (Recommended)

Certificate-based authentication is the most secure method and bypasses Conditional Access restrictions.

#### 4A: Generate Self-Signed Certificate

Open PowerShell and run:

```powershell
# Generate certificate
$cert = New-SelfSignedCertificate `
    -Subject "CN=SharePointTemplateApp" `
    -CertStoreLocation "Cert:\CurrentUser\My" `
    -KeyExportPolicy Exportable `
    -KeySpec Signature `
    -KeyLength 2048 `
    -KeyAlgorithm RSA `
    -HashAlgorithm SHA256 `
    -NotAfter (Get-Date).AddYears(2)

# Export certificate file
Export-Certificate -Cert $cert -FilePath "C:\Temp\SharePointAppCert.cer"

# Display thumbprint (save this!)
Write-Host "Certificate Thumbprint: $($cert.Thumbprint)" -ForegroundColor Cyan
```

#### 4B: Upload Certificate to Azure AD

1. In Azure Portal, go to your app registration
2. Click **Certificates & secrets** → **Certificates** tab
3. Click **Upload certificate**
4. Upload the file: `C:\Temp\SharePointAppCert.cer`
5. Click **Add**
6. Verify the thumbprint matches what PowerShell displayed

### Step 5: Optional - Create Client Secret (Fallback)

**Note**: Certificate authentication is recommended. Only add a client secret if needed for fallback.

1. Go to **Certificates & secrets** → **Client secrets** tab
2. Click **New client secret**
3. Description: `SharePoint Template Tool`
4. Expires: Choose 24 months
5. Click **Add**
6. **CRITICAL**: Copy the secret **Value** immediately (NOT the Secret ID)
   - The **Value** looks like: `ABC123~xxxxxxxxxxxxxxxxxxxxxxxxxxxx`
   - You cannot view it again after leaving this page!

### Step 6: Create Configuration File

Create a file named `app-config.json` in the script directory:

#### For Certificate Authentication (Recommended):

```json
{
  "tenantId": "YOUR-TENANT-ID-GUID-HERE",
  "clientId": "YOUR-CLIENT-ID-GUID-HERE",
  "certificateThumbprint": "YOUR-CERTIFICATE-THUMBPRINT-HERE",
  "tenantDomain": "yourtenant.onmicrosoft.com"
}
```

#### For Certificate + Client Secret Fallback:

```json
{
  "tenantId": "YOUR-TENANT-ID-GUID-HERE",
  "clientId": "YOUR-CLIENT-ID-GUID-HERE",
  "certificateThumbprint": "YOUR-CERTIFICATE-THUMBPRINT-HERE",
  "clientSecret": "YOUR-CLIENT-SECRET-VALUE-HERE",
  "tenantDomain": "yourtenant.onmicrosoft.com"
}
```

**Values**:
- `tenantId`: Directory (tenant) ID from Step 2
- `clientId`: Application (client) ID from Step 2
- `certificateThumbprint`: Certificate thumbprint from Step 4
- `clientSecret`: Secret value from Step 5 (if using fallback)
- `tenantDomain`: Your tenant domain (e.g., contoso.onmicrosoft.com)

### Step 7: Test the Setup

```powershell
# Test certificate authentication
$config = Get-Content "app-config.json" | ConvertFrom-Json

Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/test" `
    -ClientId $config.clientId `
    -Thumbprint $config.certificateThumbprint `
    -Tenant $config.tenantId

# Test query
Get-PnPWeb | Select-Object Title, Url

# If successful, you'll see your site information!
```

### Step 8: Run Export/Import Scripts

Now you can use the scripts:

```powershell
# Export a site
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://yourtenant.sharepoint.com/sites/Source" `
    -IncludeContent

# Import to target
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://yourtenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\PSReports\SiteTemplates\YourTemplate.pnp"
```

## Authentication Methods Comparison

| Method | Security | Bypasses CA | Automation | Setup Complexity |
|--------|----------|-------------|------------|------------------|
| **Certificate** | ⭐⭐⭐⭐⭐ | ✅ Yes | ✅ Yes | Medium |
| Client Secret | ⭐⭐⭐ | ❌ No | ✅ Yes | Easy |
| Interactive | ⭐⭐ | ❌ No | ❌ No | None |

**Recommendation**: Use certificate authentication for production environments.

## Troubleshooting

### Certificate Not Found

**Symptoms**: `Cannot find certificate with thumbprint...`

**Solution**:
```powershell
# List certificates in your store
Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint

# Verify thumbprint matches app-config.json
```

### Wrong API Permissions

**Symptoms**: "Unauthorized" error even with correct certificate

**Cause**: Using Microsoft Graph permissions instead of SharePoint

**Solution**: 
1. Go to app → API permissions
2. Verify you see **SharePoint** (not Microsoft Graph)
3. If wrong, remove Graph permissions and add SharePoint permissions

### Admin Consent Not Granted

**Symptoms**: "Need admin approval" error

**Solution**:
1. Go to app → API permissions
2. Click "Grant admin consent for [Organization]"
3. Confirm by clicking "Yes"
4. Verify green checkmarks appear

### Conditional Access Blocking

**Symptoms**: Error 53003 when using client secret

**Solution**: Use certificate authentication - it bypasses CA policies

## Security Considerations

### Certificate Security
- ✅ Certificates stored in Windows certificate store
- ✅ Private key marked as exportable (for backup)
- ✅ Certificate valid for 2 years (adjust as needed)
- ⚠️ Rotate certificates before expiration
- ⚠️ Keep certificate thumbprint secure

### Permission Scope
- App has **Sites.FullControl.All** - very powerful
- Only use for migration/template operations
- Consider creating separate apps for different purposes
- Revoke access when operations complete

### Credential Storage
- `app-config.json` contains sensitive data
- Keep secure and don't commit to source control
- Use file encryption if storing long-term
- Consider using Azure Key Vault for production

## Certificate Management

### View Installed Certificates
```powershell
Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint, NotAfter
```

### Export Certificate for Backup
```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq "YOUR-THUMBPRINT"}
$pwd = ConvertTo-SecureString -String "YourPassword" -Force -AsPlainText
Export-PfxCertificate -Cert $cert -FilePath "backup.pfx" -Password $pwd
```

### Import Certificate on Another Machine
```powershell
$pwd = ConvertTo-SecureString -String "YourPassword" -Force -AsPlainText
Import-PfxCertificate -FilePath "backup.pfx" -CertStoreLocation Cert:\CurrentUser\My -Password $pwd
```

### Check Certificate Expiration
```powershell
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq "YOUR-THUMBPRINT"}
$daysUntilExpiration = ($cert.NotAfter - (Get-Date)).Days
Write-Host "Certificate expires in $daysUntilExpiration days" -ForegroundColor $(if($daysUntilExpiration -lt 30){"Red"}else{"Green"})
```

## Quick Reference

### App Registration Checklist

- [ ] App registration created
- [ ] Application (client) ID saved
- [ ] Directory (tenant) ID saved
- [ ] **SharePoint** API permission added (Sites.FullControl.All)
- [ ] Admin consent granted
- [ ] Certificate generated and uploaded
- [ ] Certificate thumbprint saved
- [ ] app-config.json created
- [ ] Test connection successful

### Key URLs

- **Azure Portal**: https://portal.azure.com
- **App Registrations**: https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps
- **Your App**: https://portal.azure.com/#view/Microsoft_AAD_IAM/ApplicationBlade/appId/YOUR-APP-ID

### PowerShell Quick Commands

```powershell
# Test connection
Connect-PnPOnline -Url "https://tenant.sharepoint.com/sites/site" -ClientId "APP-ID" -Thumbprint "THUMBPRINT" -Tenant "TENANT-ID"

# List certificates
Get-ChildItem Cert:\CurrentUser\My

# View config
Get-Content app-config.json | ConvertFrom-Json
```

## Support

For script usage and examples, see [README.md](README.md)

### Step 1: Access Azure Portal

1. Go to: https://portal.azure.com
2. Sign in with your admin account
3. Navigate to **Azure Active Directory** → **App registrations** → **New registration**

   *Direct link:* https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps

### Step 2: Register the Application

1. **Name**: `SharePoint Site Template Manager`
2. **Supported account types**: `Accounts in this organizational directory only (Single tenant)`
3. **Redirect URI**: Leave blank for now
4. Click **Register**

### Step 3: Configure API Permissions

1. In your new app registration, go to **API permissions**
2. Click **Add a permission**
3. Select **Microsoft Graph** → **Application permissions**
4. Add these permissions:
   - `Sites.FullControl.All` (Read and write items in all site collections)
   - `Sites.Read.All` (Read items in all site collections)
   - `Sites.ReadWrite.All` (Read and write items in all site collections)
   - Optional: `User.Read.All` (if you need to read user information)

5. Click **Add permissions**
6. Click **Grant admin consent for [Your Organization]**
   - This requires Global Admin or Privileged Role Admin permissions

### Step 4: Create Authentication Credentials

Choose **ONE** of these options:

#### Option A: Client Secret (Easier, but less secure)

1. Go to **Certificates & secrets**
2. Click **New client secret**
3. Description: `SharePoint Template Tool`
4. Expires: Choose 24 months
5. Click **Add**
6. **CRITICAL**: Copy the secret **Value** column immediately (NOT the Secret ID column)
   - The **Value** looks like: `XyZ8Q~aBcDeFgHiJkLmNoPqRsTuVwXyZ1234567`
   - The **Secret ID** looks like: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` ❌ (DON'T use this)
   - You cannot view the Value again after leaving this page!

#### Option B: Certificate (More secure, recommended)

1. Open PowerShell and run:
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SharePointTemplateApp" -CertStoreLocation "Cert:\CurrentUser\My" -KeyExportPolicy Exportable -KeySpec Signature -KeyLength 2048 -KeyAlgorithm RSA -HashAlgorithm SHA256 -NotAfter (Get-Date).AddYears(2)
   
   Export-Certificate -Cert $cert -FilePath "C:\Temp\SharePointAppCert.cer"
   
   $cert.Thumbprint
   ```

2. In Azure Portal, go to **Certificates & secrets** → **Certificates** tab
3. Click **Upload certificate**
4. Upload the `SharePointAppCert.cer` file
5. Note the **Thumbprint** from PowerShell output

### Step 5: Note Important Values

From the **Overview** page of your app registration, copy:

1. **Application (client) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
2. **Directory (tenant) ID**: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`
3. **Client Secret** (from Step 4A) OR **Certificate Thumbprint** (from Step 4B)

### Step 6: Create Configuration File

Create a file named `app-config.json` in this folder with:

```json
{
  "tenantId": "YOUR-TENANT-ID-GUID-HERE",
  "clientId": "YOUR-CLIENT-ID-GUID-HERE",
  "clientSecret": "YOUR-CLIENT-SECRET-VALUE-HERE",
  "tenantDomain": "yourtenant.onmicrosoft.com"
}
```

### Step 7: Update Your Scripts

The `Export-SharePointSiteTemplate.ps1` and `Import-SharePointSiteTemplate.ps1` scripts should now be updated to use these credentials instead of interactive authentication.

Would you like me to update those scripts to use this app registration?

---

## Why This Happened

**Error 53003** occurs when:
- Your device is not registered/joined to Azure AD
- Conditional Access policies require device compliance
- The Microsoft Graph CLI app is blocked by policy
- Your location/IP address triggers conditional access

## Alternative Solutions

1. **Use Azure Cloud Shell**
   - Go to https://shell.azure.com
   - Cloud Shell is pre-authenticated and bypasses device restrictions
   - Upload the scripts and run them there

2. **Use a Compliant Device**
   - Use a device that's already registered with your organization
   - Device should show as "Compliant" in Azure AD

3. **Request Policy Exception**
   - Ask IT admin to exclude the Microsoft Graph Command Line Tools app from conditional access
   - App ID: `14d82eec-204b-4c2f-b7e8-296a70dab67e`

## Need Help?

Contact your IT administrator with this information:
- Error Code: 53003
- App: Microsoft Graph Command Line Tools
- App ID: 14d82eec-204b-4c2f-b7e8-296a70dab67e
- Request: Need to create app registration for SharePoint automation
