# SharePoint Site Template Export/Import Scripts

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![PnP PowerShell](https://img.shields.io/badge/PnP.PowerShell-2.12%2B-green.svg)](https://pnp.github.io/powershell/)

PowerShell scripts for exporting and importing complete SharePoint sites using PnP templates with **certificate-based authentication**.

## Overview

These scripts enable you to:
- **Export** a complete SharePoint site including structure, content, pages, and lists
- **Import** the template to another SharePoint site (same or different target tenant)
- Use **secure certificate-based authentication** (no interactive login required)
- Maintain full audit trail of operations

## Prerequisites

1. **PowerShell Modules**:
   ```powershell
   Install-Module PnP.PowerShell -Scope CurrentUser
   ```

2. **Azure AD App Registration** with:
   - Certificate-based authentication
   - SharePoint API permissions: `Sites.FullControl.All` (Application permission)
   - Admin consent granted

3. **Permissions**:
   - Site Collection Administrator on source and target sites
   - Global Administrator or Application Administrator (for app registration setup)

## Quick Start

### Step 1: Set Up App Registration

See [MANUAL-APP-REGISTRATION.md](MANUAL-APP-REGISTRATION.md) for detailed setup instructions.

**Summary**:
1. Create Azure AD app registration
2. Add **SharePoint** API permission: `Sites.FullControl.All`
3. Generate certificate:
   ```powershell
   $cert = New-SelfSignedCertificate -Subject "CN=SharePointTemplateApp" `
       -CertStoreLocation "Cert:\CurrentUser\My" `
       -KeyExportPolicy Exportable `
       -KeySpec Signature `
       -KeyLength 2048 `
       -NotAfter (Get-Date).AddYears(2)
   
   Export-Certificate -Cert $cert -FilePath "SharePointAppCert.cer"
   $cert.Thumbprint  # Save this
   ```
4. Upload certificate to Azure AD app registration
5. Grant admin consent

### Step 2: Configure Authentication

Copy the sample configuration and edit with your values:

```powershell
Copy-Item app-config.sample.json app-config.json
```

Edit `app-config.json` with your Azure AD app credentials (see [CONFIG-README.md](CONFIG-README.md) for detailed guidance):

```json
{
  "tenantId": "YOUR-TENANT-ID-GUID",
  "clientId": "YOUR-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "YOUR-CERTIFICATE-THUMBPRINT",
  "tenantDomain": "yourtenant.onmicrosoft.com"
}
```

**Validate your configuration:**
```powershell
.\Test-Configuration.ps1
```

This will check:
- Configuration file exists and is valid JSON
- All required fields are present
- Certificate exists and hasn't expired
- GUIDs are properly formatted

```json
{
  "tenantId": "YOUR-TENANT-ID-GUID",
  "clientId": "YOUR-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "YOUR-CERTIFICATE-THUMBPRINT",
  "tenantDomain": "yourtenant.onmicrosoft.com"
}
```

### Step 3: Export a Site

```powershell
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/SourceSite" `
    -IncludeContent
```

Output: `C:\PSReports\SiteTemplates\SiteTemplate_YYYYMMDD_HHMMSS.pnp`

### Step 4: Import to Target Site

**IMPORTANT**: The target site must exist before importing. Create it manually first:

**Option 1: SharePoint Admin Center**
1. Go to your tenant's admin center: `https://yourtenant-admin.sharepoint.com/_layouts/15/online/SiteCollections.aspx`
2. Click "Create" → Choose site type **that matches your source site**:
   - "Team site" if source was a Team Site
   - "Communication site" if source was a Communication Site
3. Enter site name and URL
4. Add owners/admins
5. Wait for provisioning to complete

**⚠️ IMPORTANT**: The target site type should match the source site type to avoid import errors. Check the export output for the source site type.

**Option 2: PowerShell**
```powershell
Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive
New-PnPSite -Type TeamSite -Title "My New Site" -Alias "MyNewSite" -Wait
Disconnect-PnPOnline
```

**Then import the template:**
```powershell
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://tenant.sharepoint.com/sites/TargetSite" `
    -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_20260203_142352.pnp"
```

## ⚠️ Cross-Tenant Scenarios

**IMPORTANT**: If your target site is in a **different tenant** than the source site:

1. **You need TWO app registrations**:
   - One in the **source tenant** (for export)
   - One in the **target tenant** (for import)

2. **Certificate setup in BOTH tenants**:
   - Same certificate can be used (upload .cer file to both app registrations)
   - Or generate separate certificates for each tenant

3. **Two configuration files**:
   ```powershell
   # Export from source tenant
   .\Export-SharePointSiteTemplate.ps1 `
       -SourceSiteUrl "https://sourcetenant.sharepoint.com/sites/Site" `
       -ConfigFile "app-config-source.json" `
       -IncludeContent
   
   # Import to target tenant
   .\Import-SharePointSiteTemplate.ps1 `
       -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" `
       -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp" `
       -ConfigFile "app-config-target.json"
   ```

4. **Each app-config file contains tenant-specific values**:
   ```json
   // app-config-source.json
   {
     "tenantId": "SOURCE-TENANT-ID",
     "clientId": "SOURCE-APP-CLIENT-ID",
     "certificateThumbprint": "CERT-THUMBPRINT",
     "tenantDomain": "sourcetenant.onmicrosoft.com"
   }
   
   // app-config-target.json
   {
     "tenantId": "TARGET-TENANT-ID",
     "clientId": "TARGET-APP-CLIENT-ID",
     "certificateThumbprint": "CERT-THUMBPRINT",
     "tenantDomain": "targettenant.onmicrosoft.com"
   }
   ```

5. **Setup steps for each tenant**:
   - Follow [MANUAL-APP-REGISTRATION.md](MANUAL-APP-REGISTRATION.md) in both tenants
   - Grant admin consent in each tenant
   - Ensure SharePoint API permissions in both apps

**Same Tenant**: If source and target are in the same tenant, use the same `app-config.json` for both operations.

## Script Reference

### Export-SharePointSiteTemplate.ps1

Exports a SharePoint site as a PnP template.

**Parameters**:
- `-SourceSiteUrl` (Required): URL of the site to export
- `-OutputPath` (Optional): Export location (default: `C:\PSReports\SiteTemplates`)
- `-TemplateName` (Optional): Template filename (default: auto-generated timestamp)
- `-IncludeContent` (Switch): Include list/library items in export
- `-ContentRowLimit` (Optional): Max items per list (default: 5000)
- `-ExcludeHandlers` (Optional): Comma-separated handlers to exclude
- `-ConfigFile` (Optional): Path to config file (default: `app-config.json`)

**Examples**:
```powershell
# Export with content
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectA" `
    -IncludeContent `
    -ContentRowLimit 10000

# Export structure only
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://contoso.sharepoint.com/sites/Template" `
    -TemplateName "SiteStructure_Only"

# Export excluding specific handlers
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://contoso.sharepoint.com/sites/Site" `
    -ExcludeHandlers "Workflows","SearchSettings"
```

### Import-SharePointSiteTemplate.ps1

Imports a PnP template to a SharePoint site.

**PREREQUISITE**: Target site must exist before running this script. See Quick Start Step 4 for site creation instructions.

**Parameters**:
- `-TargetSiteUrl` (Required): URL of the destination site (must already exist)
- `-TemplatePath` (Required): Path to the .pnp template file
- `-ClearNavigation` (Switch): Clear existing navigation before applying
- `-OverwriteSystemPropertyBagValues` (Switch): Overwrite system properties
- `-ProvisionFieldsToSite` (Switch): Provision fields to site collection
- `-ConfigFile` (Optional): Path to config file (default: `app-config.json`)
- `-WhatIf` (Switch): Preview changes without applying

**Examples**:
```powershell
# Standard import (site must exist)
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/NewSite" `
    -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp"

# Preview changes (WhatIf)
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/NewSite" `
    -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp" `
    -WhatIf

# Import with navigation clear
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://contoso.sharepoint.com/sites/NewSite" `
    -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp" `
    -ClearNavigation

# Cross-tenant import (use different config file)
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" `
    -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp" `
    -ConfigFile "app-config-target.json"
```

## Authentication Flow

The scripts use **certificate-based authentication** with the following priority:

1. **Certificate Authentication** (Primary)
   - Reads `certificateThumbprint` from `app-config.json`
   - Looks up certificate in `Cert:\CurrentUser\My`
   - Uses modern authentication with Azure AD
   - ✅ Bypasses Conditional Access restrictions
   - ✅ Most secure method
   - ✅ Fully automated

2. **Client Secret Authentication** (Fallback)
   - Uses `clientSecret` from `app-config.json`
   - Requires ACS (Azure Access Control Service) enabled
   - ⚠️ Legacy authentication method
   - ⚠️ May not work with strict security policies

3. **Interactive Authentication** (Last Resort)
   - Falls back if config file missing or incomplete
   - Opens browser for authentication
   - ❌ May be blocked by Conditional Access

## File Structure

```
SharepointTemplateExport/
├── Export-SharePointSiteTemplate.ps1     # Export site to PnP template
├── Import-SharePointSiteTemplate.ps1     # Import PnP template to site
├── Remove-DeletedSharePointSite.ps1      # Cleanup deleted sites from recycle bin
├── Test-Configuration.ps1                # Validate configuration files
├── Register-SharePointApp.ps1            # App registration helper (optional)
├── app-config.json                       # Authentication configuration (git-ignored)
├── app-config.sample.json                # Configuration template
├── app-config-source.example.json        # Example for source tenant
├── app-config-target.example.json        # Example for target tenant
├── CONFIG-README.md                      # Configuration setup guide
├── MANUAL-APP-REGISTRATION.md            # Detailed setup guide
└── Configuration Issues

**Before running any scripts**, validate your configuration:
```powershell
.\Test-Configuration.ps1
```

This comprehensive test will identify:
- Missing or invalid configuration files
- Incorrect JSON format
- Missing required fields
- Certificate availability and expiration
- Invalid GUID formats

### README.md                             # This file
```

## Troubleshooting

### "Unauthorized" Error (401)

**Symptoms**: Connection succeeds but queries fail

**Causes**:
1. Wrong API permissions (check you have **SharePoint** not Microsoft Graph)
2. Admin consent not granted
3. Certificate not uploaded or mismatched thumbprint

**Solutions**:
```powershell
# Verify certificate
Get-ChildItem Cert:\CurrentUser\My | Where-Object {
    $_.Thumbprint -eq "YOUR-THUMBPRINT-FROM-CONFIG"
}

# Check app permissions in Azure Portal:
# - Go to your app → API permissions
# - Should see: SharePoint - Sites.FullControl.All (Granted)
```

### Certificate Not Found

**Symptoms**: `Cannot find certificate with thumbprint...`

**Solution**:
```powershell
# List all certificates
Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint

# Verify thumbprint in app-config.json matches
```

### Conditional Access Blocking (Error 53003)

**Solution**: Use certificate authentication - it bypasses device restrictions

### Target Site Does Not Exist

**Symptoms**: "Target site does not exist" error when importing

**Solution**: Create the target site before running the import:
1. Via Admin Center: `https://yourtenant-admin.sharepoint.com/_layouts/15/online/SiteCollections.aspx`
2. Via PowerShell: `New-PnPSite -Type TeamSite -Title "Site Name" -Alias "SiteName" -Wait`
3. Then re-run the import script

### Site Type Mismatch Warnings

**Symptoms**: Warning during import: "source site had a base template ID value of X, while target has Y"

**Causes**: 
- Source was Communication Site, target is Team Site (or vice versa)
- Different site templates between source and target

**Impact**: 
- Some features may not import correctly
- Publishing pages may not work on Team Sites
- Certain lists/libraries may be skipped

**Solution**:
1. **Recommended**: Create target site matching source type:
   - Check export output for source site type
   - Create new target site with matching type
   - Re-run import
2. **Alternative**: Continue import with `-IgnoreDuplicateDataRowErrors` (some features may not work)

### ACS Disabled

**Symptoms**: Client secret authentication fails

**Solution**:
1. Use certificate authentication (recommended), OR
2. Enable ACS: SharePoint Admin → Policies → Access control → Allow apps that don't use modern authentication

## Security Best Practices

1. **Use Certificate Authentication**
   - More secure than client secrets
   - Uses modern authentication
   - Bypasses most Conditional Access restrictions

2. **Protect Credentials**
   - Keep `app-config.json` secure
   - Don't commit to source control
   - Add to `.gitignore`

3. **Certificate Management**
   - Store certificates securely
   - Rotate before expiration (recommend 2-year validity)
   - Keep private key protected

4. **Audit Logging**
   - All operations logged to `C:\PSReports\SiteTemplates\`
   - Review logs for security and compliance
   - Transcript files contain full operation history

5. **Least Privilege**
   - App has `Sites.FullControl.All` - only use for migrations
   - Consider creating separate apps for different purposes
   - Revoke access when migration complete

## Advanced Scenarios

### Cleanup Deleted Sites

After testing migrations, you may need to permanently remove sites from the recycle bin:

```powershell
# List all deleted sites
.\Remove-DeletedSharePointSite.ps1 -ListAll

# Permanently delete a specific site
.\Remove-DeletedSharePointSite.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/TestSite"

# Permanently delete ALL deleted sites (use with caution!)
.\Remove-DeletedSharePointSite.ps1 -DeleteAll -Force
```

**⚠️ WARNING**: This permanently deletes sites. They cannot be recovered.

### Exclude Specific Content

```powershell
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Site" `
    -ExcludeHandlers "Workflows","AuditSettings","SitePolicy"
```

### Large Site Export

```powershell
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/LargeSite" `
    -IncludeContent `
    -ContentRowLimit 10000  # Adjust based on needs
```

### Multiple Site Migration

```powershell
# Export multiple sites
$sites = @(
    "https://tenant.sharepoint.com/sites/Site1",
    "https://tenant.sharepoint.com/sites/Site2",
    "https://tenant.sharepoint.com/sites/Site3"
)

foreach ($site in $sites) {
    $siteName = $site.Split('/')[-1]
    .\Export-SharePointSiteTemplate.ps1 `
        -SourceSiteUrl $site `
        -TemplateName $siteName `
        -IncludeContent
}
```

## Logs and Output

All operations create detailed logs:

- **Export logs**: `C:\PSReports\SiteTemplates\SiteTemplate_*.log`
- **Import logs**: `C:\PSReports\SiteTemplates\ImportLogs\Import_*.log`
- **Template files**: `C:\PSReports\SiteTemplates\*.pnp`

Logs include:
- Connection details (masked credentials)
- Site information before/after
- List of items exported/imported
- Errors and warnings
- Operation duration

## Support

For detailed Azure AD app registration setup, see:
- [MANUAL-APP-REGISTRATION.md](MANUAL-APP-REGISTRATION.md) - Complete step-by-step guide

## Version History

- **v2.0** (February 2026)
  - Certificate-based authentication (primary method)
  - Client secret fallback support
  - Improved error handling and logging
  - Conditional Access bypass support

- **v1.0** (Initial Release)
  - Basic export/import functionality
  - Interactive authentication only
