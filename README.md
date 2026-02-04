# SharePoint Site Template Export/Import Scripts

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue.svg)](https://github.com/PowerShell/PowerShell)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![PnP PowerShell](https://img.shields.io/badge/PnP.PowerShell-2.12%2B-green.svg)](https://pnp.github.io/powershell/)

PowerShell scripts for exporting and importing complete SharePoint sites using PnP templates with **certificate-based authentication**.

> **For Contributors**: See [DEVELOPER.md](DEVELOPER.md) for code standards, testing requirements, and contribution guidelines.

## Overview

These scripts enable you to:
- **Export** a complete SharePoint site including structure, content, pages, and lists
- **Import** the template to another SharePoint site (same or different target tenant)
- **Map users** during cross-tenant migrations (similar to BitTitan MigrationWiz and Sharegate)
- Use **secure certificate-based authentication** (no interactive login required)
- Maintain full audit trail of operations

### ✨ Key Features

#### 🔄 **Cross-Tenant User Mapping** (New!)
- Automatically extract all users from source site or template
- Generate user mapping CSV template
- Validate target users before migration
- Map site permissions, metadata, and people picker fields
- Comprehensive audit trail

#### 🔐 **Secure Authentication**
- Certificate-based authentication (recommended)
- Azure AD App Registration support
- No passwords in scripts or config files
- Multi-tenant support

#### 📦 **Complete Site Migration**
- Site structure (lists, libraries, content types)
- Site pages and navigation
- Document libraries with files
- List items with all columns
- Site permissions and groups
- Custom fields and metadata

#### 🛡️ **Enterprise-Ready**
- Detailed logging and error handling
- WhatIf mode for preview
- Validation before import
- Configurable content limits
- Cross-tenant migration support

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

**For cross-tenant migrations, validate both configurations:**
```powershell
.\Test-Configuration.ps1 -SourceConfigFile "app-config-source.json" -TargetConfigFile "app-config-target.json"
```

This will check both source and target configurations and show a side-by-side comparison to ensure you're ready for cross-tenant migration.

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

## 🔄 Cross-Tenant User Mapping

When migrating SharePoint sites between different tenants, user identities need to be mapped from source to target tenant. This toolset includes comprehensive user mapping functionality similar to enterprise migration tools like BitTitan MigrationWiz and Sharegate.

### Why User Mapping is Needed

SharePoint stores user references in:
- **Site permissions** (administrators, group members, role assignments)
- **List/library permissions** (item-level security)
- **Metadata fields** (Created By, Modified By, Author, Editor)
- **People picker fields** (custom user columns in lists)
- **Workflow assignments** and **task assignments**

During cross-tenant migration, source tenant users (e.g., `user@source.onmicrosoft.com`) must be mapped to target tenant users (e.g., `user@target.onmicrosoft.com`).

### User Mapping Workflow

#### 1. Export the Site Template

First, export the site from the source tenant:

```powershell
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://sourcetenant.sharepoint.com/sites/ProjectA" `
    -ConfigFile "app-config-source.json" `
    -IncludeContent
```

Output: `C:\PSReports\SiteTemplates\SiteTemplate_20260204_143022.pnp`

#### 2. Generate User Mapping Template

The `New-UserMappingTemplate.ps1` script extracts all unique users from the exported template and creates a CSV mapping file:

```powershell
.\New-UserMappingTemplate.ps1 `
    -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_20260204_143022.pnp"
```

This creates `user-mapping-template.csv` with the following structure:

```csv
SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes
john.smith@sourcetenant.com,john.smith@sourcetenant.com,John Smith,John Smith,Found in: Site User
sarah.jones@sourcetenant.com,sarah.jones@sourcetenant.com,Sarah Jones,Sarah Jones,Found in: Group: Members
admin@sourcetenant.com,admin@sourcetenant.com,Admin User,Admin User,Found in: Site Administrator
```

**Alternative**: Scan a live source site instead of a template file:

```powershell
.\New-UserMappingTemplate.ps1 `
    -SiteUrl "https://sourcetenant.sharepoint.com/sites/ProjectA" `
    -ConfigFile "app-config-source.json" `
    -OutputPath "C:\Migrations\users.csv"
```

#### 3. Edit the User Mapping File

Open `user-mapping-template.csv` and update the **TargetUser** column with target tenant email addresses:

```csv
SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes
john.smith@sourcetenant.com,john.smith@targettenant.com,John Smith,John Smith,Same person different tenant
sarah.jones@sourcetenant.com,sarah.j@targettenant.com,Sarah Jones,Sarah Jones,Email changed in target
admin@sourcetenant.com,it.admin@targettenant.com,Admin User,IT Admin,Role reassigned
old.user@sourcetenant.com,,Old User,,No longer with company - skip
```

**Important**:
- Leave **TargetUser** empty to skip mapping (user will be unmapped)
- Update **TargetDisplayName** if names differ in target tenant
- All target users must exist in the target tenant before import

#### 4. Validate Target Users (Optional but Recommended)

Before performing the full import, validate that all target users exist:

```powershell
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/ProjectA" `
    -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_20260204_143022.pnp" `
    -UserMappingFile "user-mapping-template.csv" `
    -ConfigFile "app-config-target.json" `
    -ValidateUsersOnly
```

This will:
- ✅ Check if each target user exists in the target tenant
- ✅ Attempt to add users to the site (if they exist in the tenant)
- ❌ Report any invalid users
- ⏸️ Stop before performing the actual import

Example output:
```
═══════════════════════════════════════════════════════
  User Validation Results
═══════════════════════════════════════════════════════
  Valid users:   15
  Invalid users: 2
═══════════════════════════════════════════════════════

Invalid Users:
  • old.user@sourcetenant.com → old.user@targettenant.com
    Reason: User not found in target tenant
  • external@partner.com → external@partner.com
    Reason: User not found in target tenant
```

Fix any errors in your CSV file and re-validate until all users pass.

#### 5. Import with User Mapping

Once validation passes, perform the import with user mapping:

```powershell
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/ProjectA" `
    -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_20260204_143022.pnp" `
    -UserMappingFile "user-mapping-template.csv" `
    -ConfigFile "app-config-target.json" `
    -IgnoreDuplicateDataRowErrors
```

**What happens during import:**
1. User mapping CSV is loaded and validated
2. Target users are validated in the target tenant
3. A temporary modified template is created with mapped users
4. The modified template is applied to the target site
5. All user references are automatically mapped:
   - Site permissions → mapped to target users
   - List item metadata (Created By, Modified By) → mapped
   - People picker fields → mapped
   - User fields in list items → mapped
6. Temporary files are cleaned up

**Parameters explained:**
- `-UserMappingFile`: Path to your user mapping CSV
- `-IgnoreDuplicateDataRowErrors`: Recommended for cross-tenant migrations to skip duplicate content errors
- `-ValidateUsersOnly`: Pre-flight validation only, no import

### User Mapping Features

✅ **Comprehensive Mapping**
- Maps users in site security (administrators, groups, roles)
- Maps users in list/library permissions
- Maps metadata fields (Author, Editor, Created By, Modified By)
- Maps people picker columns and custom user fields

✅ **Pre-Flight Validation**
- Validates target users exist before import
- Reports missing or invalid users
- Prevents failed imports due to user issues

✅ **Flexible Mapping**
- CSV-based for easy editing and version control
- Supports one-to-one mapping (same user in different tenant)
- Supports role changes (old admin → new admin)
- Skip unmapped users (leave TargetUser empty)

✅ **Audit Trail**
- All mappings logged during import
- Source and target users tracked
- Import logs include user mapping details

### User Mapping CSV Format

The CSV file must have the following columns:

| Column | Required | Description |
|--------|----------|-------------|
| `SourceUser` | ✅ Yes | Source tenant user email (case-insensitive) |
| `TargetUser` | ⚠️ Optional | Target tenant user email. Leave empty to skip mapping. |
| `SourceDisplayName` | ℹ️ Info | Display name in source tenant (for reference) |
| `TargetDisplayName` | ℹ️ Info | Display name in target tenant (for reference) |
| `Notes` | ℹ️ Info | Any notes (e.g., "Same person", "Role changed") |

**Example** (`user-mapping.sample.csv` provided):
```csv
SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes
john.smith@source.onmicrosoft.com,john.smith@target.onmicrosoft.com,John Smith,John Smith,Same user different tenant
sarah.jones@source.onmicrosoft.com,sarah.jones@target.onmicrosoft.com,Sarah Jones,Sarah Jones,Same user different tenant
old.admin@source.onmicrosoft.com,new.admin@target.onmicrosoft.com,Old Admin,New Admin,Admin role changed
legacy.user@source.onmicrosoft.com,,Legacy User,,User no longer exists - will be skipped
```

### Troubleshooting User Mapping

**Issue**: "User not found in target tenant"
- **Solution**: Ensure the user exists in the target tenant (check Azure AD)
- **Solution**: User may need to be licensed in the target tenant
- **Solution**: External users may need to be invited as guests first

**Issue**: "User validation failed"
- **Solution**: Run with `-ValidateUsersOnly` to see specific errors
- **Solution**: Update CSV to fix invalid mappings or remove them
- **Solution**: Use `-IgnoreDuplicateDataRowErrors` to continue despite some errors

**Issue**: "Some permissions missing after import"
- **Solution**: Verify all users in mapping file
- **Solution**: Some system accounts cannot be mapped (e.g., SharePoint App)
- **Solution**: Manually review and assign permissions post-migration

**Issue**: "Mapping not applied to some content"
- **Solution**: PnP template limitations may skip some user references
- **Solution**: Use `-IncludeContent` during export to capture list items
- **Solution**: Re-run mapping after initial import if needed

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
- `-UserMappingFile` (Optional): Path to user mapping CSV for cross-tenant migrations
- `-ValidateUsersOnly` (Switch): Only validate target users, don't perform import
- `-IgnoreDuplicateDataRowErrors` (Switch): Continue on duplicate/malformed data errors (recommended for cross-tenant)
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

# Cross-tenant with user mapping
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" `
    -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp" `
    -UserMappingFile "user-mapping.csv" `
    -ConfigFile "app-config-target.json" `
    -IgnoreDuplicateDataRowErrors

# Validate users before import
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" `
    -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp" `
    -UserMappingFile "user-mapping.csv" `
    -ValidateUsersOnly
```

### New-UserMappingTemplate.ps1

Generates a user mapping template CSV from a SharePoint site or template file.

**Parameters**:
- `-TemplatePath` (Required for template mode): Path to exported .pnp template file
- `-SiteUrl` (Required for live site mode): URL of SharePoint site to scan
- `-OutputPath` (Optional): Output CSV path (default: `user-mapping-template.csv`)
- `-IncludeSystemAccounts` (Switch): Include system accounts like SharePoint App
- `-ConfigFile` (Optional): Config file for site authentication (when using `-SiteUrl`)

**Examples**:
```powershell
# Generate from exported template
.\New-UserMappingTemplate.ps1 `
    -TemplatePath "C:\PSReports\SiteTemplates\Template.pnp"

# Generate from live site
.\New-UserMappingTemplate.ps1 `
    -SiteUrl "https://contoso.sharepoint.com/sites/ProjectA" `
    -ConfigFile "app-config-source.json"

# Custom output path
.\New-UserMappingTemplate.ps1 `
    -TemplatePath "Template.pnp" `
    -OutputPath "C:\Migrations\ProjectA\users.csv"

# Include system accounts
.\New-UserMappingTemplate.ps1 `
    -TemplatePath "Template.pnp" `
    -IncludeSystemAccounts
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
├── Core Scripts
│   ├── Export-SharePointSiteTemplate.ps1     # Export site to PnP template (with selective export)
│   ├── Import-SharePointSiteTemplate.ps1     # Import PnP template to site (with selective import)
│   ├── New-UserMappingTemplate.ps1           # Generate user mapping CSV for cross-tenant migrations
│   ├── Get-TemplateContent.ps1               # Inspect and analyze .pnp template files (NEW v3.0)
│   └── Compare-Templates.ps1                 # Compare two templates and show differences (NEW v3.0)
├── Utility Scripts
│   ├── Remove-DeletedSharePointSite.ps1      # Cleanup deleted sites from recycle bin
│   ├── Test-Configuration.ps1                # Validate configuration files
│   └── Register-SharePointApp.ps1            # App registration helper (optional)
├── Configuration Files
│   ├── app-config.json                       # Authentication config (git-ignored)
│   ├── app-config.sample.json                # Configuration template
│   ├── app-config-source.example.json        # Example for source tenant
│   ├── app-config-target.example.json        # Example for target tenant
│   └── user-mapping.sample.csv               # User mapping CSV template
├── Documentation
│   ├── README.md                             # This file - main documentation
│   ├── CONFIG-README.md                      # Configuration setup guide
│   ├── MANUAL-APP-REGISTRATION.md            # Detailed app registration guide
│   ├── USER-MAPPING-QUICK-REF.md             # Quick reference card (NEW)
│   ├── USER-MAPPING-TEST-GUIDE.md            # Testing scenarios (NEW)
│   └── DEVELOPER.md                          # Developer guide and contribution standards (NEW)
└── LICENSE                                   # MIT License
```

## Configuration Issues

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

---

## Selective Export & Import Features

Gain precise control over what gets exported and imported with selective migration capabilities.

### 🔍 Template Inspection

Before exporting or importing, inspect template contents to make informed decisions:

```powershell
.\Get-TemplateContent.ps1 -TemplatePath "C:\PSReports\SiteTemplates\MySite.pnp" -Detailed -ShowUsers -ShowContent
```

**Displays:**
- 📋 Lists and Libraries (with item counts and types)
- 📄 Pages (with titles and layouts)
- 👥 Users (extracted from permissions and metadata)
- 📑 Content Types
- 🔧 Site Columns/Fields
- ⚙️ Features
- 🔒 Security (groups, permissions)
- 🧭 Navigation

**Options:**
```powershell
# Show specific information
.\Get-TemplateContent.ps1 -TemplatePath "template.pnp" -ShowUsers

# Show detailed analysis with user and content information
.\Get-TemplateContent.ps1 -TemplatePath "template.pnp" -Detailed -ShowUsers -ShowContent

# Export analysis to JSON
.\Get-TemplateContent.ps1 -TemplatePath "template.pnp" -OutputFormat JSON -OutputPath "analysis.json"

# Compare two templates
.\Get-TemplateContent.ps1 -TemplatePath "template1.pnp" -CompareTo "template2.pnp"
```

### 📤 Selective Export

Export only what you need instead of the entire site:

#### Export Specific Lists/Libraries

```powershell
# Include only specific lists
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Project" `
    -IncludeLists "Documents","Project Tasks","Issues" `
    -IncludeContent

# Exclude specific lists
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Project" `
    -ExcludeLists "Archive","Old Documents","Temp Data" `
    -IncludeContent
```

#### Export Structure Only (No Content)

```powershell
# Export list/library schemas without any data
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Project" `
    -StructureOnly
```

Perfect for creating site templates, setting up development environments, or schema-only migrations.

#### Export Without Pages

```powershell
# Exclude all site pages from export
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Project" `
    -ExcludePages `
    -IncludeContent
```

#### Preview Export

See what will be exported without creating the template file:

```powershell
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Project" `
    -Preview
```

**Preview shows:**
- Export configuration (content mode, structure only, etc.)
- Lists/libraries that will be exported (with item counts)
- Pages that will be exported (if not excluded)

### 📥 Selective Import

Control what gets imported to avoid overwriting critical content:

#### Inspect Before Importing

```powershell
# Preview what's in the template without importing
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://tenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\source.pnp" `
    -InspectOnly
```

#### Import Specific Components

```powershell
# Import only lists and pages
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://tenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\source.pnp" `
    -ImportComponents Lists,Pages
```

**Available Components:** All, Lists, Libraries, Pages, Navigation, Security, ContentTypes, Fields, Features

#### Import Specific Lists

```powershell
# Import only specific lists
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://tenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\source.pnp" `
    -IncludeLists "Documents","Project Tasks"

# Import all except specific lists
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://tenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\source.pnp" `
    -ExcludeLists "Archive","Temp Data"
```

#### Import Structure Only

```powershell
# Import list schemas without content
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://tenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\source.pnp" `
    -StructureOnly
```

#### Skip Existing Lists

```powershell
# Don't recreate lists that already exist
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://tenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\source.pnp" `
    -SkipExisting
```

### ⚖️ Template Comparison

Compare two templates to see differences:

```powershell
.\Compare-Templates.ps1 `
    -Template1Path "C:\Templates\before.pnp" `
    -Template2Path "C:\Templates\after.pnp"
```

**Comparison shows:**
- Items only in Template 1 (removed)
- Items only in Template 2 (added)
- Items in both (unchanged)
- Total differences per component

**Export formats:**

```powershell
# Save as HTML report
.\Compare-Templates.ps1 -Template1Path "t1.pnp" -Template2Path "t2.pnp" `
    -OutputFormat HTML -OutputPath "comparison.html"

# Save as JSON or CSV
.\Compare-Templates.ps1 -Template1Path "t1.pnp" -Template2Path "t2.pnp" `
    -OutputFormat JSON -OutputPath "comparison.json"
```

**Compare specific components:**

```powershell
.\Compare-Templates.ps1 `
    -Template1Path "v1.pnp" `
    -Template2Path "v2.pnp" `
    -CompareComponents Lists,Pages,Users
```

### 🎯 Complete Selective Migration Workflow

```powershell
# 1. Preview source site export
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Source" `
    -Preview

# 2. Export selectively
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://tenant.sharepoint.com/sites/Source" `
    -ExcludeLists "Archive","Temp" `
    -ExcludePages `
    -IncludeContent `
    -TemplateName "SourceSite_Selective"

# 3. Inspect exported template
.\Get-TemplateContent.ps1 `
    -TemplatePath "C:\PSReports\SiteTemplates\SourceSite_Selective.pnp" `
    -Detailed -ShowUsers -ShowContent

# 4. Compare with existing target (if applicable)
.\Compare-Templates.ps1 `
    -Template1Path "C:\Templates\Target_Current.pnp" `
    -Template2Path "C:\Templates\SourceSite_Selective.pnp" `
    -OutputFormat HTML

# 5. Inspect before importing
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\SourceSite_Selective.pnp" `
    -InspectOnly

# 6. Import selectively with user mapping
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Target" `
    -TemplatePath "C:\Templates\SourceSite_Selective.pnp" `
    -ImportComponents Lists,Navigation,Security `
    -SkipExisting `
    -UserMappingFile "user-mapping.csv" `
    -ConfigFile "app-config-target.json"
```

---

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

For detailed guides and documentation:
- [MANUAL-APP-REGISTRATION.md](MANUAL-APP-REGISTRATION.md) - Complete Azure AD app setup
- [CONFIG-README.md](CONFIG-README.md) - Configuration file guidance
- [USER-MAPPING-QUICK-REF.md](USER-MAPPING-QUICK-REF.md) - User mapping quick reference
- [USER-MAPPING-TEST-GUIDE.md](USER-MAPPING-TEST-GUIDE.md) - Testing scenarios
- [DEVELOPER.md](DEVELOPER.md) - Contribution guidelines and code standards

## Contributing

We welcome contributions! Please see [DEVELOPER.md](DEVELOPER.md) for:
- Code standards and style guide
- Feature development workflow
- Testing requirements
- Pull request process
- Security guidelines

## Version History

- **v3.0** (February 2026)
  - **Selective Export & Import**: Granular control over what gets migrated
  - **Template Inspection Tool**: Analyze .pnp files before import
  - **Template Comparison**: Compare two templates to identify differences
  - Export/Import specific lists by name (whitelist or blacklist)
  - Structure-only mode for schema migration
  - Preview mode to see what will be exported
  - InspectOnly mode to analyze templates before importing
  - Component-level filtering (Lists, Pages, Navigation, etc.)
  - SkipExisting option to avoid overwriting content
  - New scripts: Get-TemplateContent.ps1, Compare-Templates.ps1

- **v2.0** (February 2026)
  - Cross-tenant user mapping functionality
  - User extraction and validation tools
  - Certificate-based authentication (primary method)
  - Client secret fallback support
  - Improved error handling and logging
  - Conditional Access bypass support
  - New scripts: New-UserMappingTemplate.ps1
  - Developer guide (DEVELOPER.md) with contribution standards

- **v1.0** (Initial Release)
  - Basic export/import functionality
  - Interactive authentication only
