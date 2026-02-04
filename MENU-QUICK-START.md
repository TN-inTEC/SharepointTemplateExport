# Interactive Menu Quick Start

## Launch the Interactive Menu

```powershell
.\Start-SharePointMigration.ps1
```

## Menu Structure

### Main Menu Options

1. **Same-Tenant Migration** - Migrate within one Microsoft 365 tenant
2. **Cross-Tenant Migration** - Migrate between different Microsoft 365 tenants  
3. **View Documentation** - Access setup guides and references
4. **Test Configuration Files** - Validate your configurations
0. **Exit** - Close the menu

---

## Same-Tenant Migration Workflow

### What You'll See

1. **Prerequisites Display**
   - Azure AD App Registration requirements
   - Certificate setup needs
   - Configuration file format example
   - Current prerequisite status check

2. **Configuration Validation**
   - Option to run `Test-Configuration.ps1`
   - Validates certificate, GUIDs, and connectivity

3. **Step-by-Step Workflow**
   - Export source site
   - Inspect template (optional)
   - Create target site
   - Import to target

### Prerequisites Needed

- ✅ PnP.PowerShell module installed
- ✅ Azure AD app registration in your tenant
- ✅ Certificate uploaded and configured
- ✅ `app-config.json` file created

### Example Configuration File

```json
{
  "tenantId": "YOUR-TENANT-ID-GUID",
  "clientId": "YOUR-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "YOUR-CERTIFICATE-THUMBPRINT",
  "tenantDomain": "yourtenant.onmicrosoft.com"
}
```

---

## Cross-Tenant Migration Workflow

### What You'll See

1. **Prerequisites Display**
   - TWO Azure AD app registrations needed (source + target)
   - Certificate setup in BOTH tenants
   - Two configuration file examples
   - Current prerequisite status check

2. **Configuration Validation**
   - Option to run `Test-Configuration.ps1` with both configs
   - Side-by-side validation of source and target
   - Warns if both are same tenant

3. **Step-by-Step Workflow with User Mapping**
   - Export from source tenant
   - Generate user mapping CSV
   - Edit user mappings
   - Validate target users
   - Create target site
   - Import with user mapping

### Prerequisites Needed

- ✅ PnP.PowerShell module installed
- ✅ Azure AD app registration in SOURCE tenant
- ✅ Azure AD app registration in TARGET tenant
- ✅ Certificate uploaded to BOTH app registrations
- ✅ `app-config-source.json` file created
- ✅ `app-config-target.json` file created

### Example Configuration Files

**Source Tenant** (`app-config-source.json`):
```json
{
  "tenantId": "SOURCE-TENANT-ID-GUID",
  "clientId": "SOURCE-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "CERT-THUMBPRINT",
  "tenantDomain": "sourcetenant.onmicrosoft.com"
}
```

**Target Tenant** (`app-config-target.json`):
```json
{
  "tenantId": "TARGET-TENANT-ID-GUID",
  "clientId": "TARGET-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "CERT-THUMBPRINT",
  "tenantDomain": "targettenant.onmicrosoft.com"
}
```

> **Note**: You can use the same certificate in both tenants - just upload the `.cer` file to both Azure AD app registrations.

---

## Documentation Menu

Access key documentation directly from the menu:

1. **README.md** - Main documentation and complete workflows
2. **MANUAL-APP-REGISTRATION.md** - Step-by-step Azure AD app setup
3. **CONFIG-README.md** - Configuration file detailed guidance
4. **USER-MAPPING-QUICK-REF.md** - User mapping CSV format and examples
5. **DEVELOPER.md** - Contribution guidelines and code standards

All documents open in Notepad for easy reading.

---

## Test Configuration Menu

Validate your configuration files before starting migration:

### Option 1: Single Configuration
- Tests one config file (for same-tenant migrations)
- Default: `app-config.json`
- Checks certificate, GUIDs, tenant connectivity

### Option 2: Cross-Tenant Configurations
- Tests both source and target configs
- Default: `app-config-source.json` + `app-config-target.json`
- Shows side-by-side comparison
- Warns if both point to same tenant

---

## Menu Navigation Tips

- **Choose options by number**: Type the number and press Enter
- **Return to main menu**: Select option `0` or press Ctrl+C
- **Validation recommended**: Always test configs before migration
- **Follow workflows**: The menu provides exact commands to run
- **Copy commands**: Commands shown can be copied/pasted into PowerShell

---

## Color Guide

The menu uses colors to help you navigate:

- 🔵 **Cyan** - Headers and main structure
- 🟡 **Yellow** - Prompts and important notes
- 🟢 **Green** - Success indicators and step numbers
- 🟣 **Magenta** - Info boxes and prerequisites
- 🔴 **Red** - Exit option and errors
- ⚪ **White** - Normal text and commands
- ⚫ **Gray** - Secondary information and examples

---

## Common Workflows

### First-Time User (Same Tenant)

1. Launch menu: `.\Start-SharePointMigration.ps1`
2. Select: `[1] Same-Tenant Migration`
3. Review prerequisites
4. Choose `[Y]` to validate configuration
5. Review workflow steps
6. Exit menu and follow the workflow commands

### First-Time User (Cross-Tenant)

1. Launch menu: `.\Start-SharePointMigration.ps1`
2. Select: `[2] Cross-Tenant Migration`
3. Review prerequisites (note: TWO app registrations needed)
4. Choose `[Y]` to validate both configurations
5. Review workflow with user mapping steps
6. Exit menu and follow the workflow commands

### Quick Config Test

1. Launch menu: `.\Start-SharePointMigration.ps1`
2. Select: `[4] Test Configuration Files`
3. Choose single or cross-tenant test
4. Enter config file paths (or use defaults)
5. Review validation results

### Documentation Lookup

1. Launch menu: `.\Start-SharePointMigration.ps1`
2. Select: `[3] View Documentation`
3. Choose document to view
4. Document opens in Notepad

---

## Troubleshooting

### Menu Doesn't Start
- Ensure you're in the correct directory: `C:\.GitLocal\SharepointTemplateExport`
- Check PowerShell execution policy: `Get-ExecutionPolicy`
- If restricted: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`

### Prerequisites Check Fails
- Install PnP.PowerShell: `Install-Module PnP.PowerShell -Scope CurrentUser`
- Create configuration files from samples
- See MANUAL-APP-REGISTRATION.md for app setup

### Configuration Validation Fails
- Verify GUIDs are correct format (no spaces/quotes)
- Check certificate exists: `Get-ChildItem Cert:\CurrentUser\My`
- Ensure certificate thumbprint matches config file
- See CONFIG-README.md for troubleshooting

### Documentation Won't Open
- Files must exist in current directory
- Notepad.exe must be available on system PATH
- Alternative: Open files manually in your preferred editor

---

## Next Steps After Using Menu

1. **Prerequisites Complete?** 
   - Follow signposted documentation if not ready
   - Run Test-Configuration.ps1 to validate

2. **Ready to Migrate?**
   - Copy workflow commands from menu output
   - Run commands one by one
   - Monitor logs in `C:\PSReports\SiteTemplates\`

3. **Cross-Tenant Migration?**
   - Generate user mapping: `.\New-UserMappingTemplate.ps1`
   - Edit CSV with target user mappings
   - Validate users before full import

4. **Need Help?**
   - Review README.md for complete documentation
   - Check troubleshooting sections
   - See DEVELOPER.md for advanced scenarios

---

## Pro Tips

💡 **Use WhatIf**: Add `-WhatIf` to import commands to preview changes

💡 **Inspect First**: Use `Get-TemplateContent.ps1` before importing to see what's in the template

💡 **Validate Users**: Always use `-ValidateUsersOnly` before cross-tenant import

💡 **Test Small**: Start with a simple site before migrating complex production sites

💡 **Keep Logs**: All operations log to `C:\PSReports\SiteTemplates\` - review for troubleshooting

💡 **Version Control**: Keep config files and user mappings in version control (but exclude from git)

---

## Getting Help

- **Quick Reference**: This document (MENU-QUICK-START.md)
- **Full Documentation**: README.md
- **Setup Guide**: MANUAL-APP-REGISTRATION.md
- **Config Help**: CONFIG-README.md
- **User Mapping**: USER-MAPPING-QUICK-REF.md
- **Contributing**: DEVELOPER.md

**Need more help?** Review the documentation files via the menu's Documentation option (option 3).
