# Cross-Tenant User Mapping - Quick Reference

## 🚀 Quick Start (4 Steps)

### 1️⃣ Export from Source Tenant
```powershell
.\Export-SharePointSiteTemplate.ps1 `
    -SourceSiteUrl "https://sourcetenant.sharepoint.com/sites/MySite" `
    -ConfigFile "app-config-source.json" `
    -IncludeContent
```
**Output**: `C:\PSReports\SiteTemplates\SiteTemplate_YYYYMMDD_HHMMSS.pnp`

---

### 2️⃣ Generate User Mapping Template
```powershell
.\New-UserMappingTemplate.ps1 `
    -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_YYYYMMDD_HHMMSS.pnp"
```
**Output**: `user-mapping-template.csv`

---

### 3️⃣ Edit User Mapping CSV

Open `user-mapping-template.csv` and update **TargetUser** column:

| SourceUser | TargetUser | Notes |
|------------|------------|-------|
| john@source.com | john@target.com | ✅ Same user |
| sarah@source.com | sarah.new@target.com | ✅ Email changed |
| old@source.com | new@target.com | ✅ Role change |
| ghost@source.com | **(leave empty)** | ⏭️ Skip this user |

---

### 4️⃣ Validate & Import to Target Tenant

**First, validate users:**
```powershell
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/MySite" `
    -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_YYYYMMDD_HHMMSS.pnp" `
    -UserMappingFile "user-mapping-template.csv" `
    -ConfigFile "app-config-target.json" `
    -ValidateUsersOnly
```

**If validation passes, import:**
```powershell
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://targettenant.sharepoint.com/sites/MySite" `
    -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_YYYYMMDD_HHMMSS.pnp" `
    -UserMappingFile "user-mapping-template.csv" `
    -ConfigFile "app-config-target.json" `
    -IgnoreDuplicateDataRowErrors
```

---

## 📋 CSV Format

**Required columns:**
- `SourceUser` - Source tenant email
- `TargetUser` - Target tenant email (or empty to skip)

**Optional columns:**
- `SourceDisplayName` - For reference
- `TargetDisplayName` - For reference
- `Notes` - Any notes

**Example:**
```csv
SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes
user1@source.com,user1@target.com,User One,User One,Mapped
user2@source.com,,User Two,,Skipped
```

---

## ✅ What Gets Mapped

- ✅ Site administrators
- ✅ Site group members (Owners, Members, Visitors)
- ✅ List/library permissions
- ✅ Created By / Modified By metadata
- ✅ Author / Editor fields
- ✅ People Picker columns
- ✅ Custom user fields in lists

---

## 🔧 Common Commands

### Generate from Live Site (Instead of Template)
```powershell
.\New-UserMappingTemplate.ps1 `
    -SiteUrl "https://sourcetenant.sharepoint.com/sites/MySite" `
    -ConfigFile "app-config-source.json"
```

### Custom Output Path
```powershell
.\New-UserMappingTemplate.ps1 `
    -TemplatePath "Template.pnp" `
    -OutputPath "C:\Migrations\ProjectA\users.csv"
```

### Include System Accounts
```powershell
.\New-UserMappingTemplate.ps1 `
    -TemplatePath "Template.pnp" `
    -IncludeSystemAccounts
```

### Preview Import (WhatIf)
```powershell
.\Import-SharePointSiteTemplate.ps1 `
    -TargetSiteUrl "https://target.sharepoint.com/sites/Site" `
    -TemplatePath "Template.pnp" `
    -UserMappingFile "users.csv" `
    -WhatIf
```

---

## 🚨 Troubleshooting

### ❌ "User not found in target tenant"
**Fix**: Ensure user exists in Azure AD and is licensed

### ❌ "User validation failed"
**Fix**: Run with `-ValidateUsersOnly` to see specific errors

### ❌ "Missing required column: SourceUser"
**Fix**: Ensure CSV has `SourceUser` and `TargetUser` columns

### ❌ Import fails with user errors
**Fix**: Add `-IgnoreDuplicateDataRowErrors` parameter

### ⚠️ Some permissions missing after import
**Fix**: Manually review and adjust permissions post-migration

---

## 📊 Validation Output Example

```
═══════════════════════════════════════════════════════
  User Validation Results
═══════════════════════════════════════════════════════
  Valid users:   15
  Invalid users: 2
═══════════════════════════════════════════════════════

Invalid Users:
  • old.user@source.com → old.user@target.com
    Reason: User not found in target tenant
```

**Action**: Update CSV to fix invalid users or remove them.

---

## 🎯 Best Practices

1. **Always validate first** with `-ValidateUsersOnly`
2. **Test in non-production** environment first
3. **Backup target site** before import
4. **Use version control** for user mapping CSV files
5. **Document role changes** in Notes column
6. **Review permissions** after migration
7. **Notify users** of email/role changes
8. **Keep mapping files** for audit trail

---

## 📞 Need Help?

- 📖 Full documentation: [README.md](README.md)
- 🧪 Testing guide: [USER-MAPPING-TEST-GUIDE.md](USER-MAPPING-TEST-GUIDE.md)
- 🔧 Configuration help: [CONFIG-README.md](CONFIG-README.md)
- 🏗️ App setup: [MANUAL-APP-REGISTRATION.md](MANUAL-APP-REGISTRATION.md)

---

**Version**: 1.0 | **Updated**: February 4, 2026
