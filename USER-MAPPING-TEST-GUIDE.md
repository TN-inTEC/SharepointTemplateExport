# User Mapping Testing Guide

## Overview

This document provides comprehensive testing scenarios and validation approaches for the cross-tenant user mapping feature in the SharePoint Template Export/Import tooling.

## Test Environment Setup

### Prerequisites

1. **Two SharePoint Tenants** (or two sites in the same tenant for testing)
   - Source tenant: `sourcetenant.onmicrosoft.com`
   - Target tenant: `targettenant.onmicrosoft.com`

2. **Test Users**
   - At least 5 test users in each tenant
   - Mix of internal and external (guest) users
   - Some users with identical emails, some different
   - At least one unmapped user (exists in source, not in target)

3. **Test Site**
   - Team site with content in source tenant
   - Multiple lists with items created by different users
   - Document library with files uploaded by different users
   - Site permissions with multiple groups and members
   - Custom list with People Picker columns

### Test Site Structure

```
Test Site: "UserMappingTestSite"
├── Site Permissions
│   ├── Owners: admin@source.com
│   ├── Members: user1@source.com, user2@source.com
│   └── Visitors: user3@source.com, external@partner.com
├── Lists
│   ├── "Tasks" (10 items, various authors/editors)
│   ├── "Project Tracker" (with AssignedTo people picker field)
│   └── "Announcements" (5 items)
└── Document Libraries
    ├── "Shared Documents" (20 files, multiple authors)
    └── "Project Files" (15 files)
```

## Test Scenarios

### Scenario 1: Basic User Extraction

**Objective**: Verify that New-UserMappingTemplate.ps1 correctly extracts all unique users from a site template.

**Steps**:
1. Export test site with content:
   ```powershell
   .\Export-SharePointSiteTemplate.ps1 `
       -SourceSiteUrl "https://sourcetenant.sharepoint.com/sites/UserMappingTestSite" `
       -IncludeContent `
       -TemplateName "UserMappingTest"
   ```

2. Generate user mapping template:
   ```powershell
   .\New-UserMappingTemplate.ps1 `
       -TemplatePath "C:\PSReports\SiteTemplates\UserMappingTest.pnp" `
       -OutputPath "test-user-mapping.csv"
   ```

3. Validate output:
   - Open `test-user-mapping.csv`
   - Verify all expected users are listed
   - Check SourceUser column contains valid emails
   - Verify SourceDisplayName is populated
   - Check Notes column shows where user was found

**Expected Results**:
- ✅ CSV file created successfully
- ✅ All site administrators extracted
- ✅ All site group members extracted
- ✅ List item authors/editors extracted
- ✅ Document authors/editors extracted
- ✅ People picker field users extracted
- ✅ No duplicate entries
- ✅ System accounts excluded by default

**Pass Criteria**: All unique users from site are in CSV with accurate source information.

---

### Scenario 2: Live Site Scanning

**Objective**: Verify that user extraction works from a live SharePoint site (not just template files).

**Steps**:
1. Generate mapping from live site:
   ```powershell
   .\New-UserMappingTemplate.ps1 `
       -SiteUrl "https://sourcetenant.sharepoint.com/sites/UserMappingTestSite" `
       -ConfigFile "app-config-source.json" `
       -OutputPath "live-site-users.csv"
   ```

2. Compare with template-based extraction:
   ```powershell
   $templateUsers = Import-Csv "test-user-mapping.csv"
   $liveUsers = Import-Csv "live-site-users.csv"
   Compare-Object $templateUsers.SourceUser $liveUsers.SourceUser
   ```

**Expected Results**:
- ✅ CSV file created successfully
- ✅ User count matches or exceeds template extraction
- ✅ Live scan may find additional users not in template export
- ✅ All critical users present

**Pass Criteria**: Live site scanning extracts at least the same users as template scanning.

---

### Scenario 3: User Mapping CSV Editing

**Objective**: Validate that manually edited CSV files are correctly processed.

**Steps**:
1. Copy sample CSV:
   ```powershell
   Copy-Item user-mapping.sample.csv test-scenario3.csv
   ```

2. Edit CSV with test cases:
   ```csv
   SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes
   user1@source.com,user1@target.com,User One,User One,Same person different tenant
   user2@source.com,user2alt@target.com,User Two,User Two (Alt),Email changed
   admin@source.com,newadmin@target.com,Source Admin,Target Admin,Role reassigned
   olduser@source.com,,Old User,,No longer exists - skip
   external@partner.com,external@partner.com,External User,External User,Guest user same email
   ```

3. Load mapping in PowerShell:
   ```powershell
   $mapping = Import-Csv test-scenario3.csv
   $mapping | Format-Table -AutoSize
   ```

**Expected Results**:
- ✅ CSV loads without errors
- ✅ Empty TargetUser cells handled correctly
- ✅ Display name differences preserved
- ✅ Notes column present (not required for processing)

**Pass Criteria**: CSV structure valid and ready for import.

---

### Scenario 4: Target User Validation (Pre-Flight Check)

**Objective**: Verify that target user validation correctly identifies valid and invalid users before import.

**Steps**:
1. Create mapping with intentionally invalid users:
   ```csv
   SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes
   user1@source.com,user1@target.com,User One,User One,Valid - exists in target
   user2@source.com,nonexistent@target.com,User Two,Ghost User,Invalid - does not exist
   user3@source.com,user3@target.com,User Three,User Three,Valid - exists in target
   ```

2. Run validation only:
   ```powershell
   .\Import-SharePointSiteTemplate.ps1 `
       -TargetSiteUrl "https://targettenant.sharepoint.com/sites/TestTarget" `
       -TemplatePath "UserMappingTest.pnp" `
       -UserMappingFile "test-scenario4.csv" `
       -ConfigFile "app-config-target.json" `
       -ValidateUsersOnly
   ```

3. Review validation output.

**Expected Results**:
- ✅ Script connects to target site
- ✅ Each target user checked
- ✅ Valid users marked with ✓
- ✅ Invalid users marked with ✗ and reason shown
- ✅ Summary shows count of valid vs invalid
- ✅ No import performed (validation only)
- ✅ Exit gracefully with actionable error message

**Pass Criteria**: Validation correctly identifies valid and invalid users; provides clear guidance to fix issues.

---

### Scenario 5: Successful Cross-Tenant Import with User Mapping

**Objective**: Perform a complete end-to-end cross-tenant migration with user mapping.

**Steps**:
1. Export from source tenant:
   ```powershell
   .\Export-SharePointSiteTemplate.ps1 `
       -SourceSiteUrl "https://sourcetenant.sharepoint.com/sites/UserMappingTestSite" `
       -ConfigFile "app-config-source.json" `
       -IncludeContent
   ```

2. Generate user mapping:
   ```powershell
   .\New-UserMappingTemplate.ps1 `
       -TemplatePath "C:\PSReports\SiteTemplates\UserMappingTestSite.pnp"
   ```

3. Edit `user-mapping-template.csv` with all valid target users.

4. Validate users:
   ```powershell
   .\Import-SharePointSiteTemplate.ps1 `
       -TargetSiteUrl "https://targettenant.sharepoint.com/sites/TestTarget" `
       -TemplatePath "C:\PSReports\SiteTemplates\UserMappingTestSite.pnp" `
       -UserMappingFile "user-mapping-template.csv" `
       -ConfigFile "app-config-target.json" `
       -ValidateUsersOnly
   ```

5. Perform import:
   ```powershell
   .\Import-SharePointSiteTemplate.ps1 `
       -TargetSiteUrl "https://targettenant.sharepoint.com/sites/TestTarget" `
       -TemplatePath "C:\PSReports\SiteTemplates\UserMappingTestSite.pnp" `
       -UserMappingFile "user-mapping-template.csv" `
       -ConfigFile "app-config-target.json" `
       -IgnoreDuplicateDataRowErrors
   ```

6. Verify target site:
   - Check site permissions
   - Check list item authors/editors
   - Check document authors/editors
   - Check people picker fields

**Expected Results**:
- ✅ Export completes successfully
- ✅ User mapping template generated
- ✅ Validation passes
- ✅ Import completes without critical errors
- ✅ Site structure imported
- ✅ Site permissions use target users
- ✅ List items show target users as authors/editors
- ✅ Documents show target users as authors/editors
- ✅ People picker fields reference target users
- ✅ Log file contains user mapping details

**Pass Criteria**: Complete migration with all user references correctly mapped to target tenant users.

---

### Scenario 6: Handling Unmapped Users

**Objective**: Verify behavior when some users are intentionally left unmapped.

**Steps**:
1. Create mapping with unmapped users:
   ```csv
   SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes
   user1@source.com,user1@target.com,User One,User One,Mapped
   user2@source.com,,User Two,,Intentionally unmapped
   user3@source.com,user3@target.com,User Three,User Three,Mapped
   ```

2. Import with this mapping:
   ```powershell
   .\Import-SharePointSiteTemplate.ps1 `
       -TargetSiteUrl "https://targettenant.sharepoint.com/sites/TestUnmapped" `
       -TemplatePath "UserMappingTest.pnp" `
       -UserMappingFile "test-scenario6.csv" `
       -ConfigFile "app-config-target.json" `
       -IgnoreDuplicateDataRowErrors
   ```

3. Check how unmapped users are handled in target site.

**Expected Results**:
- ✅ Script loads mapping and skips unmapped users
- ✅ Warning shown for skipped users
- ✅ Import proceeds
- ✅ Mapped users updated correctly
- ✅ Unmapped users remain as source tenant references (may show as deleted users in UI)
- ✅ No critical failures

**Pass Criteria**: Import completes; unmapped users handled gracefully without blocking import.

---

### Scenario 7: Large-Scale User Mapping

**Objective**: Test performance and reliability with a large number of users.

**Steps**:
1. Generate or create a CSV with 100+ user mappings.
2. Import a template with extensive content and user references.
3. Monitor:
   - Import duration
   - Memory usage
   - Errors or warnings
   - Completeness of mapping

**Expected Results**:
- ✅ Script handles large mapping file
- ✅ Import completes within reasonable time
- ✅ No memory issues
- ✅ All users mapped correctly
- ✅ Log file contains all mapping operations

**Pass Criteria**: Large-scale mapping completes successfully without performance degradation.

---

### Scenario 8: WhatIf Mode with User Mapping

**Objective**: Verify that -WhatIf mode previews user mapping operations without making changes.

**Steps**:
1. Run import with -WhatIf and user mapping:
   ```powershell
   .\Import-SharePointSiteTemplate.ps1 `
       -TargetSiteUrl "https://targettenant.sharepoint.com/sites/TestTarget" `
       -TemplatePath "UserMappingTest.pnp" `
       -UserMappingFile "user-mapping.csv" `
       -WhatIf
   ```

2. Check that no changes are made to target site.

**Expected Results**:
- ✅ Script loads user mapping
- ✅ User validation runs
- ✅ "WhatIf: Would apply template..." message shown
- ✅ "WhatIf: Would apply user mappings..." message shown
- ✅ No actual changes to target site
- ✅ No modified template files created

**Pass Criteria**: WhatIf mode shows intent without making changes.

---

### Scenario 9: Error Handling and Recovery

**Objective**: Test error handling for common failure scenarios.

**Test Cases**:

**9.1: Invalid CSV File**
- Missing required columns
- Malformed CSV (bad quotes, wrong delimiter)
- Empty file

**9.2: Missing Certificate/Auth Failure**
- Invalid configuration file
- Expired certificate
- Insufficient permissions

**9.3: Target Site Does Not Exist**
- Template valid, but target site URL is wrong

**9.4: User Mapping File Not Found**
- Specified file path does not exist

**9.5: Partial Import Failure**
- Some content fails to import
- Some users fail validation

**Expected Results**:
- ✅ Clear error messages for each scenario
- ✅ Helpful guidance on how to fix
- ✅ No data corruption
- ✅ Graceful exit
- ✅ Log files capture errors
- ✅ Cleanup of temporary files

**Pass Criteria**: Errors are caught and reported clearly; no undefined behavior.

---

## Validation Checklist

After each test scenario, verify:

### Site Structure
- [ ] Lists and libraries created
- [ ] List items imported
- [ ] Documents uploaded
- [ ] Site pages created
- [ ] Navigation configured

### User Mapping
- [ ] Site owners mapped correctly
- [ ] Site members mapped correctly
- [ ] Site visitors mapped correctly
- [ ] List item "Created By" mapped
- [ ] List item "Modified By" mapped
- [ ] Document "Author" mapped
- [ ] Document "Editor" mapped
- [ ] People Picker columns mapped
- [ ] Custom user fields mapped

### Permissions
- [ ] Site collection administrators correct
- [ ] Group memberships correct
- [ ] Item-level permissions preserved
- [ ] No broken permission inheritance

### Audit Trail
- [ ] Import log file created
- [ ] User mappings logged
- [ ] Timestamp and duration recorded
- [ ] Errors and warnings captured

### Cleanup
- [ ] Temporary user-mapped template removed
- [ ] No orphaned files left

## Automated Testing Script

```powershell
# Test-UserMapping.ps1 - Automated test runner
param(
    [string]$SourceTenant,
    [string]$TargetTenant,
    [string]$TestSiteName = "UserMappingTestSite"
)

$results = @()

# Test 1: User Extraction from Template
Write-Host "Running Test 1: User Extraction from Template..." -ForegroundColor Cyan
try {
    .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl "https://$SourceTenant.sharepoint.com/sites/$TestSiteName" -IncludeContent
    $exportResult = $?
    
    .\New-UserMappingTemplate.ps1 -TemplatePath (Get-ChildItem "C:\PSReports\SiteTemplates\" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName
    $extractResult = $?
    
    $results += [PSCustomObject]@{
        Test = "User Extraction from Template"
        Passed = ($exportResult -and $extractResult)
        Notes = "Check CSV output"
    }
}
catch {
    $results += [PSCustomObject]@{
        Test = "User Extraction from Template"
        Passed = $false
        Notes = $_.Exception.Message
    }
}

# Test 2: User Validation
Write-Host "Running Test 2: User Validation..." -ForegroundColor Cyan
try {
    .\Import-SharePointSiteTemplate.ps1 `
        -TargetSiteUrl "https://$TargetTenant.sharepoint.com/sites/TestTarget" `
        -TemplatePath (Get-ChildItem "C:\PSReports\SiteTemplates\" | Sort-Object LastWriteTime -Descending | Select-Object -First 1).FullName `
        -UserMappingFile "user-mapping-template.csv" `
        -ValidateUsersOnly
    
    $results += [PSCustomObject]@{
        Test = "User Validation"
        Passed = $?
        Notes = "Check validation output"
    }
}
catch {
    $results += [PSCustomObject]@{
        Test = "User Validation"
        Passed = $false
        Notes = $_.Exception.Message
    }
}

# Display results
Write-Host ""
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
Write-Host "  Test Results" -ForegroundColor Green
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
$results | Format-Table -AutoSize

$passedCount = ($results | Where-Object { $_.Passed -eq $true }).Count
$totalCount = $results.Count

Write-Host ""
Write-Host "Passed: $passedCount / $totalCount" -ForegroundColor $(if ($passedCount -eq $totalCount) { "Green" } else { "Yellow" })
```

## Performance Benchmarks

Target performance metrics:

| Operation | Small Site (<100 items) | Medium Site (100-1000 items) | Large Site (>1000 items) |
|-----------|-------------------------|------------------------------|--------------------------|
| User Extraction (Template) | < 5 seconds | < 15 seconds | < 60 seconds |
| User Extraction (Live) | < 30 seconds | < 2 minutes | < 5 minutes |
| User Validation | < 10 seconds | < 30 seconds | < 2 minutes |
| User Mapping Application | < 5 seconds | < 15 seconds | < 60 seconds |
| Full Import (with mapping) | < 2 minutes | < 10 minutes | < 30 minutes |

## Known Limitations

1. **System Accounts**: SharePoint system accounts (e.g., `sharepoint\system`) cannot be mapped
2. **External Users**: Guest users must exist in target tenant before import
3. **Deleted Users**: Users deleted from source tenant may not extract completely
4. **Claims Encoding**: Some claims-based identities may require special handling
5. **Workflow Assignments**: Workflow task assignments may need manual review post-migration
6. **Historical Data**: Audit logs and version history user references may not update

## Reporting Issues

When reporting issues with user mapping:
1. Include script version and PowerShell version
2. Attach sanitized CSV mapping file
3. Include relevant log file excerpts
4. Describe expected vs actual behavior
5. Note tenant types (commercial, GCC, etc.)

---

**Document Version**: 1.0  
**Last Updated**: February 4, 2026  
**Maintained By**: IT Support Team
