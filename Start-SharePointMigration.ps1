<#
.SYNOPSIS
    Interactive menu system for SharePoint site migrations.

.DESCRIPTION
    Provides a guided, menu-driven workflow for SharePoint migrations with:
    - Same-tenant or cross-tenant migration options
    - Prerequisites validation and guidance
    - Configuration file validation
    - Step-by-step workflow with all necessary commands

.PARAMETER ConfigFile
    Path to configuration file (optional, defaults to app-config.json).
    
.PARAMETER SourceConfigFile
    Path to source tenant configuration file for cross-tenant migrations.
    
.PARAMETER TargetConfigFile
    Path to target tenant configuration file for cross-tenant migrations.

.EXAMPLE
    .\Start-SharePointMigration.ps1
    
    Launches interactive menu to guide you through the migration process.

.NOTES
    Author: SharePoint Migration Toolset
    Version: 3.1
    Requires: PnP.PowerShell 2.12+
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigFile = "app-config.json",
    
    [Parameter(Mandatory = $false)]
    [string]$SourceConfigFile,
    
    [Parameter(Mandatory = $false)]
    [string]$TargetConfigFile
)

# Set strict mode
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helper Functions

function Write-Header {
    param([string]$Text)
    Write-Host "`n$('=' * 70)" -ForegroundColor Cyan
    Write-Host "  $Text" -ForegroundColor Cyan
    Write-Host "$('=' * 70)`n" -ForegroundColor Cyan
}

function Write-SubHeader {
    param([string]$Text)
    Write-Host "`n$Text" -ForegroundColor Yellow
    Write-Host "$('-' * $Text.Length)" -ForegroundColor Yellow
}

function Write-Step {
    param(
        [int]$StepNumber,
        [string]$Description
    )
    Write-Host "`n[$StepNumber] " -ForegroundColor Green -NoNewline
    Write-Host $Description -ForegroundColor White
}

function Write-InfoBox {
    param([string[]]$Lines)
    $maxLength = ($Lines | Measure-Object -Property Length -Maximum).Maximum
    $width = [Math]::Max($maxLength + 4, 50)
    
    Write-Host "`n╔$('═' * $width)╗" -ForegroundColor Magenta
    foreach ($line in $Lines) {
        $padding = ' ' * ($width - $line.Length - 2)
        Write-Host "║ $line$padding ║" -ForegroundColor Magenta
    }
    Write-Host "╚$('═' * $width)╝`n" -ForegroundColor Magenta
}

function Write-Prerequisite {
    param(
        [string]$Item,
        [bool]$IsMet = $false
    )
    $icon = if ($IsMet) { "✓" } else { "○" }
    $color = if ($IsMet) { "Green" } else { "Yellow" }
    Write-Host "  $icon " -ForegroundColor $color -NoNewline
    Write-Host $Item
}

function Show-Menu {
    param(
        [string]$Title,
        [string[]]$Options,
        [string]$Prompt = "Enter your choice"
    )
    
    Write-Header $Title
    
    for ($i = 0; $i -lt $Options.Count; $i++) {
        Write-Host "  [$($i + 1)] " -ForegroundColor Cyan -NoNewline
        Write-Host $Options[$i]
    }
    Write-Host "  [0] " -ForegroundColor Red -NoNewline
    Write-Host "Exit"
    
    Write-Host "`n$Prompt" -ForegroundColor Yellow -NoNewline
    Write-Host ": " -NoNewline
    $choice = Read-Host
    
    return $choice
}

function Test-Prerequisites {
    param(
        [switch]$CrossTenant
    )
    
    Write-SubHeader "Prerequisites Check"
    
    $allMet = $true
    
    # Check PnP.PowerShell module
    $pnpModule = Get-Module -ListAvailable -Name PnP.PowerShell | Select-Object -First 1
    $pnpInstalled = $null -ne $pnpModule
    Write-Prerequisite "PnP.PowerShell module installed (v2.12+)" -IsMet $pnpInstalled
    if ($pnpInstalled -and $pnpModule.Version) {
        Write-Host "    Version: $($pnpModule.Version)" -ForegroundColor Gray
    }
    $allMet = $allMet -and $pnpInstalled
    
    # Check configuration files
    if ($CrossTenant) {
        $sourceConfigExists = Test-Path $SourceConfigFile -ErrorAction SilentlyContinue
        $targetConfigExists = Test-Path $TargetConfigFile -ErrorAction SilentlyContinue
        
        Write-Prerequisite "Source config file: $SourceConfigFile" -IsMet $sourceConfigExists
        Write-Prerequisite "Target config file: $TargetConfigFile" -IsMet $targetConfigExists
        
        $allMet = $allMet -and $sourceConfigExists -and $targetConfigExists
    }
    else {
        $configExists = Test-Path $ConfigFile -ErrorAction SilentlyContinue
        Write-Prerequisite "Config file: $ConfigFile" -IsMet $configExists
        $allMet = $allMet -and $configExists
    }
    
    # Check Test-Configuration.ps1 exists
    $testScriptExists = Test-Path ".\Test-Configuration.ps1" -ErrorAction SilentlyContinue
    Write-Prerequisite "Test-Configuration.ps1 script available" -IsMet $testScriptExists
    $allMet = $allMet -and $testScriptExists
    
    return $allMet
}

function Show-SameTenantPrerequisites {
    Write-Header "Same-Tenant Migration Prerequisites"
    
    Write-InfoBox @(
        "Same-Tenant Migration Requirements:",
        "",
        "1. Azure AD App Registration (in your tenant)",
        "2. App permissions: SharePoint Sites.FullControl.All",
        "3. Certificate uploaded to app registration",
        "4. Admin consent granted",
        "5. Configuration file: app-config.json"
    )
    
    Write-Host "Configuration file should contain:" -ForegroundColor Cyan
    Write-Host @"
{
  "tenantId": "YOUR-TENANT-ID-GUID",
  "clientId": "YOUR-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "YOUR-CERTIFICATE-THUMBPRINT",
  "tenantDomain": "yourtenant.onmicrosoft.com"
}
"@ -ForegroundColor Gray
    
    Write-Host "`nFor detailed setup instructions:" -ForegroundColor Yellow
    Write-Host "  • See MANUAL-APP-REGISTRATION.md for app registration" -ForegroundColor White
    Write-Host "  • See CONFIG-README.md for configuration file setup`n" -ForegroundColor White
    
    # Check current prerequisites
    Test-Prerequisites | Out-Null
    
    Write-Host "`nReady to validate your configuration?" -ForegroundColor Yellow
    Write-Host "  [Y] " -ForegroundColor Green -NoNewline
    Write-Host "Yes - Run Test-Configuration.ps1"
    Write-Host "  [N] " -ForegroundColor Red -NoNewline
    Write-Host "No - Return to main menu"
    
    $response = Read-Host "`nYour choice"
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-Host "`nRunning configuration validation..." -ForegroundColor Cyan
        if (Test-Path ".\Test-Configuration.ps1") {
            & ".\Test-Configuration.ps1" -ConfigFile $ConfigFile
        }
        else {
            Write-Host "ERROR: Test-Configuration.ps1 not found in current directory" -ForegroundColor Red
        }
        
        Write-Host "`nPress any key to continue..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

function Show-CrossTenantPrerequisites {
    Write-Header "Cross-Tenant Migration Prerequisites"
    
    Write-InfoBox @(
        "Cross-Tenant Migration Requirements:",
        "",
        "1. TWO Azure AD App Registrations:",
        "   - One in SOURCE tenant",
        "   - One in TARGET tenant",
        "2. Each app needs: SharePoint Sites.FullControl.All",
        "3. Certificate uploaded to BOTH app registrations",
        "4. Admin consent granted in BOTH tenants",
        "5. Two configuration files:",
        "   - app-config-source.json (source tenant)",
        "   - app-config-target.json (target tenant)"
    )
    
    Write-Host "Source configuration file (app-config-source.json):" -ForegroundColor Cyan
    Write-Host @"
{
  "tenantId": "SOURCE-TENANT-ID-GUID",
  "clientId": "SOURCE-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "CERT-THUMBPRINT",
  "tenantDomain": "sourcetenant.onmicrosoft.com"
}
"@ -ForegroundColor Gray
    
    Write-Host "`nTarget configuration file (app-config-target.json):" -ForegroundColor Cyan
    Write-Host @"
{
  "tenantId": "TARGET-TENANT-ID-GUID",
  "clientId": "TARGET-APP-CLIENT-ID-GUID",
  "certificateThumbprint": "CERT-THUMBPRINT",
  "tenantDomain": "targettenant.onmicrosoft.com"
}
"@ -ForegroundColor Gray
    
    Write-Host "`nFor detailed setup instructions:" -ForegroundColor Yellow
    Write-Host "  • See MANUAL-APP-REGISTRATION.md for app registration in BOTH tenants" -ForegroundColor White
    Write-Host "  • See CONFIG-README.md for configuration file setup" -ForegroundColor White
    Write-Host "  • You can use the same certificate in both tenants`n" -ForegroundColor White
    
    # Check current prerequisites
    $SourceConfigFile = if ([string]::IsNullOrWhiteSpace($script:SourceConfigFile)) { "app-config-source.json" } else { $script:SourceConfigFile }
    $TargetConfigFile = if ([string]::IsNullOrWhiteSpace($script:TargetConfigFile)) { "app-config-target.json" } else { $script:TargetConfigFile }
    
    Test-Prerequisites -CrossTenant | Out-Null
    
    Write-Host "`nReady to validate your configurations?" -ForegroundColor Yellow
    Write-Host "  [Y] " -ForegroundColor Green -NoNewline
    Write-Host "Yes - Run Test-Configuration.ps1 for both configs"
    Write-Host "  [N] " -ForegroundColor Red -NoNewline
    Write-Host "No - Return to main menu"
    
    $response = Read-Host "`nYour choice"
    
    if ($response -eq 'Y' -or $response -eq 'y') {
        Write-Host "`nRunning cross-tenant configuration validation..." -ForegroundColor Cyan
        if (Test-Path ".\Test-Configuration.ps1") {
            & ".\Test-Configuration.ps1" -SourceConfigFile $SourceConfigFile -TargetConfigFile $TargetConfigFile
        }
        else {
            Write-Host "ERROR: Test-Configuration.ps1 not found in current directory" -ForegroundColor Red
        }
        
        Write-Host "`nPress any key to continue..." -ForegroundColor Gray
        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    }
}

function Show-SameTenantWorkflow {
    Write-Header "Same-Tenant Migration Workflow"
    
    Write-InfoBox @(
        "This workflow will guide you through migrating a SharePoint site",
        "within the same Microsoft 365 tenant."
    )
    
    Write-Step 1 "Export the source site"
    Write-Host "    Run:" -ForegroundColor Gray
    Write-Host @"
    .\Export-SharePointSiteTemplate.ps1 ``
        -SourceSiteUrl "https://yourtenant.sharepoint.com/sites/SourceSite" ``
        -IncludeContent
"@ -ForegroundColor White
    
    Write-Step 2 "Inspect the exported template (optional)"
    Write-Host "    Run:" -ForegroundColor Gray
    Write-Host @"
    .\Get-TemplateContent.ps1 ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -Detailed -ShowUsers -ShowContent
"@ -ForegroundColor White
    
    Write-Step 3 "Create the target site"
    Write-Host "    Options:" -ForegroundColor Gray
    Write-Host "      a) Via SharePoint Admin Center:" -ForegroundColor White
    Write-Host "         https://yourtenant-admin.sharepoint.com" -ForegroundColor Gray
    Write-Host "      b) Via PowerShell:" -ForegroundColor White
    Write-Host @"
         Connect-PnPOnline -Url "https://yourtenant-admin.sharepoint.com" -Interactive
         New-PnPSite -Type TeamSite -Title "Target Site" -Alias "TargetSite" -Wait
"@ -ForegroundColor Gray
    
    Write-Step 4 "Import to the target site"
    Write-Host "    Run:" -ForegroundColor Gray
    Write-Host @"
    .\Import-SharePointSiteTemplate.ps1 ``
        -TargetSiteUrl "https://yourtenant.sharepoint.com/sites/TargetSite" ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp"
"@ -ForegroundColor White
    
    Write-Host "`n" + "─" * 70 -ForegroundColor Cyan
    Write-Host "TIP: Use -Preview on export and -InspectOnly on import to see" -ForegroundColor Yellow
    Write-Host "     what will be included before running the full migration." -ForegroundColor Yellow
    Write-Host "─" * 70 + "`n" -ForegroundColor Cyan
    
    Write-Host "Press any key to return to main menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

function Show-CrossTenantWorkflow {
    Write-Header "Cross-Tenant Migration Workflow"
    
    Write-InfoBox @(
        "This workflow will guide you through migrating a SharePoint site",
        "from one Microsoft 365 tenant to another tenant."
    )
    
    Write-Step 1 "Export from source tenant"
    Write-Host "    Run:" -ForegroundColor Gray
    Write-Host @"
    .\Export-SharePointSiteTemplate.ps1 ``
        -SourceSiteUrl "https://sourcetenant.sharepoint.com/sites/Site" ``
        -ConfigFile "app-config-source.json" ``
        -IncludeContent
"@ -ForegroundColor White
    
    Write-Step 2 "Generate user mapping template"
    Write-Host "    Run:" -ForegroundColor Gray
    Write-Host @"
    .\New-UserMappingTemplate.ps1 ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -OutputPath "user-mapping.csv"
"@ -ForegroundColor White
    
    Write-Step 3 "Edit user mapping CSV"
    Write-Host "    Actions:" -ForegroundColor Gray
    Write-Host "      • Open user-mapping.csv in Excel or text editor" -ForegroundColor White
    Write-Host "      • Update TargetUser column with target tenant emails" -ForegroundColor White
    Write-Host "      • Update TargetDisplayName if names differ" -ForegroundColor White
    Write-Host "      • Leave TargetUser empty to skip unmapped users" -ForegroundColor White
    Write-Host "`n    Example:" -ForegroundColor Gray
    Write-Host "      john@source.com → john@target.com" -ForegroundColor White
    Write-Host "      admin@source.com → it.admin@target.com" -ForegroundColor White
    
    Write-Step 4 "Validate target users"
    Write-Host "    Run:" -ForegroundColor Gray
    Write-Host @"
    .\Import-SharePointSiteTemplate.ps1 ``
        -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -UserMappingFile "user-mapping.csv" ``
        -ConfigFile "app-config-target.json" ``
        -ValidateUsersOnly
"@ -ForegroundColor White
    
    Write-Host "`n    Fix any validation errors in user-mapping.csv and re-validate" -ForegroundColor Yellow
    
    Write-Step 5 "Create target site in target tenant"
    Write-Host "    Options:" -ForegroundColor Gray
    Write-Host "      a) Via Target Tenant Admin Center:" -ForegroundColor White
    Write-Host "         https://targettenant-admin.sharepoint.com" -ForegroundColor Gray
    Write-Host "      b) Via PowerShell:" -ForegroundColor White
    Write-Host @"
         Connect-PnPOnline -Url "https://targettenant-admin.sharepoint.com" ``
             -ClientId "TARGET-APP-ID" ``
             -Thumbprint "CERT-THUMBPRINT" ``
             -Tenant "targettenant.onmicrosoft.com"
         New-PnPSite -Type TeamSite -Title "Target Site" -Alias "TargetSite" -Wait
"@ -ForegroundColor Gray
    
    Write-Step 6 "Import to target tenant with user mapping"
    Write-Host "    Run:" -ForegroundColor Gray
    Write-Host @"
    .\Import-SharePointSiteTemplate.ps1 ``
        -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -UserMappingFile "user-mapping.csv" ``
        -ConfigFile "app-config-target.json" ``
        -IgnoreDuplicateDataRowErrors
"@ -ForegroundColor White
    
    Write-Host "`n" + "─" * 70 -ForegroundColor Cyan
    Write-Host "TIP: Always validate users before the full import to catch" -ForegroundColor Yellow
    Write-Host "     missing or invalid target users early in the process." -ForegroundColor Yellow
    Write-Host "─" * 70 + "`n" -ForegroundColor Cyan
    
    Write-Host "Press any key to return to main menu..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

#endregion

#region Main Menu

function Show-MainMenu {
    Clear-Host
    
    Write-Host @"

   _____ _                 _____       _       _     __  __ _                 _   _             
  / ____| |               |  __ \     (_)     | |   |  \/  (_)               | | (_)            
 | (___ | |__   __ _ _ __ | |__) |__   _ _ __ | |_  | \  / |_  __ _ _ __ __ _| |_ _  ___  _ __  
  \___ \| '_ \ / _' | '__||  ___/ _ \ | | '_ \| __| | |\/| | |/ _' | '__/ _' | __| |/ _ \| '_ \ 
  ____) | | | | (_| | |   | |  | (_) || | | | | |_  | |  | | | (_| | | | (_| | |_| | (_) | | | |
 |_____/|_| |_|\__,_|_|   |_|   \___/ |_|_| |_|\__| |_|  |_|_|\__, |_|  \__,_|\__|_|\___/|_| |_|
                                                                __/ |                            
                                                               |___/                             
"@ -ForegroundColor Cyan
    
    Write-Host "                        Interactive Migration Assistant v3.1" -ForegroundColor Gray
    Write-Host ""
    
    $choice = Show-Menu -Title "SharePoint Migration - Main Menu" `
        -Options @(
            "Same-Tenant Migration (within one Microsoft 365 tenant)",
            "Cross-Tenant Migration (between different Microsoft 365 tenants)",
            "View Documentation",
            "Test Configuration Files"
        ) `
        -Prompt "Select migration type"
    
    return $choice
}

function Show-DocumentationMenu {
    $choice = Show-Menu -Title "Documentation" `
        -Options @(
            "README.md - Main documentation and workflows",
            "MANUAL-APP-REGISTRATION.md - Azure AD app setup guide",
            "CONFIG-README.md - Configuration file guidance",
            "USER-MAPPING-QUICK-REF.md - User mapping quick reference",
            "DEVELOPER.md - Contribution guidelines"
        ) `
        -Prompt "Select document to view"
    
    switch ($choice) {
        "1" { if (Test-Path ".\README.md") { & notepad.exe ".\README.md" } }
        "2" { if (Test-Path ".\MANUAL-APP-REGISTRATION.md") { & notepad.exe ".\MANUAL-APP-REGISTRATION.md" } }
        "3" { if (Test-Path ".\CONFIG-README.md") { & notepad.exe ".\CONFIG-README.md" } }
        "4" { if (Test-Path ".\USER-MAPPING-QUICK-REF.md") { & notepad.exe ".\USER-MAPPING-QUICK-REF.md" } }
        "5" { if (Test-Path ".\DEVELOPER.md") { & notepad.exe ".\DEVELOPER.md" } }
        "0" { return }
        default { 
            Write-Host "`nInvalid choice. Press any key to continue..." -ForegroundColor Red
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
}

function Show-TestConfigMenu {
    $choice = Show-Menu -Title "Test Configuration Files" `
        -Options @(
            "Test single configuration file (same-tenant)",
            "Test cross-tenant configurations (source + target)"
        ) `
        -Prompt "Select test mode"
    
    switch ($choice) {
        "1" {
            Write-Host "`nEnter path to configuration file (or press Enter for 'app-config.json'): " -ForegroundColor Yellow -NoNewline
            $configPath = Read-Host
            if ([string]::IsNullOrWhiteSpace($configPath)) {
                $configPath = "app-config.json"
            }
            
            if (Test-Path ".\Test-Configuration.ps1") {
                Write-Host "`nRunning validation..." -ForegroundColor Cyan
                & ".\Test-Configuration.ps1" -ConfigFile $configPath
            }
            else {
                Write-Host "`nERROR: Test-Configuration.ps1 not found" -ForegroundColor Red
            }
            
            Write-Host "`nPress any key to continue..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "2" {
            Write-Host "`nEnter path to SOURCE configuration file (or press Enter for 'app-config-source.json'): " -ForegroundColor Yellow -NoNewline
            $sourceConfig = Read-Host
            if ([string]::IsNullOrWhiteSpace($sourceConfig)) {
                $sourceConfig = "app-config-source.json"
            }
            
            Write-Host "Enter path to TARGET configuration file (or press Enter for 'app-config-target.json'): " -ForegroundColor Yellow -NoNewline
            $targetConfig = Read-Host
            if ([string]::IsNullOrWhiteSpace($targetConfig)) {
                $targetConfig = "app-config-target.json"
            }
            
            if (Test-Path ".\Test-Configuration.ps1") {
                Write-Host "`nRunning cross-tenant validation..." -ForegroundColor Cyan
                & ".\Test-Configuration.ps1" -SourceConfigFile $sourceConfig -TargetConfigFile $targetConfig
            }
            else {
                Write-Host "`nERROR: Test-Configuration.ps1 not found" -ForegroundColor Red
            }
            
            Write-Host "`nPress any key to continue..." -ForegroundColor Gray
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
        "0" { return }
        default {
            Write-Host "`nInvalid choice. Press any key to continue..." -ForegroundColor Red
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        }
    }
}

#endregion

#region Main Program Loop

try {
    # Set default config files if not provided
    if ([string]::IsNullOrWhiteSpace($SourceConfigFile)) {
        $SourceConfigFile = "app-config-source.json"
    }
    if ([string]::IsNullOrWhiteSpace($TargetConfigFile)) {
        $TargetConfigFile = "app-config-target.json"
    }
    
    # Main loop
    $running = $true
    while ($running) {
        $mainChoice = Show-MainMenu
        
        switch ($mainChoice) {
            "1" {
                # Same-Tenant Migration
                Show-SameTenantPrerequisites
                $continue = Read-Host "`nContinue to workflow guide? (Y/N)"
                if ($continue -eq 'Y' -or $continue -eq 'y') {
                    Show-SameTenantWorkflow
                }
            }
            "2" {
                # Cross-Tenant Migration
                Show-CrossTenantPrerequisites
                $continue = Read-Host "`nContinue to workflow guide? (Y/N)"
                if ($continue -eq 'Y' -or $continue -eq 'y') {
                    Show-CrossTenantWorkflow
                }
            }
            "3" {
                # Documentation
                Show-DocumentationMenu
            }
            "4" {
                # Test Configuration
                Show-TestConfigMenu
            }
            "0" {
                # Exit
                Write-Host "`nThank you for using SharePoint Migration Assistant!" -ForegroundColor Cyan
                $running = $false
            }
            default {
                Write-Host "`nInvalid choice. Press any key to continue..." -ForegroundColor Red
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
        }
    }
}
catch {
    Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
}

#endregion
