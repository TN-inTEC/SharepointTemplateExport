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
    
    Write-Host "`nFor detailed setup instructions:" -ForegroundColor Yellow
    Write-Host "  • See MANUAL-APP-REGISTRATION.md for app registration in BOTH tenants" -ForegroundColor White
    Write-Host "  • See CONFIG-README.md for configuration file setup and examples" -ForegroundColor White
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
    
    $workflowRunning = $true
    $exportedTemplate = $null
    
    while ($workflowRunning) {
        $choice = Show-Menu -Title "Same-Tenant Workflow Steps" `
            -Options @(
                "Export source site to template",
                "Inspect exported template",
                "Import template to target site",
                "View all commands (reference)",
                "Return to main menu"
            ) `
            -Prompt "Select step to execute"
        
        switch ($choice) {
            "1" {
                # Export source site
                Write-SubHeader "Step 1: Export Source Site"
                
                Write-Host "`nEnter source site URL: " -ForegroundColor Yellow -NoNewline
                $sourceUrl = Read-Host
                
                if ([string]::IsNullOrWhiteSpace($sourceUrl)) {
                    Write-Host "Source URL is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nInclude content (lists/libraries data)? [Y/N] (default: Y): " -ForegroundColor Yellow -NoNewline
                $includeContent = Read-Host
                $includeContentSwitch = if ($includeContent -eq 'N' -or $includeContent -eq 'n') { "" } else { "-IncludeContent" }
                
                Write-Host "`nRun in PREVIEW mode (no export, just show what would be included)? [Y/N] (default: N): " -ForegroundColor Yellow -NoNewline
                $preview = Read-Host
                $previewSwitch = if ($preview -eq 'Y' -or $preview -eq 'y') { "-Preview" } else { "" }
                
                if ($previewSwitch) {
                    Write-Host "`nExecuting preview (no files will be created)..." -ForegroundColor Cyan
                }
                else {
                    Write-Host "`nExecuting export..." -ForegroundColor Cyan
                }
                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                Write-Host ".\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl `"$sourceUrl`" -ConfigFile `"$ConfigFile`" $includeContentSwitch $previewSwitch" -ForegroundColor White
                
                if (Test-Path ".\Export-SharePointSiteTemplate.ps1") {
                    try {
                        if ($previewSwitch -and $includeContentSwitch) {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $ConfigFile -IncludeContent -Preview
                        }
                        elseif ($previewSwitch) {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $ConfigFile -Preview
                        }
                        elseif ($includeContentSwitch) {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $ConfigFile -IncludeContent
                        }
                        else {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $ConfigFile
                        }
                        
                        # Try to find the most recent template (only if not preview mode)
                        if (-not $previewSwitch) {
                            $templates = Get-ChildItem "C:\PSReports\SiteTemplates\*.pnp" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
                            if ($templates) {
                                $exportedTemplate = $templates[0].FullName
                                Write-Host "`n✓ Export completed!" -ForegroundColor Green
                                Write-Host "Template saved: $exportedTemplate" -ForegroundColor Cyan
                            }
                        }
                        else {
                            Write-Host "`n✓ Preview completed (no files created)" -ForegroundColor Green
                        }
                    }
                    catch {
                        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nERROR: Export-SharePointSiteTemplate.ps1 not found" -ForegroundColor Red
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "2" {
                # Inspect template
                Write-SubHeader "Step 2: Inspect Exported Template"
                
                if ($exportedTemplate) {
                    Write-Host "`nMost recent template: " -ForegroundColor Cyan
                    Write-Host $exportedTemplate -ForegroundColor White
                    Write-Host "`nUse this template? [Y/N] (default: Y): " -ForegroundColor Yellow -NoNewline
                    $useRecent = Read-Host
                    
                    if ($useRecent -eq 'N' -or $useRecent -eq 'n') {
                        Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                        $templatePath = Read-Host
                    }
                    else {
                        $templatePath = $exportedTemplate
                    }
                }
                else {
                    Write-Host "`nEnter template path: " -ForegroundColor Yellow -NoNewline
                    $templatePath = Read-Host
                }
                
                if ([string]::IsNullOrWhiteSpace($templatePath)) {
                    Write-Host "Template path is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nExecuting inspection..." -ForegroundColor Cyan
                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                Write-Host ".\Get-TemplateContent.ps1 -TemplatePath `"$templatePath`" -Detailed -ShowUsers -ShowContent" -ForegroundColor White
                
                if (Test-Path ".\Get-TemplateContent.ps1") {
                    try {
                        & ".\Get-TemplateContent.ps1" -TemplatePath $templatePath -Detailed -ShowUsers -ShowContent
                    }
                    catch {
                        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nERROR: Get-TemplateContent.ps1 not found" -ForegroundColor Red
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "3" {
                # Import to target
                Write-SubHeader "Step 3: Import to Target Site"
                
                Write-Host "`n⚠️  PREREQUISITE: Target site must already exist!" -ForegroundColor Yellow
                Write-Host "Create via Admin Center or PowerShell before proceeding.`n" -ForegroundColor Yellow
                
                Write-Host "Enter target site URL: " -ForegroundColor Yellow -NoNewline
                $targetUrl = Read-Host
                
                if ([string]::IsNullOrWhiteSpace($targetUrl)) {
                    Write-Host "Target URL is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                if ($exportedTemplate) {
                    Write-Host "`nMost recent template: " -ForegroundColor Cyan
                    Write-Host $exportedTemplate -ForegroundColor White
                    Write-Host "`nUse this template? [Y/N] (default: Y): " -ForegroundColor Yellow -NoNewline
                    $useRecent = Read-Host
                    
                    if ($useRecent -eq 'N' -or $useRecent -eq 'n') {
                        Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                        $templatePath = Read-Host
                    }
                    else {
                        $templatePath = $exportedTemplate
                    }
                }
                else {
                    Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                    $templatePath = Read-Host
                }
                
                if ([string]::IsNullOrWhiteSpace($templatePath)) {
                    Write-Host "Template path is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nRun in INSPECT-ONLY mode (validation only, no import)? [Y/N] (default: N): " -ForegroundColor Yellow -NoNewline
                $inspectOnly = Read-Host
                $inspectSwitch = if ($inspectOnly -eq 'Y' -or $inspectOnly -eq 'y') { "-InspectOnly" } else { "" }
                
                if (-not $inspectSwitch) {
                    Write-Host "`n⚠️  FINAL CONFIRMATION" -ForegroundColor Yellow
                    Write-Host "About to import to: $targetUrl" -ForegroundColor White
                    Write-Host "From template: $templatePath" -ForegroundColor White
                    Write-Host "`nThis will modify the target site!" -ForegroundColor Yellow
                    Write-Host "Proceed with import? [Y/N]: " -ForegroundColor Yellow -NoNewline
                    $confirm = Read-Host
                    
                    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
                        Write-Host "Import cancelled." -ForegroundColor Cyan
                        Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                        $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                        continue
                    }
                }
                
                if ($inspectSwitch) {
                    Write-Host "`nExecuting inspection (no changes will be made)..." -ForegroundColor Cyan
                }
                else {
                    Write-Host "`nExecuting import..." -ForegroundColor Cyan
                }
                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                Write-Host ".\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl `"$targetUrl`" -TemplatePath `"$templatePath`" -ConfigFile `"$ConfigFile`" $inspectSwitch" -ForegroundColor White
                
                if (Test-Path ".\Import-SharePointSiteTemplate.ps1") {
                    try {
                        if ($inspectSwitch) {
                            & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $targetUrl -TemplatePath $templatePath -ConfigFile $ConfigFile -InspectOnly
                            Write-Host "`n✓ Inspection completed (no changes made)" -ForegroundColor Green
                        }
                        else {
                            & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $targetUrl -TemplatePath $templatePath -ConfigFile $ConfigFile
                            Write-Host "`n✓ Import completed!" -ForegroundColor Green
                        }
                    }
                    catch {
                        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nERROR: Import-SharePointSiteTemplate.ps1 not found" -ForegroundColor Red
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "4" {
                # View reference commands
                Write-SubHeader "Complete Workflow Reference"
                
                Write-Step 1 "Export the source site"
                Write-Host @"
    .\Export-SharePointSiteTemplate.ps1 ``
        -SourceSiteUrl "https://yourtenant.sharepoint.com/sites/SourceSite" ``
        -ConfigFile "$ConfigFile" ``
        -IncludeContent
"@ -ForegroundColor White
                
                Write-Step 2 "Inspect the exported template (optional)"
                Write-Host @"
    .\Get-TemplateContent.ps1 ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -Detailed -ShowUsers -ShowContent
"@ -ForegroundColor White
                
                Write-Step 3 "Import to the target site"
                Write-Host @"
    .\Import-SharePointSiteTemplate.ps1 ``
        -TargetSiteUrl "https://yourtenant.sharepoint.com/sites/TargetSite" ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -ConfigFile "$ConfigFile"
"@ -ForegroundColor White
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "5" {
                # Return to main menu
                $workflowRunning = $false
            }
            "0" {
                $workflowRunning = $false
            }
            default {
                Write-Host "`nInvalid choice. Press any key to continue..." -ForegroundColor Red
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
        }
    }
}

function Show-CrossTenantWorkflow {
    Write-Header "Cross-Tenant Migration Workflow"
    
    Write-InfoBox @(
        "This workflow will guide you through migrating a SharePoint site",
        "from one Microsoft 365 tenant to another tenant."
    )
    
    # Load config files to get tenant domains
    $sourceConfig = $null
    $targetConfig = $null
    $sourceDomain = ""
    $targetDomain = ""
    
    if (Test-Path $SourceConfigFile) {
        try {
            $sourceConfig = Get-Content $SourceConfigFile -Raw | ConvertFrom-Json
            $sourceDomain = $sourceConfig.tenantDomain
        }
        catch {
            Write-Host "Warning: Could not read source config file" -ForegroundColor Yellow
        }
    }
    
    if (Test-Path $TargetConfigFile) {
        try {
            $targetConfig = Get-Content $TargetConfigFile -Raw | ConvertFrom-Json
            $targetDomain = $targetConfig.tenantDomain
        }
        catch {
            Write-Host "Warning: Could not read target config file" -ForegroundColor Yellow
        }
    }
    
    $workflowRunning = $true
    $exportedTemplate = $null
    $userMappingFile = "user-mapping.csv"
    
    while ($workflowRunning) {
        $choice = Show-Menu -Title "Cross-Tenant Workflow Steps" `
            -Options @(
                "Export from source tenant",
                "Generate user mapping template",
                "Edit user mapping CSV",
                "Validate target users",
                "Import to target tenant with user mapping",
                "View all commands (reference)",
                "Return to main menu"
            ) `
            -Prompt "Select step to execute"
        
        switch ($choice) {
            "1" {
                # Export from source
                Write-SubHeader "Step 1: Export from Source Tenant"
                
                if ($sourceDomain) {
                    Write-Host "`nSource tenant: " -ForegroundColor Cyan -NoNewline
                    Write-Host $sourceDomain -ForegroundColor White
                }
                
                Write-Host "`nEnter source site URL: " -ForegroundColor Yellow -NoNewline
                $sourceUrl = Read-Host
                
                if ([string]::IsNullOrWhiteSpace($sourceUrl)) {
                    Write-Host "Source URL is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nInclude content? [Y/N] (default: Y): " -ForegroundColor Yellow -NoNewline
                $includeContent = Read-Host
                $includeContentSwitch = if ($includeContent -eq 'N' -or $includeContent -eq 'n') { "" } else { "-IncludeContent" }
                
                Write-Host "`nRun in PREVIEW mode (no export, just show what would be included)? [Y/N] (default: N): " -ForegroundColor Yellow -NoNewline
                $preview = Read-Host
                $previewSwitch = if ($preview -eq 'Y' -or $preview -eq 'y') { "-Preview" } else { "" }
                
                if ($previewSwitch) {
                    Write-Host "`nExecuting preview (no files will be created)..." -ForegroundColor Cyan
                }
                else {
                    Write-Host "`nExecuting export..." -ForegroundColor Cyan
                }
                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                Write-Host ".\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl `"$sourceUrl`" -ConfigFile `"$SourceConfigFile`" $includeContentSwitch $previewSwitch" -ForegroundColor White
                
                if (Test-Path ".\Export-SharePointSiteTemplate.ps1") {
                    try {
                        if ($previewSwitch -and $includeContentSwitch) {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $SourceConfigFile -IncludeContent -Preview
                        }
                        elseif ($previewSwitch) {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $SourceConfigFile -Preview
                        }
                        elseif ($includeContentSwitch) {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $SourceConfigFile -IncludeContent
                        }
                        else {
                            & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $sourceUrl -ConfigFile $SourceConfigFile
                        }
                        
                        # Try to find the most recent template (only if not preview mode)
                        if (-not $previewSwitch) {
                            $templates = Get-ChildItem "C:\PSReports\SiteTemplates\*.pnp" -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
                            if ($templates) {
                                $exportedTemplate = $templates[0].FullName
                                Write-Host "`n✓ Export completed!" -ForegroundColor Green
                                Write-Host "Template saved: $exportedTemplate" -ForegroundColor Cyan
                            }
                        }
                        else {
                            Write-Host "`n✓ Preview completed (no files created)" -ForegroundColor Green
                        }
                    }
                    catch {
                        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nERROR: Export-SharePointSiteTemplate.ps1 not found" -ForegroundColor Red
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "2" {
                # Generate user mapping
                Write-SubHeader "Step 2: Generate User Mapping Template"
                
                if ($exportedTemplate) {
                    Write-Host "`nMost recent template: " -ForegroundColor Cyan
                    Write-Host $exportedTemplate -ForegroundColor White
                    Write-Host "`nUse this template? [Y/N] (default: Y): " -ForegroundColor Yellow -NoNewline
                    $useRecent = Read-Host
                    
                    if ($useRecent -eq 'N' -or $useRecent -eq 'n') {
                        Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                        $templatePath = Read-Host
                    }
                    else {
                        $templatePath = $exportedTemplate
                    }
                }
                else {
                    Write-Host "`nEnter template path: " -ForegroundColor Yellow -NoNewline
                    $templatePath = Read-Host
                }
                
                if ([string]::IsNullOrWhiteSpace($templatePath)) {
                    Write-Host "Template path is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nOutput CSV path (default: user-mapping.csv): " -ForegroundColor Yellow -NoNewline
                $outputPath = Read-Host
                if ([string]::IsNullOrWhiteSpace($outputPath)) {
                    $outputPath = "user-mapping.csv"
                }
                $userMappingFile = $outputPath
                
                Write-Host "`nExecuting user mapping generation..." -ForegroundColor Cyan
                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                Write-Host ".\New-UserMappingTemplate.ps1 -TemplatePath `"$templatePath`" -OutputPath `"$outputPath`"" -ForegroundColor White
                
                if (Test-Path ".\New-UserMappingTemplate.ps1") {
                    try {
                        & ".\New-UserMappingTemplate.ps1" -TemplatePath $templatePath -OutputPath $outputPath
                        Write-Host "`n✓ User mapping template created!" -ForegroundColor Green
                        Write-Host "File: $outputPath" -ForegroundColor Cyan
                        
                        if ($sourceDomain -and $targetDomain) {
                            Write-Host "`nℹ️  TIP: Update TargetUser column in CSV:" -ForegroundColor Yellow
                            Write-Host "   Replace '@$sourceDomain' with '@$targetDomain'" -ForegroundColor White
                        }
                    }
                    catch {
                        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nERROR: New-UserMappingTemplate.ps1 not found" -ForegroundColor Red
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "3" {
                # Edit user mapping
                Write-SubHeader "Step 3: Edit User Mapping CSV"
                
                Write-Host "`nUser mapping file (default: $userMappingFile): " -ForegroundColor Yellow -NoNewline
                $mappingPath = Read-Host
                if ([string]::IsNullOrWhiteSpace($mappingPath)) {
                    $mappingPath = $userMappingFile
                }
                
                if (Test-Path $mappingPath) {
                    Write-Host "`nOpening $mappingPath in default editor..." -ForegroundColor Cyan
                    
                    if ($sourceDomain -and $targetDomain) {
                        Write-Host "`nREMINDER: Update TargetUser column:" -ForegroundColor Yellow
                        Write-Host "  Source domain: @$sourceDomain" -ForegroundColor White
                        Write-Host "  Target domain: @$targetDomain" -ForegroundColor White
                        Write-Host "`nExample mappings:" -ForegroundColor Yellow
                        Write-Host "  john.smith@$sourceDomain → john.smith@$targetDomain" -ForegroundColor White
                        Write-Host "  admin@$sourceDomain → it.admin@$targetDomain" -ForegroundColor White
                    }
                    
                    try {
                        Start-Process $mappingPath
                        Write-Host "`n✓ File opened in default application" -ForegroundColor Green
                        Write-Host "Edit the file and save when complete." -ForegroundColor Cyan
                    }
                    catch {
                        Write-Host "`nERROR: Could not open file: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "Please open manually: $mappingPath" -ForegroundColor Yellow
                    }
                }
                else {
                    Write-Host "`nERROR: File not found: $mappingPath" -ForegroundColor Red
                    Write-Host "Generate the user mapping template first (option 2)" -ForegroundColor Yellow
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "4" {
                # Validate users
                Write-SubHeader "Step 4: Validate Target Users"
                
                if ($targetDomain) {
                    Write-Host "`nTarget tenant: " -ForegroundColor Cyan -NoNewline
                    Write-Host $targetDomain -ForegroundColor White
                }
                
                Write-Host "`nEnter target site URL: " -ForegroundColor Yellow -NoNewline
                $targetUrl = Read-Host
                
                if ([string]::IsNullOrWhiteSpace($targetUrl)) {
                    Write-Host "Target URL is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                if ($exportedTemplate) {
                    Write-Host "`nMost recent template: " -ForegroundColor Cyan
                    Write-Host $exportedTemplate -ForegroundColor White
                    Write-Host "`nUse this template? [Y/N] (default: Y): " -ForegroundColor Yellow -NoNewline
                    $useRecent = Read-Host
                    
                    if ($useRecent -eq 'N' -or $useRecent -eq 'n') {
                        Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                        $templatePath = Read-Host
                    }
                    else {
                        $templatePath = $exportedTemplate
                    }
                }
                else {
                    Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                    $templatePath = Read-Host
                }
                
                if ([string]::IsNullOrWhiteSpace($templatePath)) {
                    Write-Host "Template path is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nUser mapping file (default: $userMappingFile): " -ForegroundColor Yellow -NoNewline
                $mappingPath = Read-Host
                if ([string]::IsNullOrWhiteSpace($mappingPath)) {
                    $mappingPath = $userMappingFile
                }
                
                if (-not (Test-Path $mappingPath)) {
                    Write-Host "`nERROR: User mapping file not found: $mappingPath" -ForegroundColor Red
                    Write-Host "Generate and edit the mapping file first (options 2 & 3)" -ForegroundColor Yellow
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nValidating users in target tenant..." -ForegroundColor Cyan
                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                Write-Host ".\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl `"$targetUrl`" -TemplatePath `"$templatePath`" -UserMappingFile `"$mappingPath`" -ConfigFile `"$TargetConfigFile`" -ValidateUsersOnly" -ForegroundColor White
                
                if (Test-Path ".\Import-SharePointSiteTemplate.ps1") {
                    try {
                        & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $targetUrl -TemplatePath $templatePath -UserMappingFile $mappingPath -ConfigFile $TargetConfigFile -ValidateUsersOnly
                        Write-Host "`n✓ User validation completed!" -ForegroundColor Green
                        Write-Host "If any users are invalid, update the CSV and re-validate." -ForegroundColor Cyan
                    }
                    catch {
                        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nERROR: Import-SharePointSiteTemplate.ps1 not found" -ForegroundColor Red
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "5" {
                # Import with user mapping
                Write-SubHeader "Step 5: Import to Target Tenant"
                
                Write-Host "`n⚠️  PREREQUISITES:" -ForegroundColor Yellow
                Write-Host "  1. Target site must exist in target tenant" -ForegroundColor Yellow
                Write-Host "  2. All target users must be validated (Step 4)`n" -ForegroundColor Yellow
                
                if ($targetDomain) {
                    Write-Host "Target tenant: " -ForegroundColor Cyan -NoNewline
                    Write-Host $targetDomain -ForegroundColor White
                }
                
                Write-Host "`nEnter target site URL: " -ForegroundColor Yellow -NoNewline
                $targetUrl = Read-Host
                
                if ([string]::IsNullOrWhiteSpace($targetUrl)) {
                    Write-Host "Target URL is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                if ($exportedTemplate) {
                    Write-Host "`nMost recent template: " -ForegroundColor Cyan
                    Write-Host $exportedTemplate -ForegroundColor White
                    Write-Host "`nUse this template? [Y/N] (default: Y): " -ForegroundColor Yellow -NoNewline
                    $useRecent = Read-Host
                    
                    if ($useRecent -eq 'N' -or $useRecent -eq 'n') {
                        Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                        $templatePath = Read-Host
                    }
                    else {
                        $templatePath = $exportedTemplate
                    }
                }
                else {
                    Write-Host "Enter template path: " -ForegroundColor Yellow -NoNewline
                    $templatePath = Read-Host
                }
                
                if ([string]::IsNullOrWhiteSpace($templatePath)) {
                    Write-Host "Template path is required. Operation cancelled." -ForegroundColor Red
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nUser mapping file (default: $userMappingFile): " -ForegroundColor Yellow -NoNewline
                $mappingPath = Read-Host
                if ([string]::IsNullOrWhiteSpace($mappingPath)) {
                    $mappingPath = $userMappingFile
                }
                
                if (-not (Test-Path $mappingPath)) {
                    Write-Host "`nERROR: User mapping file not found: $mappingPath" -ForegroundColor Red
                    Write-Host "Generate and edit the mapping file first (options 2 & 3)" -ForegroundColor Yellow
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`n⚠️  FINAL CONFIRMATION" -ForegroundColor Yellow
                Write-Host "About to import to: $targetUrl" -ForegroundColor White
                Write-Host "From template: $templatePath" -ForegroundColor White
                Write-Host "With user mapping: $mappingPath" -ForegroundColor White
                Write-Host "`nProceed with import? [Y/N]: " -ForegroundColor Yellow -NoNewline
                $confirm = Read-Host
                
                if ($confirm -ne 'Y' -and $confirm -ne 'y') {
                    Write-Host "Import cancelled." -ForegroundColor Cyan
                    Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                    continue
                }
                
                Write-Host "`nExecuting import with user mapping..." -ForegroundColor Cyan
                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                Write-Host ".\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl `"$targetUrl`" -TemplatePath `"$templatePath`" -UserMappingFile `"$mappingPath`" -ConfigFile `"$TargetConfigFile`" -IgnoreDuplicateDataRowErrors" -ForegroundColor White
                
                if (Test-Path ".\Import-SharePointSiteTemplate.ps1") {
                    try {
                        & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $targetUrl -TemplatePath $templatePath -UserMappingFile $mappingPath -ConfigFile $TargetConfigFile -IgnoreDuplicateDataRowErrors
                        Write-Host "`n✓ Import completed!" -ForegroundColor Green
                        Write-Host "Cross-tenant migration finished. Review logs for details." -ForegroundColor Cyan
                    }
                    catch {
                        Write-Host "`nERROR: $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host "`nERROR: Import-SharePointSiteTemplate.ps1 not found" -ForegroundColor Red
                }
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "6" {
                # View reference commands
                Write-SubHeader "Complete Workflow Reference"
                
                Write-Step 1 "Export from source tenant"
                Write-Host @"
    .\Export-SharePointSiteTemplate.ps1 ``
        -SourceSiteUrl "https://sourcetenant.sharepoint.com/sites/Site" ``
        -ConfigFile "$SourceConfigFile" ``
        -IncludeContent
"@ -ForegroundColor White
                
                Write-Step 2 "Generate user mapping template"
                Write-Host @"
    .\New-UserMappingTemplate.ps1 ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -OutputPath "$userMappingFile"
"@ -ForegroundColor White
                
                Write-Step 3 "Edit user mapping CSV"
                Write-Host "    Edit $userMappingFile - update TargetUser column" -ForegroundColor White
                if ($sourceDomain -and $targetDomain) {
                    Write-Host "    Replace @$sourceDomain with @$targetDomain" -ForegroundColor Gray
                }
                
                Write-Step 4 "Validate target users"
                Write-Host @"
    .\Import-SharePointSiteTemplate.ps1 ``
        -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -UserMappingFile "$userMappingFile" ``
        -ConfigFile "$TargetConfigFile" ``
        -ValidateUsersOnly
"@ -ForegroundColor White
                
                Write-Step 5 "Import with user mapping"
                Write-Host @"
    .\Import-SharePointSiteTemplate.ps1 ``
        -TargetSiteUrl "https://targettenant.sharepoint.com/sites/Site" ``
        -TemplatePath "C:\PSReports\SiteTemplates\SiteTemplate_*.pnp" ``
        -UserMappingFile "$userMappingFile" ``
        -ConfigFile "$TargetConfigFile" ``
        -IgnoreDuplicateDataRowErrors
"@ -ForegroundColor White
                
                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
            "7" {
                # Return to main menu
                $workflowRunning = $false
            }
            "0" {
                $workflowRunning = $false
            }
            default {
                Write-Host "`nInvalid choice. Press any key to continue..." -ForegroundColor Red
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
        }
    }
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
    $menuRunning = $true
    
    while ($menuRunning) {
        $choice = Show-Menu -Title "Documentation" `
            -Options @(
                "README.md - Main documentation and workflows",
                "MANUAL-APP-REGISTRATION.md - Azure AD app setup guide",
                "CONFIG-README.md - Configuration file guidance",
                "USER-MAPPING-QUICK-REF.md - User mapping quick reference",
                "DEVELOPER.md - Contribution guidelines",
                "Return to main menu"
            ) `
            -Prompt "Select document to view"
        
        switch ($choice) {
            "1" { if (Test-Path ".\README.md") { & notepad.exe ".\README.md" } }
            "2" { if (Test-Path ".\MANUAL-APP-REGISTRATION.md") { & notepad.exe ".\MANUAL-APP-REGISTRATION.md" } }
            "3" { if (Test-Path ".\CONFIG-README.md") { & notepad.exe ".\CONFIG-README.md" } }
            "4" { if (Test-Path ".\USER-MAPPING-QUICK-REF.md") { & notepad.exe ".\USER-MAPPING-QUICK-REF.md" } }
            "5" { if (Test-Path ".\DEVELOPER.md") { & notepad.exe ".\DEVELOPER.md" } }
            "6" { $menuRunning = $false }
            "0" { $menuRunning = $false }
            default { 
                Write-Host "`nInvalid choice. Press any key to continue..." -ForegroundColor Red
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
        }
    }
}

function Show-TestConfigMenu {
    $menuRunning = $true
    
    while ($menuRunning) {
        $choice = Show-Menu -Title "Test Configuration Files" `
            -Options @(
                "Test single configuration file (same-tenant)",
                "Test cross-tenant configurations (source + target)",
                "Return to main menu"
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
            "3" { $menuRunning = $false }
            "0" { $menuRunning = $false }
            default {
                Write-Host "`nInvalid choice. Press any key to continue..." -ForegroundColor Red
                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            }
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
                $continue = Read-Host "`nContinue to Migration options? (Y/N)"
                if ($continue -eq 'Y' -or $continue -eq 'y') {
                    Show-SameTenantWorkflow
                }
            }
            "2" {
                # Cross-Tenant Migration
                Show-CrossTenantPrerequisites
                $continue = Read-Host "`nContinue to Migration options? (Y/N)"
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
