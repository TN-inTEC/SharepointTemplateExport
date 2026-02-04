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
    param(
        [string]$TenantDomain = "",
        [string]$TenantRootUrl = "",
        [bool]$ConfigIsExample = $false
    )
    
    Write-Header "Same-Tenant Migration Workflow"
    
    Write-InfoBox @(
        "This workflow will guide you through migrating a SharePoint site",
        "within the same Microsoft 365 tenant."
    )
    
    if ($TenantDomain) {
        Write-Host "`nTenant: " -ForegroundColor Cyan -NoNewline
        Write-Host $TenantDomain -ForegroundColor White
        if ($TenantRootUrl) {
            Write-Host "Default URL: " -ForegroundColor Gray -NoNewline
            Write-Host $TenantRootUrl -ForegroundColor White
        }
    }
    
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
                
                # Parameter collection menu with config-based defaults
                $exportParams = @{
                    SourceUrl = if (-not $ConfigIsExample -and $TenantRootUrl) { $TenantRootUrl } else { "" }
                    IncludeContent = $true
                    PreviewMode = $false
                }
                
                $paramMenuRunning = $true
                while ($paramMenuRunning) {
                    $contentStatus = if ($exportParams.IncludeContent) { "Yes" } else { "No" }
                    $previewStatus = if ($exportParams.PreviewMode) { "Yes (no files created)" } else { "No (will create template)" }
                    $urlDisplay = if ($exportParams.SourceUrl) { $exportParams.SourceUrl } else { "(not set)" }
                    
                    Write-Host "`n$('=' * 70)" -ForegroundColor Cyan
                    Write-Host "  Export Configuration" -ForegroundColor Cyan
                    Write-Host "$('=' * 70)" -ForegroundColor Cyan
                    Write-Host "  Source URL:       $urlDisplay" -ForegroundColor White
                    Write-Host "  Include Content:  $contentStatus" -ForegroundColor White
                    Write-Host "  Preview Mode:     $previewStatus" -ForegroundColor White
                    Write-Host "$('=' * 70)`n" -ForegroundColor Cyan
                    
                    $paramChoice = Show-Menu -Title "Configure Export Parameters" `
                        -Options @(
                            "Set source site URL",
                            "Toggle include content (currently: $contentStatus)",
                            "Toggle preview mode (currently: $previewStatus)",
                            "Execute export",
                            "Cancel and return to workflow menu"
                        ) `
                        -Prompt "Select option"
                    
                    switch ($paramChoice) {
                        "1" {
                            Write-Host "`nEnter source site URL: " -ForegroundColor Yellow -NoNewline
                            $url = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($url)) {
                                $exportParams.SourceUrl = $url
                                Write-Host "✓ URL set" -ForegroundColor Green
                            }
                            Start-Sleep -Milliseconds 500
                        }
                        "2" {
                            $exportParams.IncludeContent = -not $exportParams.IncludeContent
                            Write-Host "`n✓ Include content toggled to: $(if ($exportParams.IncludeContent) { 'Yes' } else { 'No' })" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        "3" {
                            $exportParams.PreviewMode = -not $exportParams.PreviewMode
                            Write-Host "`n✓ Preview mode toggled to: $(if ($exportParams.PreviewMode) { 'Yes' } else { 'No' })" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        "4" {
                            # Execute export
                            if ([string]::IsNullOrWhiteSpace($exportParams.SourceUrl)) {
                                Write-Host "`nERROR: Source URL is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            
                            if ($exportParams.PreviewMode) {
                                Write-Host "`nExecuting preview (no files will be created)..." -ForegroundColor Cyan
                            }
                            else {
                                Write-Host "`nExecuting export..." -ForegroundColor Cyan
                            }
                            
                            $includeContentSwitch = if ($exportParams.IncludeContent) { "-IncludeContent" } else { "" }
                            $previewSwitch = if ($exportParams.PreviewMode) { "-Preview" } else { "" }
                            
                            Write-Host "Command: " -ForegroundColor Gray -NoNewline
                            Write-Host ".\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl `"$($exportParams.SourceUrl)`" -ConfigFile `"$ConfigFile`" $includeContentSwitch $previewSwitch" -ForegroundColor White
                            
                            if (Test-Path ".\Export-SharePointSiteTemplate.ps1") {
                                try {
                                    if ($exportParams.PreviewMode -and $exportParams.IncludeContent) {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $ConfigFile -IncludeContent -Preview
                                    }
                                    elseif ($exportParams.PreviewMode) {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $ConfigFile -Preview
                                    }
                                    elseif ($exportParams.IncludeContent) {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $ConfigFile -IncludeContent
                                    }
                                    else {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $ConfigFile
                                    }
                                    
                                    if (-not $exportParams.PreviewMode) {
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
                            $paramMenuRunning = $false
                        }
                        "5" {
                            $paramMenuRunning = $false
                        }
                        "0" {
                            $paramMenuRunning = $false
                        }
                        default {
                            Write-Host "`nInvalid choice" -ForegroundColor Red
                            Start-Sleep -Milliseconds 500
                        }
                    }
                }
            }
            "2" {
                # Inspect template
                Write-SubHeader "Step 2: Inspect Exported Template"
                
                # Parameter collection menu
                $inspectParams = @{
                    TemplatePath = if ($exportedTemplate) { $exportedTemplate } else { "" }
                }
                
                $paramMenuRunning = $true
                while ($paramMenuRunning) {
                    $pathDisplay = if ($inspectParams.TemplatePath) { $inspectParams.TemplatePath } else { "(not set)" }
                    
                    Write-Host "`n$('=' * 70)" -ForegroundColor Cyan
                    Write-Host "  Template Inspection Configuration" -ForegroundColor Cyan
                    Write-Host "$('=' * 70)" -ForegroundColor Cyan
                    Write-Host "  Template Path: $pathDisplay" -ForegroundColor White
                    Write-Host "$('=' * 70)`n" -ForegroundColor Cyan
                    
                    $options = @(
                        "Set template path"
                    )
                    if ($exportedTemplate) {
                        $options += "Use most recent export: $(Split-Path $exportedTemplate -Leaf)"
                    }
                    $options += @(
                        "Execute inspection",
                        "Cancel and return to workflow menu"
                    )
                    
                    $paramChoice = Show-Menu -Title "Configure Template Inspection" `
                        -Options $options `
                        -Prompt "Select option"
                    
                    switch ($paramChoice) {
                        "1" {
                            Write-Host "`nEnter template path: " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $inspectParams.TemplatePath = $path
                                Write-Host "✓ Template path set" -ForegroundColor Green
                            }
                            Start-Sleep -Milliseconds 500
                        }
                        "2" {
                            if ($exportedTemplate) {
                                $inspectParams.TemplatePath = $exportedTemplate
                                Write-Host "`n✓ Using most recent export" -ForegroundColor Green
                                Start-Sleep -Milliseconds 500
                            }
                            else {
                                # Execute inspection
                                if ([string]::IsNullOrWhiteSpace($inspectParams.TemplatePath)) {
                                    Write-Host "`nERROR: Template path is required" -ForegroundColor Red
                                    Write-Host "Press any key to continue..." -ForegroundColor Gray
                                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                    continue
                                }
                                
                                Write-Host "`nExecuting inspection..." -ForegroundColor Cyan
                                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                                Write-Host ".\Get-TemplateContent.ps1 -TemplatePath `"$($inspectParams.TemplatePath)`" -Detailed -ShowUsers -ShowContent" -ForegroundColor White
                                
                                if (Test-Path ".\Get-TemplateContent.ps1") {
                                    try {
                                        & ".\Get-TemplateContent.ps1" -TemplatePath $inspectParams.TemplatePath -Detailed -ShowUsers -ShowContent
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
                                $paramMenuRunning = $false
                            }
                        }
                        "3" {
                            if ($exportedTemplate) {
                                # Execute inspection
                                if ([string]::IsNullOrWhiteSpace($inspectParams.TemplatePath)) {
                                    Write-Host "`nERROR: Template path is required" -ForegroundColor Red
                                    Write-Host "Press any key to continue..." -ForegroundColor Gray
                                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                    continue
                                }
                                
                                Write-Host "`nExecuting inspection..." -ForegroundColor Cyan
                                Write-Host "Command: " -ForegroundColor Gray -NoNewline
                                Write-Host ".\Get-TemplateContent.ps1 -TemplatePath `"$($inspectParams.TemplatePath)`" -Detailed -ShowUsers -ShowContent" -ForegroundColor White
                                
                                if (Test-Path ".\Get-TemplateContent.ps1") {
                                    try {
                                        & ".\Get-TemplateContent.ps1" -TemplatePath $inspectParams.TemplatePath -Detailed -ShowUsers -ShowContent
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
                                $paramMenuRunning = $false
                            }
                            else {
                                $paramMenuRunning = $false
                            }
                        }
                        "4" {
                            $paramMenuRunning = $false
                        }
                        "0" {
                            $paramMenuRunning = $false
                        }
                        default {
                            Write-Host "`nInvalid choice" -ForegroundColor Red
                            Start-Sleep -Milliseconds 500
                        }
                    }
                }
            }
            "3" {
                # Import to target - menu-based parameter collection
                Write-SubHeader "Step 3: Import to Target Site"
                
                # Initialize parameters with config-based defaults
                $importParams = @{
                    TargetUrl = if (-not $ConfigIsExample -and $TenantRootUrl) { $TenantRootUrl } else { "" }
                    TemplatePath = if ($exportedTemplate) { $exportedTemplate } else { "" }
                    InspectOnly = $false
                }
                
                $importMenuRunning = $true
                while ($importMenuRunning) {
                    Write-Host "`n═══ Current Configuration ═══" -ForegroundColor Cyan
                    Write-Host "  Target Site URL: " -NoNewline -ForegroundColor Gray
                    if ($importParams.TargetUrl) {
                        Write-Host $importParams.TargetUrl -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  Template Path: " -NoNewline -ForegroundColor Gray
                    if ($importParams.TemplatePath) {
                        Write-Host $importParams.TemplatePath -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  Mode: " -NoNewline -ForegroundColor Gray
                    if ($importParams.InspectOnly) {
                        Write-Host "INSPECT ONLY (validation only, no import)" -ForegroundColor Cyan
                    } else {
                        Write-Host "IMPORT (will modify target site)" -ForegroundColor Yellow
                    }
                    
                    Write-Host "`n⚠️  PREREQUISITE: Target site must already exist!" -ForegroundColor Yellow
                    
                    $importOptions = @(
                        "Set target site URL"
                    )
                    if (-not $ConfigIsExample -and $TenantRootUrl) {
                        $importOptions += "Reset to default tenant URL: $TenantRootUrl"
                    }
                    $importOptions += "Set template path"
                    if ($exportedTemplate) {
                        $importOptions += "Use most recent export: $(Split-Path $exportedTemplate -Leaf)"
                    }
                    $importOptions += @(
                        "Toggle mode (Inspect Only / Import)",
                        "Execute import",
                        "Cancel and return to workflow menu"
                    )
                    
                    $importChoice = Show-Menu -Title "Import Options" -Options $importOptions
                    
                    # Adjust switch based on dynamic menu
                    $optionOffset = if (-not $ConfigIsExample -and $TenantRootUrl) { 1 } else { 0 }
                    $useRecentOffset = if ($exportedTemplate) { 1 } else { 0 }
                    
                    switch ($importChoice) {
                        "1" {
                            Write-Host "`nEnter target site URL: " -ForegroundColor Yellow -NoNewline
                            $url = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($url)) {
                                $importParams.TargetUrl = $url
                            }
                        }
                        { $_ -eq "2" -and (-not $ConfigIsExample -and $TenantRootUrl) } {
                            $importParams.TargetUrl = $TenantRootUrl
                            Write-Host "`n✓ Reset to default tenant URL" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (2 + $optionOffset) } {
                            Write-Host "`nEnter template path: " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $importParams.TemplatePath = $path
                            }
                        }
                        { $_ -eq (3 + $optionOffset) -and $exportedTemplate } {
                            $importParams.TemplatePath = $exportedTemplate
                            Write-Host "`n✓ Using most recent export" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (3 + $optionOffset + $useRecentOffset) } {
                            $importParams.InspectOnly = -not $importParams.InspectOnly
                            Write-Host "`n✓ Mode toggled" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (4 + $optionOffset + $useRecentOffset) } {
                            # Execute
                            if ([string]::IsNullOrWhiteSpace($importParams.TargetUrl)) {
                                Write-Host "`n✗ Target URL is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            if ([string]::IsNullOrWhiteSpace($importParams.TemplatePath)) {
                                Write-Host "`n✗ Template path is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            
                            if (-not $importParams.InspectOnly) {
                                Write-Host "`n⚠️  FINAL CONFIRMATION" -ForegroundColor Yellow
                                Write-Host "About to import to: $($importParams.TargetUrl)" -ForegroundColor White
                                Write-Host "From template: $($importParams.TemplatePath)" -ForegroundColor White
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
                            
                            if ($importParams.InspectOnly) {
                                Write-Host "`nExecuting inspection (no changes will be made)..." -ForegroundColor Cyan
                            }
                            else {
                                Write-Host "`nExecuting import..." -ForegroundColor Cyan
                            }
                            
                            $inspectSwitch = if ($importParams.InspectOnly) { "-InspectOnly" } else { "" }
                            Write-Host "Command: " -ForegroundColor Gray -NoNewline
                            Write-Host ".\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl `"$($importParams.TargetUrl)`" -TemplatePath `"$($importParams.TemplatePath)`" -ConfigFile `"$ConfigFile`" $inspectSwitch" -ForegroundColor White
                            
                            if (Test-Path ".\Import-SharePointSiteTemplate.ps1") {
                                try {
                                    if ($importParams.InspectOnly) {
                                        & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $importParams.TargetUrl -TemplatePath $importParams.TemplatePath -ConfigFile $ConfigFile -InspectOnly
                                        Write-Host "`n✓ Inspection completed (no changes made)" -ForegroundColor Green
                                    }
                                    else {
                                        & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $importParams.TargetUrl -TemplatePath $importParams.TemplatePath -ConfigFile $ConfigFile
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
                            $importMenuRunning = $false
                        }
                        { $_ -eq (5 + $optionOffset + $useRecentOffset) } {
                            # Cancel and return
                            $importMenuRunning = $false
                        }
                    }
                }
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
    
    # Helper function to detect if config contains example/placeholder values
    function Test-IsExampleConfig {
        param($tenantDomain)
        return ($tenantDomain -match "example\.onmicrosoft\.com|targettenant\.onmicrosoft\.com|yourtenant\.onmicrosoft\.com" -or 
                [string]::IsNullOrWhiteSpace($tenantDomain))
    }
    
    # Helper function to construct SharePoint root URL from tenant domain
    function Get-SharePointRootUrl {
        param($tenantDomain)
        if ([string]::IsNullOrWhiteSpace($tenantDomain)) { return "" }
        # Extract tenant name (e.g., "slm.co.uk" -> "slm")
        $tenantName = $tenantDomain -replace '\.onmicrosoft\.com$|\..+$', ''
        return "https://$tenantName.sharepoint.com"
    }
    
    # Load config files to get tenant domains
    $sourceConfig = $null
    $targetConfig = $null
    $sourceDomain = ""
    $targetDomain = ""
    $sourceRootUrl = ""
    $targetRootUrl = ""
    $sourceConfigIsExample = $false
    $targetConfigIsExample = $false
    
    if (Test-Path $SourceConfigFile) {
        try {
            $sourceConfig = Get-Content $SourceConfigFile -Raw | ConvertFrom-Json
            $sourceDomain = $sourceConfig.tenantDomain
            $sourceConfigIsExample = Test-IsExampleConfig $sourceDomain
            if (-not $sourceConfigIsExample) {
                $sourceRootUrl = Get-SharePointRootUrl $sourceDomain
            }
        }
        catch {
            Write-Host "Warning: Could not read source config file" -ForegroundColor Yellow
        }
    }
    
    if (Test-Path $TargetConfigFile) {
        try {
            $targetConfig = Get-Content $TargetConfigFile -Raw | ConvertFrom-Json
            $targetDomain = $targetConfig.tenantDomain
            $targetConfigIsExample = Test-IsExampleConfig $targetDomain
            if (-not $targetConfigIsExample) {
                $targetRootUrl = Get-SharePointRootUrl $targetDomain
            }
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
                # Export from source - menu-based parameter collection
                Write-SubHeader "Step 1: Export from Source Tenant"
                
                if ($sourceDomain) {
                    Write-Host "`nSource tenant: " -ForegroundColor Cyan -NoNewline
                    Write-Host $sourceDomain -ForegroundColor White
                    if ($sourceRootUrl) {
                        Write-Host "Default URL: " -ForegroundColor Gray -NoNewline
                        Write-Host $sourceRootUrl -ForegroundColor White
                    }
                }
                
                # Initialize parameters with config-based defaults
                $exportParams = @{
                    SourceUrl = if (-not $sourceConfigIsExample -and $sourceRootUrl) { $sourceRootUrl } else { "" }
                    IncludeContent = $true
                    PreviewMode = $false
                }
                
                $exportMenuRunning = $true
                while ($exportMenuRunning) {
                    Write-Host "`n═══ Current Configuration ═══" -ForegroundColor Cyan
                    Write-Host "  Source Site URL: " -NoNewline -ForegroundColor Gray
                    if ($exportParams.SourceUrl) {
                        Write-Host $exportParams.SourceUrl -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  Include Content: " -NoNewline -ForegroundColor Gray
                    Write-Host $(if ($exportParams.IncludeContent) { "Yes" } else { "No" }) -ForegroundColor White
                    Write-Host "  Preview Mode: " -NoNewline -ForegroundColor Gray
                    Write-Host $(if ($exportParams.PreviewMode) { "Yes (no files created)" } else { "No" }) -ForegroundColor White
                    
                    $exportOptions = @(
                        "Set source site URL"
                    )
                    if (-not $sourceConfigIsExample -and $sourceRootUrl) {
                        $exportOptions += "Reset to default tenant URL: $sourceRootUrl"
                    }
                    $exportOptions += @(
                        "Toggle include content (currently: $(if ($exportParams.IncludeContent) { 'Yes' } else { 'No' }))",
                        "Toggle preview mode (currently: $(if ($exportParams.PreviewMode) { 'Yes' } else { 'No' }))",
                        "Execute export",
                        "Cancel and return to workflow menu"
                    )
                    
                    $exportChoice = Show-Menu -Title "Export Options" -Options $exportOptions
                    
                    # Adjust switch based on dynamic menu
                    $optionOffset = if (-not $sourceConfigIsExample -and $sourceRootUrl) { 1 } else { 0 }
                    
                    switch ($exportChoice) {
                        "1" {
                            Write-Host "`nEnter source site URL: " -ForegroundColor Yellow -NoNewline
                            $url = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($url)) {
                                $exportParams.SourceUrl = $url
                            }
                        }
                        { $_ -eq "2" -and (-not $sourceConfigIsExample -and $sourceRootUrl) } {
                            $exportParams.SourceUrl = $sourceRootUrl
                            Write-Host "`n✓ Reset to default tenant URL" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (2 + $optionOffset) } {
                            $exportParams.IncludeContent = -not $exportParams.IncludeContent
                            Write-Host "`n✓ Include content toggled to: $(if ($exportParams.IncludeContent) { 'Yes' } else { 'No' })" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (3 + $optionOffset) } {
                            $exportParams.PreviewMode = -not $exportParams.PreviewMode
                            Write-Host "`n✓ Preview mode toggled to: $(if ($exportParams.PreviewMode) { 'Yes' } else { 'No' })" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (4 + $optionOffset) } {
                            # Execute
                            if ([string]::IsNullOrWhiteSpace($exportParams.SourceUrl)) {
                                Write-Host "`n✗ Source URL is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            
                            if ($exportParams.PreviewMode) {
                                Write-Host "`nExecuting preview (no files will be created)..." -ForegroundColor Cyan
                            }
                            else {
                                Write-Host "`nExecuting export..." -ForegroundColor Cyan
                            }
                            
                            $cmdPreview = if ($exportParams.PreviewMode) { " -Preview" } else { "" }
                            $cmdContent = if ($exportParams.IncludeContent) { " -IncludeContent" } else { "" }
                            Write-Host "Command: " -ForegroundColor Gray -NoNewline
                            Write-Host ".\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl `"$($exportParams.SourceUrl)`" -ConfigFile `"$SourceConfigFile`"$cmdContent$cmdPreview" -ForegroundColor White
                            
                            if (Test-Path ".\Export-SharePointSiteTemplate.ps1") {
                                try {
                                    if ($exportParams.PreviewMode -and $exportParams.IncludeContent) {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $SourceConfigFile -IncludeContent -Preview
                                    }
                                    elseif ($exportParams.PreviewMode) {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $SourceConfigFile -Preview
                                    }
                                    elseif ($exportParams.IncludeContent) {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $SourceConfigFile -IncludeContent
                                    }
                                    else {
                                        & ".\Export-SharePointSiteTemplate.ps1" -SourceSiteUrl $exportParams.SourceUrl -ConfigFile $SourceConfigFile
                                    }
                                    
                                    if (-not $exportParams.PreviewMode) {
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
                            $exportMenuRunning = $false
                        }
                        { $_ -eq (5 + $optionOffset) } {
                            # Cancel and return
                            $exportMenuRunning = $false
                        }
                    }
                }
            }
            "2" {
                # Generate user mapping - menu-based parameter collection
                Write-SubHeader "Step 2: Generate User Mapping Template"
                
                # Initialize parameters
                $mappingParams = @{
                    TemplatePath = if ($exportedTemplate) { $exportedTemplate } else { "" }
                    OutputPath = "user-mapping.csv"
                }
                
                $mappingMenuRunning = $true
                while ($mappingMenuRunning) {
                    Write-Host "`n═══ Current Configuration ═══" -ForegroundColor Cyan
                    Write-Host "  Template Path: " -NoNewline -ForegroundColor Gray
                    if ($mappingParams.TemplatePath) {
                        Write-Host $mappingParams.TemplatePath -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  Output CSV Path: " -NoNewline -ForegroundColor Gray
                    Write-Host $mappingParams.OutputPath -ForegroundColor White
                    
                    if ($sourceDomain -and $targetDomain) {
                        Write-Host "`n  Domain mapping hint: " -ForegroundColor Gray
                        Write-Host "    Source: @$sourceDomain" -ForegroundColor Cyan
                        Write-Host "    Target: @$targetDomain" -ForegroundColor Cyan
                    }
                    
                    $mappingOptions = @(
                        "Set template path"
                    )
                    if ($exportedTemplate) {
                        $mappingOptions += "Use most recent export: $(Split-Path $exportedTemplate -Leaf)"
                    }
                    $mappingOptions += @(
                        "Set output CSV path",
                        "Execute generation",
                        "Cancel and return to workflow menu"
                    )
                    
                    $mappingChoice = Show-Menu -Title "User Mapping Generation Options" -Options $mappingOptions
                    
                    switch ($mappingChoice) {
                        "1" {
                            Write-Host "`nEnter template path: " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $mappingParams.TemplatePath = $path
                            }
                        }
                        { $_ -eq "2" -and $exportedTemplate } {
                            $mappingParams.TemplatePath = $exportedTemplate
                            Write-Host "`n✓ Using most recent export" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { ($_ -eq "2" -and -not $exportedTemplate) -or ($_ -eq "3" -and $exportedTemplate) } {
                            Write-Host "`nOutput CSV path (press Enter for default: user-mapping.csv): " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $mappingParams.OutputPath = $path
                            }
                        }
                        { ($_ -eq "3" -and -not $exportedTemplate) -or ($_ -eq "4" -and $exportedTemplate) } {
                            # Execute
                            if ([string]::IsNullOrWhiteSpace($mappingParams.TemplatePath)) {
                                Write-Host "`n✗ Template path is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            
                            Write-Host "`nExecuting user mapping generation..." -ForegroundColor Cyan
                            Write-Host "Command: " -ForegroundColor Gray -NoNewline
                            Write-Host ".\New-UserMappingTemplate.ps1 -TemplatePath `"$($mappingParams.TemplatePath)`" -OutputPath `"$($mappingParams.OutputPath)`"" -ForegroundColor White
                            
                            if (Test-Path ".\New-UserMappingTemplate.ps1") {
                                try {
                                    & ".\New-UserMappingTemplate.ps1" -TemplatePath $mappingParams.TemplatePath -OutputPath $mappingParams.OutputPath
                                    $userMappingFile = $mappingParams.OutputPath
                                    Write-Host "`n✓ User mapping template created!" -ForegroundColor Green
                                    Write-Host "File: $userMappingFile" -ForegroundColor Cyan
                                    
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
                            $mappingMenuRunning = $false
                        }
                        { ($_ -eq "4" -and -not $exportedTemplate) -or ($_ -eq "5" -and $exportedTemplate) } {
                            # Cancel and return
                            $mappingMenuRunning = $false
                        }
                    }
                }
            }
            "3" {
                # Edit user mapping - menu-based with confirmation
                Write-SubHeader "Step 3: Edit User Mapping CSV"
                
                # Initialize parameters
                $editMappingPath = $userMappingFile
                
                $editMenuRunning = $true
                while ($editMenuRunning) {
                    Write-Host "`n═══ Current Configuration ═══" -ForegroundColor Cyan
                    Write-Host "  Mapping File: " -NoNewline -ForegroundColor Gray
                    Write-Host $editMappingPath -ForegroundColor White
                    Write-Host "  File Exists: " -NoNewline -ForegroundColor Gray
                    if (Test-Path $editMappingPath) {
                        Write-Host "Yes" -ForegroundColor Green
                    } else {
                        Write-Host "No" -ForegroundColor Red
                    }
                    
                    if ($sourceDomain -and $targetDomain) {
                        Write-Host "`n  Domain mapping reminder: " -ForegroundColor Gray
                        Write-Host "    Source: @$sourceDomain" -ForegroundColor Cyan
                        Write-Host "    Target: @$targetDomain" -ForegroundColor Cyan
                        Write-Host "`n  Example mappings:" -ForegroundColor Gray
                        Write-Host "    john.smith@$sourceDomain → john.smith@$targetDomain" -ForegroundColor White
                        Write-Host "    admin@$sourceDomain → it.admin@$targetDomain" -ForegroundColor White
                    }
                    
                    $editChoice = Show-Menu -Title "Edit User Mapping Options" -Options @(
                        "Set mapping file path",
                        "Open file in default editor",
                        "Cancel and return to workflow menu"
                    )
                    
                    switch ($editChoice) {
                        "1" {
                            Write-Host "`nMapping file path (press Enter for default: $userMappingFile): " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $editMappingPath = $path
                            }
                        }
                        "2" {
                            # Open file
                            if (Test-Path $editMappingPath) {
                                Write-Host "`nOpening $editMappingPath in default editor..." -ForegroundColor Cyan
                                
                                try {
                                    Start-Process $editMappingPath
                                    Write-Host "`n✓ File opened in default application" -ForegroundColor Green
                                    Write-Host "Edit the file and save when complete." -ForegroundColor Cyan
                                }
                                catch {
                                    Write-Host "`nERROR: Could not open file: $($_.Exception.Message)" -ForegroundColor Red
                                    Write-Host "Please open manually: $editMappingPath" -ForegroundColor Yellow
                                }
                                
                                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                            }
                            else {
                                Write-Host "`n✗ File not found: $editMappingPath" -ForegroundColor Red
                                Write-Host "Generate the user mapping template first (option 2)" -ForegroundColor Yellow
                                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                            }
                            $editMenuRunning = $false
                        }
                        "3" {
                            # Cancel and return
                            $editMenuRunning = $false
                        }
                    }
                }
            }
            "4" {
                # Validate users - menu-based parameter collection
                Write-SubHeader "Step 4: Validate Target Users"
                
                if ($targetDomain) {
                    Write-Host "`nTarget tenant: " -ForegroundColor Cyan -NoNewline
                    Write-Host $targetDomain -ForegroundColor White
                    if ($targetRootUrl) {
                        Write-Host "Default URL: " -ForegroundColor Gray -NoNewline
                        Write-Host $targetRootUrl -ForegroundColor White
                    }
                }
                
                # Initialize parameters with config-based defaults
                $validateParams = @{
                    TargetUrl = if (-not $targetConfigIsExample -and $targetRootUrl) { $targetRootUrl } else { "" }
                    TemplatePath = if ($exportedTemplate) { $exportedTemplate } else { "" }
                    MappingFile = $userMappingFile
                }
                
                $validateMenuRunning = $true
                while ($validateMenuRunning) {
                    Write-Host "`n═══ Current Configuration ═══" -ForegroundColor Cyan
                    Write-Host "  Target Site URL: " -NoNewline -ForegroundColor Gray
                    if ($validateParams.TargetUrl) {
                        Write-Host $validateParams.TargetUrl -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  Template Path: " -NoNewline -ForegroundColor Gray
                    if ($validateParams.TemplatePath) {
                        Write-Host $validateParams.TemplatePath -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  User Mapping File: " -NoNewline -ForegroundColor Gray
                    Write-Host $validateParams.MappingFile -ForegroundColor White
                    Write-Host "  Mapping File Exists: " -NoNewline -ForegroundColor Gray
                    if (Test-Path $validateParams.MappingFile) {
                        Write-Host "Yes" -ForegroundColor Green
                    } else {
                        Write-Host "No" -ForegroundColor Red
                    }
                    
                    $validateOptions = @(
                        "Set target site URL"
                    )
                    if (-not $targetConfigIsExample -and $targetRootUrl) {
                        $validateOptions += "Reset to default tenant URL: $targetRootUrl"
                    }
                    $validateOptions += "Set template path"
                    if ($exportedTemplate) {
                        $validateOptions += "Use most recent export: $(Split-Path $exportedTemplate -Leaf)"
                    }
                    $validateOptions += @(
                        "Set user mapping file path",
                        "Execute validation",
                        "Cancel and return to workflow menu"
                    )
                    
                    $validateChoice = Show-Menu -Title "User Validation Options" -Options $validateOptions
                    
                    # Adjust switch based on dynamic menu
                    $optionOffset = if (-not $targetConfigIsExample -and $targetRootUrl) { 1 } else { 0 }
                    $useRecentOffset = if ($exportedTemplate) { 1 } else { 0 }
                    
                    switch ($validateChoice) {
                        "1" {
                            Write-Host "`nEnter target site URL: " -ForegroundColor Yellow -NoNewline
                            $url = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($url)) {
                                $validateParams.TargetUrl = $url
                            }
                        }
                        { $_ -eq "2" -and (-not $targetConfigIsExample -and $targetRootUrl) } {
                            $validateParams.TargetUrl = $targetRootUrl
                            Write-Host "`n✓ Reset to default tenant URL" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (2 + $optionOffset) } {
                            Write-Host "`nEnter template path: " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $validateParams.TemplatePath = $path
                            }
                        }
                        { $_ -eq (3 + $optionOffset) -and $exportedTemplate } {
                            $validateParams.TemplatePath = $exportedTemplate
                            Write-Host "`n✓ Using most recent export" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (3 + $optionOffset + $useRecentOffset) } {
                            Write-Host "`nMapping file path (press Enter for default: $userMappingFile): " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $validateParams.MappingFile = $path
                            }
                        }
                        { $_ -eq (4 + $optionOffset + $useRecentOffset) } {
                            # Execute validation
                            if ([string]::IsNullOrWhiteSpace($validateParams.TargetUrl)) {
                                Write-Host "`n✗ Target URL is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            if ([string]::IsNullOrWhiteSpace($validateParams.TemplatePath)) {
                                Write-Host "`n✗ Template path is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            if (-not (Test-Path $validateParams.MappingFile)) {
                                Write-Host "`n✗ User mapping file not found: $($validateParams.MappingFile)" -ForegroundColor Red
                                Write-Host "Generate and edit the mapping file first (options 2 & 3)" -ForegroundColor Yellow
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            
                            Write-Host "`nValidating users in target tenant..." -ForegroundColor Cyan
                            Write-Host "Command: " -ForegroundColor Gray -NoNewline
                            Write-Host ".\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl `"$($validateParams.TargetUrl)`" -TemplatePath `"$($validateParams.TemplatePath)`" -UserMappingFile `"$($validateParams.MappingFile)`" -ConfigFile `"$TargetConfigFile`" -ValidateUsersOnly" -ForegroundColor White
                            
                            if (Test-Path ".\Import-SharePointSiteTemplate.ps1") {
                                try {
                                    & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $validateParams.TargetUrl -TemplatePath $validateParams.TemplatePath -UserMappingFile $validateParams.MappingFile -ConfigFile $TargetConfigFile -ValidateUsersOnly
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
                            $validateMenuRunning = $false
                        }
                        { $_ -eq (5 + $optionOffset + $useRecentOffset) } {
                            # Cancel and return
                            $validateMenuRunning = $false
                        }
                    }
                }
            }
            "5" {
                # Import with user mapping - menu-based parameter collection
                Write-SubHeader "Step 5: Import to Target Tenant"
                
                Write-Host "`n⚠️  PREREQUISITES:" -ForegroundColor Yellow
                Write-Host "  1. Target site must exist in target tenant" -ForegroundColor Yellow
                Write-Host "  2. All target users must be validated (Step 4)`n" -ForegroundColor Yellow
                
                if ($targetDomain) {
                    Write-Host "Target tenant: " -ForegroundColor Cyan -NoNewline
                    Write-Host $targetDomain -ForegroundColor White
                    if ($targetRootUrl) {
                        Write-Host "Default URL: " -ForegroundColor Gray -NoNewline
                        Write-Host $targetRootUrl -ForegroundColor White
                    }
                }
                
                # Initialize parameters with config-based defaults
                $crossTenantImportParams = @{
                    TargetUrl = if (-not $targetConfigIsExample -and $targetRootUrl) { $targetRootUrl } else { "" }
                    TemplatePath = if ($exportedTemplate) { $exportedTemplate } else { "" }
                    MappingFile = $userMappingFile
                }
                
                $crossTenantImportMenuRunning = $true
                while ($crossTenantImportMenuRunning) {
                    Write-Host "`n═══ Current Configuration ═══" -ForegroundColor Cyan
                    Write-Host "  Target Site URL: " -NoNewline -ForegroundColor Gray
                    if ($crossTenantImportParams.TargetUrl) {
                        Write-Host $crossTenantImportParams.TargetUrl -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  Template Path: " -NoNewline -ForegroundColor Gray
                    if ($crossTenantImportParams.TemplatePath) {
                        Write-Host $crossTenantImportParams.TemplatePath -ForegroundColor White
                    } else {
                        Write-Host "(not set)" -ForegroundColor Yellow
                    }
                    Write-Host "  User Mapping File: " -NoNewline -ForegroundColor Gray
                    Write-Host $crossTenantImportParams.MappingFile -ForegroundColor White
                    Write-Host "  Mapping File Exists: " -NoNewline -ForegroundColor Gray
                    if (Test-Path $crossTenantImportParams.MappingFile) {
                        Write-Host "Yes" -ForegroundColor Green
                    } else {
                        Write-Host "No" -ForegroundColor Red
                    }
                    
                    $crossTenantImportOptions = @(
                        "Set target site URL"
                    )
                    if (-not $targetConfigIsExample -and $targetRootUrl) {
                        $crossTenantImportOptions += "Reset to default tenant URL: $targetRootUrl"
                    }
                    $crossTenantImportOptions += "Set template path"
                    if ($exportedTemplate) {
                        $crossTenantImportOptions += "Use most recent export: $(Split-Path $exportedTemplate -Leaf)"
                    }
                    $crossTenantImportOptions += @(
                        "Set user mapping file path",
                        "Execute import (requires confirmation)",
                        "Cancel and return to workflow menu"
                    )
                    
                    $crossTenantImportChoice = Show-Menu -Title "Cross-Tenant Import Options" -Options $crossTenantImportOptions
                    
                    # Adjust switch based on dynamic menu
                    $optionOffset = if (-not $targetConfigIsExample -and $targetRootUrl) { 1 } else { 0 }
                    $useRecentOffset = if ($exportedTemplate) { 1 } else { 0 }
                    
                    switch ($crossTenantImportChoice) {
                        "1" {
                            Write-Host "`nEnter target site URL: " -ForegroundColor Yellow -NoNewline
                            $url = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($url)) {
                                $crossTenantImportParams.TargetUrl = $url
                            }
                        }
                        { $_ -eq "2" -and (-not $targetConfigIsExample -and $targetRootUrl) } {
                            $crossTenantImportParams.TargetUrl = $targetRootUrl
                            Write-Host "`n✓ Reset to default tenant URL" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (2 + $optionOffset) } {
                            Write-Host "`nEnter template path: " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $crossTenantImportParams.TemplatePath = $path
                            }
                        }
                        { $_ -eq (3 + $optionOffset) -and $exportedTemplate } {
                            $crossTenantImportParams.TemplatePath = $exportedTemplate
                            Write-Host "`n✓ Using most recent export" -ForegroundColor Green
                            Start-Sleep -Milliseconds 500
                        }
                        { $_ -eq (3 + $optionOffset + $useRecentOffset) } {
                            Write-Host "`nMapping file path (press Enter for default: $userMappingFile): " -ForegroundColor Yellow -NoNewline
                            $path = Read-Host
                            if (-not [string]::IsNullOrWhiteSpace($path)) {
                                $crossTenantImportParams.MappingFile = $path
                            }
                        }
                        { $_ -eq (4 + $optionOffset + $useRecentOffset) } {
                            # Execute import with full validation and confirmation
                            if ([string]::IsNullOrWhiteSpace($crossTenantImportParams.TargetUrl)) {
                                Write-Host "`n✗ Target URL is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            if ([string]::IsNullOrWhiteSpace($crossTenantImportParams.TemplatePath)) {
                                Write-Host "`n✗ Template path is required" -ForegroundColor Red
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            if (-not (Test-Path $crossTenantImportParams.MappingFile)) {
                                Write-Host "`n✗ User mapping file not found: $($crossTenantImportParams.MappingFile)" -ForegroundColor Red
                                Write-Host "Generate and edit the mapping file first (options 2 & 3)" -ForegroundColor Yellow
                                Write-Host "Press any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            
                            Write-Host "`n⚠️  FINAL CONFIRMATION" -ForegroundColor Yellow
                            Write-Host "About to import to: $($crossTenantImportParams.TargetUrl)" -ForegroundColor White
                            Write-Host "From template: $($crossTenantImportParams.TemplatePath)" -ForegroundColor White
                            Write-Host "With user mapping: $($crossTenantImportParams.MappingFile)" -ForegroundColor White
                            Write-Host "`nThis will modify the target site!" -ForegroundColor Yellow
                            Write-Host "Proceed with import? [Y/N]: " -ForegroundColor Yellow -NoNewline
                            $confirm = Read-Host
                            
                            if ($confirm -ne 'Y' -and $confirm -ne 'y') {
                                Write-Host "Import cancelled." -ForegroundColor Cyan
                                Write-Host "`nPress any key to continue..." -ForegroundColor Gray
                                $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
                                continue
                            }
                            
                            Write-Host "`nExecuting import with user mapping..." -ForegroundColor Cyan
                            Write-Host "Command: " -ForegroundColor Gray -NoNewline
                            Write-Host ".\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl `"$($crossTenantImportParams.TargetUrl)`" -TemplatePath `"$($crossTenantImportParams.TemplatePath)`" -UserMappingFile `"$($crossTenantImportParams.MappingFile)`" -ConfigFile `"$TargetConfigFile`" -IgnoreDuplicateDataRowErrors" -ForegroundColor White
                            
                            if (Test-Path ".\Import-SharePointSiteTemplate.ps1") {
                                try {
                                    & ".\Import-SharePointSiteTemplate.ps1" -TargetSiteUrl $crossTenantImportParams.TargetUrl -TemplatePath $crossTenantImportParams.TemplatePath -UserMappingFile $crossTenantImportParams.MappingFile -ConfigFile $TargetConfigFile -IgnoreDuplicateDataRowErrors
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
                            $crossTenantImportMenuRunning = $false
                        }
                        { $_ -eq (5 + $optionOffset + $useRecentOffset) } {
                            # Cancel and return
                            $crossTenantImportMenuRunning = $false
                        }
                    }
                }
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
    
    # Load same-tenant config for tenant URL detection
    $sameTenantDomain = ""
    $sameTenantRootUrl = ""
    $sameTenantConfigIsExample = $false
    
    if (Test-Path $ConfigFile) {
        try {
            $sameTenantConfig = Get-Content $ConfigFile -Raw | ConvertFrom-Json
            $sameTenantDomain = $sameTenantConfig.tenantDomain
            # Check if it's an example config
            $sameTenantConfigIsExample = ($sameTenantDomain -match "example\.onmicrosoft\.com|yourtenant\.onmicrosoft\.com" -or 
                                         [string]::IsNullOrWhiteSpace($sameTenantDomain))
            # Generate root URL if not example
            if (-not $sameTenantConfigIsExample) {
                $tenantName = $sameTenantDomain -replace '\.onmicrosoft\.com$|\..+$', ''
                $sameTenantRootUrl = "https://$tenantName.sharepoint.com"
            }
        }
        catch {
            Write-Host "Warning: Could not read same-tenant config file $ConfigFile" -ForegroundColor Yellow
        }
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
                    Show-SameTenantWorkflow -TenantDomain $sameTenantDomain -TenantRootUrl $sameTenantRootUrl -ConfigIsExample $sameTenantConfigIsExample
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
