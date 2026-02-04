<#
.SYNOPSIS
    Imports a PnP site template to a SharePoint site for deployment.

.DESCRIPTION
    This script applies a PnP template to a target SharePoint site, importing structure,
    configuration, pages, lists, libraries, and content. Designed for site migrations
    like deploying SLM Academy template to SMC SharePoint.

.PARAMETER TargetSiteUrl
    The URL of the SharePoint site where the template will be applied.

.PARAMETER TemplatePath
    Full path to the .pnp template file to import.

.PARAMETER ClearNavigation
    Clear existing navigation before applying template.

.PARAMETER OverwriteSystemPropertyBagValues
    Overwrite system property bag values during import.

.PARAMETER ProvisionFieldsStoSite
    Provision fields to the site collection.
.PARAMETER IgnoreDuplicateDataRowErrors
    Continue importing even if there are duplicate or malformed data row errors. Recommended for cross-tenant migrations.
.PARAMETER ClientId
    Azure AD App Client ID for authentication. Uses PnP Management Shell if not specified.

.PARAMETER Tenant
    Tenant name (e.g., contoso) for authentication.

.PARAMETER ConfigFile
    Path to app-config.json file. Default: app-config.json in script directory.

.PARAMETER WhatIf
    Shows what would happen if the script runs without making changes.

.EXAMPLE
    .\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/SMC" `
        -TemplatePath "C:\PSReports\SiteTemplates\SLM_Academy_Full.pnp"

.EXAMPLE
    .\Import-SharePointSiteTemplate.ps1 -TargetSiteUrl "https://contoso.sharepoint.com/sites/SMC" `
        -TemplatePath "C:\PSReports\SiteTemplates\SLM_Academy_Full.pnp" -ClearNavigation -WhatIf

.NOTES
    Author: IT Support
    Date: February 3, 2026
    Requires: PnP.PowerShell module
    
    IMPORTANT: 
    - Test in a non-production environment first
    - Ensure you have Site Collection Admin permissions
    - Backup target site before applying template
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[^/]+\.sharepoint\.com/.*$')]
    [string]$TargetSiteUrl,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Template file not found: $_"
        }
        if ($_ -notmatch '\.pnp$') {
            throw "File must be a .pnp template file"
        }
        return $true
    })]
    [string]$TemplatePath,

    [Parameter(Mandatory = $false)]
    [switch]$ClearNavigation,

    [Parameter(Mandatory = $false)]
    [switch]$OverwriteSystemPropertyBagValues,

    [Parameter(Mandatory = $false)]
    [switch]$ProvisionFieldsToSite,

    [Parameter(Mandatory = $false)]
    [switch]$IgnoreDuplicateDataRowErrors,

    [Parameter(Mandatory = $false)]
    [string]$ClientId,

    [Parameter(Mandatory = $false)]
    [string]$Tenant,

    [Parameter(Mandatory = $false)]
    [string]$ConfigFile = "app-config.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Functions

function Ensure-PnPModule {
    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        throw "PnP.PowerShell module not found. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
    }
}

function Test-ConfigurationFile {
    param(
        [string]$ConfigFilePath
    )
    
    $fullPath = if ([System.IO.Path]::IsPathRooted($ConfigFilePath)) {
        $ConfigFilePath
    } else {
        Join-Path $PSScriptRoot $ConfigFilePath
    }
    
    if (-not (Test-Path $fullPath)) {
        Write-Host ""
        Write-Host "ERROR: Configuration file not found: $ConfigFilePath" -ForegroundColor Red
        Write-Host ""
        Write-Host "Required: Create a configuration file with your Azure AD app credentials." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Quick setup:" -ForegroundColor Cyan
        Write-Host "  1. Copy the sample file:" -ForegroundColor White
        Write-Host "     Copy-Item app-config.sample.json app-config.json" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  2. Edit app-config.json with your values:" -ForegroundColor White
        Write-Host "     - tenantId: Your Azure AD tenant ID" -ForegroundColor Gray
        Write-Host "     - clientId: Your app registration client ID" -ForegroundColor Gray
        Write-Host "     - certificateThumbprint: Your certificate thumbprint" -ForegroundColor Gray
        Write-Host "     - tenantDomain: Your tenant domain (e.g., contoso.onmicrosoft.com)" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  3. See MANUAL-APP-REGISTRATION.md for detailed setup instructions" -ForegroundColor White
        Write-Host ""
        return $false
    }
    
    try {
        $config = Get-Content $fullPath -Raw | ConvertFrom-Json
        
        # Validate required fields
        $missingFields = @()
        if (-not $config.tenantId) { $missingFields += "tenantId" }
        if (-not $config.clientId) { $missingFields += "clientId" }
        if (-not $config.tenantDomain) { $missingFields += "tenantDomain" }
        
        # Must have either certificate or client secret
        $hasAuth = $false
        if ($config.certificateThumbprint) { $hasAuth = $true }
        if ($config.clientSecret) { $hasAuth = $true }
        
        if ($missingFields.Count -gt 0) {
            Write-Host ""
            Write-Host "ERROR: Configuration file is incomplete" -ForegroundColor Red
            Write-Host "Missing required fields: $($missingFields -join ', ')" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Your config file must contain:" -ForegroundColor White
            Write-Host "  - tenantId (Azure AD tenant ID)" -ForegroundColor Gray
            Write-Host "  - clientId (App registration client ID)" -ForegroundColor Gray
            Write-Host "  - tenantDomain (e.g., contoso.onmicrosoft.com)" -ForegroundColor Gray
            Write-Host "  - certificateThumbprint OR clientSecret" -ForegroundColor Gray
            Write-Host ""
            return $false
        }
        
        if (-not $hasAuth) {
            Write-Host ""
            Write-Host "ERROR: No authentication method configured" -ForegroundColor Red
            Write-Host "You must provide either:" -ForegroundColor Yellow
            Write-Host "  - certificateThumbprint (recommended)" -ForegroundColor Gray
            Write-Host "  - clientSecret (fallback)" -ForegroundColor Gray
            Write-Host ""
            return $false
        }
        
        # If certificate specified, validate it exists
        if ($config.certificateThumbprint) {
            $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $config.certificateThumbprint }
            if (-not $cert) {
                Write-Host ""
                Write-Host "ERROR: Certificate not found" -ForegroundColor Red
                Write-Host "Thumbprint specified: $($config.certificateThumbprint)" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "Available certificates:" -ForegroundColor White
                Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint | Format-Table -AutoSize
                Write-Host ""
                Write-Host "To generate a new certificate, see MANUAL-APP-REGISTRATION.md" -ForegroundColor Cyan
                Write-Host ""
                return $false
            }
        }
        
        return $true
    }
    catch {
        Write-Host ""
        Write-Host "ERROR: Failed to read configuration file" -ForegroundColor Red
        Write-Host "File: $ConfigFilePath" -ForegroundColor Yellow
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Ensure the file contains valid JSON format." -ForegroundColor White
        Write-Host ""
        return $false
    }
}

function Test-TemplateFile {
    param(
        [string]$TemplatePath
    )
    
    if (-not (Test-Path $TemplatePath)) {
        Write-Host ""
        Write-Host "ERROR: Template file not found" -ForegroundColor Red
        Write-Host "Path specified: $TemplatePath" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Ensure you have exported a site template first using:" -ForegroundColor White
        Write-Host "  .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl 'URL' -IncludeContent" -ForegroundColor Gray
        Write-Host ""
        return $false
    }
    
    if ([System.IO.Path]::GetExtension($TemplatePath) -ne ".pnp") {
        Write-Host ""
        Write-Host "WARNING: Template file should have .pnp extension" -ForegroundColor Yellow
        Write-Host "File: $TemplatePath" -ForegroundColor Gray
        Write-Host ""
    }
    
    # Check file size
    $fileInfo = Get-Item $TemplatePath
    if ($fileInfo.Length -eq 0) {
        Write-Host ""
        Write-Host "ERROR: Template file is empty (0 bytes)" -ForegroundColor Red
        Write-Host "File: $TemplatePath" -ForegroundColor Yellow
        Write-Host ""
        return $false
    }
    
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    Write-ProgressMessage "Template file validated: $fileSizeMB MB" -Type "Success"
    
    return $true
}

function Connect-SharePoint {
    param(
        [string]$SiteUrl,
        [string]$ConfigFilePath,
        [string]$ClientIdParam,
        [string]$TenantParam
    )
    
    # Try app registration authentication first
    $configPath = Join-Path $PSScriptRoot $ConfigFilePath
    if (Test-Path $configPath) {
        Write-ProgressMessage "Loading configuration from: $ConfigFilePath" -Type "Info"
        try {
            $config = Get-Content $configPath | ConvertFrom-Json
            
            # Determine tenant value (prefer domain over ID for PnP)
            $tenantValue = if ($config.tenantDomain) { $config.tenantDomain } else { $config.tenantId }
            
            # Try certificate authentication first (more secure and uses modern auth)
            if ($config.clientId -and $config.certificateThumbprint -and $config.tenantId) {
                Write-ProgressMessage "Authenticating with App Registration + Certificate" -Type "Info"
                Write-ProgressMessage "Using modern authentication with certificate" -Type "Info"
                
                Connect-PnPOnline -Url $SiteUrl -ClientId $config.clientId -Thumbprint $config.certificateThumbprint -Tenant $config.tenantId -WarningAction SilentlyContinue
                
                Write-ProgressMessage "Connected using certificate authentication" -Type "Success"
                return
            }
            # Fall back to client secret (ACS)
            elseif ($config.clientId -and $config.clientSecret -and $tenantValue) {
                Write-ProgressMessage "Authenticating with App Registration (ClientId: $($config.clientId))" -Type "Info"
                Write-ProgressMessage "Using Azure Access Control Service (ACS) authentication" -Type "Info"
                
                # Note: PnP.PowerShell uses ACS auth with plain text client secret
                Connect-PnPOnline -Url $SiteUrl -ClientId $config.clientId -ClientSecret $config.clientSecret -WarningAction SilentlyContinue
                
                Write-ProgressMessage "Connected using app registration" -Type "Success"
                return
            }
        }
        catch {
            Write-ProgressMessage "App registration auth failed: $($_.Exception.Message)" -Type "Warning"
            Write-ProgressMessage "Falling back to interactive authentication..." -Type "Info"
        }
    }
    
    # Fall back to PnP Management Shell (built-in app with all permissions)
    Write-ProgressMessage "Using PnP Management Shell authentication" -Type "Info"
    Write-ProgressMessage "(This uses Microsoft's pre-registered app with full permissions)" -Type "Info"
    try {
        # PnP Management Shell is a Microsoft-registered app that bypasses most issues
        Connect-PnPOnline -Url $SiteUrl -Interactive -WarningAction SilentlyContinue
        Write-ProgressMessage "Connected using PnP Management Shell" -Type "Success"
    }
    catch {
        throw "Failed to connect: $($_.Exception.Message)"
    }
}

function Write-ProgressMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [string]$Type = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Type) {
        "Success" { Write-Host "[$timestamp] ✓ $Message" -ForegroundColor Green }
        "Warning" { Write-Host "[$timestamp] ⚠ $Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "[$timestamp] ✗ $Message" -ForegroundColor Red }
        default   { Write-Host "[$timestamp] ℹ $Message" -ForegroundColor Cyan }
    }
}

function Get-SiteInfo {
    try {
        $web = Get-PnPWeb -Includes Title, Description, ServerRelativeUrl, WebTemplate, Configuration -ErrorAction Stop
        $lists = Get-PnPList | Where-Object { -not $_.Hidden } | Measure-Object
        
        # Determine site type from template
        $templateId = "$($web.WebTemplate)#$($web.Configuration)"
        $siteType = switch -Wildcard ($templateId) {
            "SITEPAGEPUBLISHING#*" { "Communication Site" }
            "GROUP#*" { "Team Site (Microsoft 365 Group)" }
            "STS#*" { "Team Site (Classic)" }
            "TEAMCHANNEL#*" { "Teams Channel Site" }
            default { "$templateId" }
        }
        
        return @{
            Title = $web.Title
            Description = $web.Description
            Url = $web.ServerRelativeUrl
            ListCount = $lists.Count
            TemplateId = $templateId
            SiteType = $siteType
            Exists = $true
        }
    }
    catch {
        return @{
            Exists = $false
            Error = $_.Exception.Message
        }
    }
}

function Show-PreImportWarning {
    param(
        [string]$TargetSite,
        [string]$Template
    )
    
    Write-Host ""
    Write-Host "╔═══════════════════════════════════════════════════════╗" -ForegroundColor Yellow
    Write-Host "║              ⚠  IMPORTANT WARNING ⚠                  ║" -ForegroundColor Yellow
    Write-Host "╚═══════════════════════════════════════════════════════╝" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  You are about to apply a template to:" -ForegroundColor Yellow
    Write-Host "  Target Site: $TargetSite" -ForegroundColor White
    Write-Host ""
    Write-Host "  Template: $Template" -ForegroundColor White
    Write-Host ""
    Write-Host "  This operation will:" -ForegroundColor Yellow
    Write-Host "    • Modify site structure and configuration" -ForegroundColor White
    Write-Host "    • Create new lists, libraries, and pages" -ForegroundColor White
    Write-Host "    • Import content and data" -ForegroundColor White
    Write-Host "    • Potentially overwrite existing items" -ForegroundColor White
    Write-Host ""
    Write-Host "  Recommendations:" -ForegroundColor Yellow
    Write-Host "    ✓ Backup your site first" -ForegroundColor Green
    Write-Host "    ✓ Test in non-production environment" -ForegroundColor Green
    Write-Host "    ✓ Verify you have Site Collection Admin rights" -ForegroundColor Green
    Write-Host ""
}

#endregion

#region Main Script

try {
    Write-ProgressMessage "Starting SharePoint site template import" -Type "Info"
    
    # Ensure PnP module is available
    Ensure-PnPModule
    
    # Validate configuration before proceeding
    Write-ProgressMessage "Validating configuration..." -Type "Info"
    if (-not (Test-ConfigurationFile -ConfigFilePath $ConfigFile)) {
        Write-Host "Import aborted due to configuration errors." -ForegroundColor Red
        exit 1
    }
    Write-ProgressMessage "Configuration validated successfully" -Type "Success"
    
    # Validate template file before proceeding
    Write-ProgressMessage "Validating template file..." -Type "Info"
    if (-not (Test-TemplateFile -TemplatePath $TemplatePath)) {
        Write-Host "Import aborted due to template file errors." -ForegroundColor Red
        exit 1
    }
    
    # Verify template file exists
    $templateFile = Get-Item $TemplatePath
    $templateSizeMB = [math]::Round($templateFile.Length / 1MB, 2)
    
    Write-ProgressMessage "Template file: $($templateFile.Name) ($templateSizeMB MB)" -Type "Info"
    
    # Create log directory and file
    $logDirectory = Join-Path (Split-Path $TemplatePath) "ImportLogs"
    if (-not (Test-Path -Path $logDirectory)) {
        New-Item -ItemType Directory -Path $logDirectory | Out-Null
    }
    
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $logPath = Join-Path $logDirectory "Import_$timestamp.log"
    
    # Start transcript
    Start-Transcript -Path $logPath
    
    # Show warning if not WhatIf
    if (-not $WhatIfPreference) {
        Show-PreImportWarning -TargetSite $TargetSiteUrl -Template $templateFile.Name
        
        $confirmation = Read-Host "Do you want to proceed? (yes/no)"
        if ($confirmation -ne 'yes') {
            Write-ProgressMessage "Import cancelled by user" -Type "Warning"
            return
        }
    }
    
    # Connect to SharePoint
    Write-ProgressMessage "Connecting to target site: $TargetSiteUrl" -Type "Info"
    
    Connect-SharePoint -SiteUrl $TargetSiteUrl -ConfigFilePath $ConfigFile -ClientIdParam $ClientId -TenantParam $Tenant
    
    # Check if target site exists
    Write-ProgressMessage "Checking if target site exists..." -Type "Info"
    $siteInfoBefore = Get-SiteInfo
    
    if (-not $siteInfoBefore.Exists) {
        # Extract site details for instructions
        $siteAlias = if ($TargetSiteUrl -match '/sites/([^/]+)') { $matches[1] } else { "YourSite" }
        $siteName = $siteAlias -replace '-', ' '
        $tenantName = if ($TargetSiteUrl -match 'https://([^\.]+)\.sharepoint\.com') { $matches[1] } else { "tenant" }
        $adminUrl = "https://$tenantName-admin.sharepoint.com"
        
        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
        Write-Host "  ✗ TARGET SITE DOES NOT EXIST" -ForegroundColor Red
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
        Write-Host ""
        Write-Host "The target site must be created before importing the template." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "MANUAL SITE CREATION INSTRUCTIONS:" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Option 1: SharePoint Admin Center (Recommended)" -ForegroundColor Green
        Write-Host "  1. Go to: $adminUrl/_layouts/15/online/SiteCollections.aspx" -ForegroundColor White
        Write-Host "  2. Click 'Create' → 'Team site' or 'Communication site'" -ForegroundColor White
        Write-Host "  3. Site name: '$siteName'" -ForegroundColor White
        Write-Host "  4. Site address: $siteAlias" -ForegroundColor White
        Write-Host "  5. Add site owner/admin" -ForegroundColor White
        Write-Host "  6. Click 'Finish' and wait for provisioning" -ForegroundColor White
        Write-Host ""
        Write-Host "Option 2: PowerShell with Interactive Login" -ForegroundColor Green
        Write-Host "  Connect-PnPOnline -Url '$adminUrl' -Interactive" -ForegroundColor White
        Write-Host "  New-PnPSite -Type TeamSite -Title '$siteName' ``" -ForegroundColor White
        Write-Host "    -Alias '$siteAlias' -Wait" -ForegroundColor White
        Write-Host "  Disconnect-PnPOnline" -ForegroundColor White
        Write-Host ""
        Write-Host "After creating the site, re-run this import script:" -ForegroundColor Cyan
        Write-Host "  .\Import-SharePointSiteTemplate.ps1 ``" -ForegroundColor White
        Write-Host "    -TargetSiteUrl '$TargetSiteUrl' ``" -ForegroundColor White
        Write-Host "    -TemplatePath '$TemplatePath'" -ForegroundColor White
        Write-Host ""
        throw "Target site does not exist: $TargetSiteUrl"
    }
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Target Site Information (Before Import)" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Title:      $($siteInfoBefore.Title)" -ForegroundColor White
    Write-Host "  URL:        $($siteInfoBefore.Url)" -ForegroundColor White
    Write-Host "  Site Type:  $($siteInfoBefore.SiteType)" -ForegroundColor Yellow
    Write-Host "  Template:   $($siteInfoBefore.TemplateId)" -ForegroundColor Gray
    Write-Host "  Lists/Libs: $($siteInfoBefore.ListCount)" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    # Build parameters for template application
    $invokeParams = @{
        Path = $TemplatePath
    }
    
    if ($ClearNavigation) {
        $invokeParams['ClearNavigation'] = $true
        Write-ProgressMessage "Navigation will be cleared" -Type "Warning"
    }
    
    if ($OverwriteSystemPropertyBagValues) {
        $invokeParams['OverwriteSystemPropertyBagValues'] = $true
    }
    
    if ($ProvisionFieldsToSite) {
        $invokeParams['ProvisionFieldsStoSite'] = $true
    }
    
    # Apply template
    if ($PSCmdlet.ShouldProcess($TargetSiteUrl, "Apply PnP template")) {
        Write-ProgressMessage "Applying template to site..." -Type "Info"
        Write-ProgressMessage "This may take several minutes depending on template size..." -Type "Info"
        
        $startTime = Get-Date
        
        Invoke-PnPSiteTemplate @invokeParams
        
        $endTime = Get-Date
        $duration = $endTime - $startTime
        
        Write-ProgressMessage "Template applied successfully in $($duration.TotalMinutes.ToString('0.00')) minutes" -Type "Success"
        
        # Get updated site information
        Write-ProgressMessage "Gathering updated site information..." -Type "Info"
        $siteInfoAfter = Get-SiteInfo
        
        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host "  Import Completed Successfully!" -ForegroundColor Green
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host ""
        Write-Host "  Target Site Information (After Import)" -ForegroundColor Cyan
        Write-Host "  ─────────────────────────────────────────────────────" -ForegroundColor Cyan
        Write-Host "  Title:      $($siteInfoAfter.Title)" -ForegroundColor White
        Write-Host "  URL:        $($siteInfoAfter.Url)" -ForegroundColor White
        Write-Host "  Lists/Libs: $($siteInfoAfter.ListCount) (was $($siteInfoBefore.ListCount))" -ForegroundColor White
        Write-Host ""
        Write-Host "  Duration:   $($duration.TotalMinutes.ToString('0.00')) minutes" -ForegroundColor White
        Write-Host "  Log File:   $logPath" -ForegroundColor White
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
        Write-Host ""
        
        Write-ProgressMessage "Next steps:" -Type "Info"
        Write-Host "  1. Verify site structure and content" -ForegroundColor White
        Write-Host "  2. Test site functionality" -ForegroundColor White
        Write-Host "  3. Review permissions and security" -ForegroundColor White
        Write-Host "  4. Notify users when ready for Go Live" -ForegroundColor White
        Write-Host ""
    }
    else {
        Write-ProgressMessage "WhatIf: Would apply template to $TargetSiteUrl" -Type "Info"
    }
}
catch {
    Write-ProgressMessage "Import failed: $($_.Exception.Message)" -Type "Error"
    Write-Host ""
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Host ""
    Write-ProgressMessage "Check log file for details: $logPath" -Type "Error"
    throw
}
finally {
    Stop-Transcript
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

#endregion
