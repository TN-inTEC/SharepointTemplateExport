<#
.SYNOPSIS
    Exports a complete SharePoint site as a PnP template with content for migration.

.DESCRIPTION
    This script exports an entire SharePoint site including structure, configuration,
    pages, lists, libraries, and optionally content data. Designed for site migrations
    like SLM Academy to SMC SharePoint deployments.
    
    For cross-tenant migrations, use New-UserMappingTemplate.ps1 to generate a user
    mapping file after exporting the template.

.PARAMETER SourceSiteUrl
    The URL of the SharePoint site to export.

.PARAMETER OutputPath
    Path where the template file will be saved. Default: C:\PSReports\SiteTemplates

.PARAMETER TemplateName
    Name for the template file (without extension). Default: SiteTemplate_<timestamp>

.PARAMETER IncludeContent
    Include list and library items in the template.

.PARAMETER ContentRowLimit
    Maximum number of rows to export per list when including content. Default: 5000

.PARAMETER ExcludeHandlers
    Handlers to exclude from export (comma-separated). 
    Options: AuditSettings, ComposedLook, CustomActions, ExtensibilityProviders, Features, 
    Fields, Files, Lists, Pages, Publishing, RegionalSettings, SearchSettings, 
    SitePolicy, SupportedUILanguages, TermGroups, Workflows, etc.

.PARAMETER IncludeLists
    Specific list/library titles to include in export. If specified, only these lists will be exported.
    Use with -IncludeContent to export data from these lists only.

.PARAMETER ExcludeLists
    Specific list/library titles to exclude from export.

.PARAMETER ExcludePages
    Exclude site pages from the export.

.PARAMETER StructureOnly
    Export only list/library structure without any content data. Overrides -IncludeContent.

.PARAMETER Preview
    Preview what will be exported without actually creating the template file.

.PARAMETER ClientId
    Azure AD App Client ID for authentication. Uses PnP Management Shell if not specified.

.PARAMETER Tenant
    Tenant name (e.g., contoso) for authentication.

.EXAMPLE
    .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/SLMAcademy" -IncludeContent

.EXAMPLE
    .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/SLMAcademy" `
        -TemplateName "SLM_Academy_Full" -IncludeContent -ContentRowLimit 10000

.EXAMPLE
    .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectSite" `
        -OutputPath "C:\Templates" -IncludeLists "Documents","Tasks","Project Tracker" -IncludeContent

.EXAMPLE
    .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectSite" `
        -OutputPath "C:\Templates" -ExcludeLists "Archive","Old Documents" -IncludeContent

.EXAMPLE
    .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectSite" `
        -OutputPath "C:\Templates" -StructureOnly

.EXAMPLE
    .\Export-SharePointSiteTemplate.ps1 -SourceSiteUrl "https://contoso.sharepoint.com/sites/ProjectSite" `
        -OutputPath "C:\Templates" -Preview

.NOTES
    Author: IT Support
    Date: February 3, 2026
    Requires: PnP.PowerShell module
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[^/]+\.sharepoint\.com/.*$')]
    [string]$SourceSiteUrl,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$OutputPath = "C:\PSReports\SiteTemplates",

    [Parameter(Mandatory = $false)]
    [string]$TemplateName,

    [Parameter(Mandatory = $false)]
    [switch]$IncludeContent,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 10000)]
    [int]$ContentRowLimit = 5000,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeHandlers,

    [Parameter(Mandatory = $false)]
    [string[]]$IncludeLists,

    [Parameter(Mandatory = $false)]
    [string[]]$ExcludeLists,

    [Parameter(Mandatory = $false)]
    [switch]$ExcludePages,

    [Parameter(Mandatory = $false)]
    [switch]$StructureOnly,

    [Parameter(Mandatory = $false)]
    [switch]$Preview,

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
    $web = Get-PnPWeb -Includes Title, Description, ServerRelativeUrl, Created, LastItemModifiedDate, WebTemplate, Configuration
    $lists = Get-PnPList | Where-Object { -not $_.Hidden } | Measure-Object
    $pages = Get-PnPPage | Measure-Object
    
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
        Created = $web.Created
        LastModified = $web.LastItemModifiedDate
        ListCount = $lists.Count
        PageCount = $pages.Count
        TemplateId = $templateId
        SiteType = $siteType
    }
}

#endregion

#region Main Script

try {
    Write-ProgressMessage "Starting SharePoint site template export" -Type "Info"
    
    # Ensure PnP module is available
    Ensure-PnPModule
    
    # Validate configuration before proceeding
    Write-ProgressMessage "Validating configuration..." -Type "Info"
    if (-not (Test-ConfigurationFile -ConfigFilePath $ConfigFile)) {
        Write-Host "Export aborted due to configuration errors." -ForegroundColor Red
        exit 1
    }
    Write-ProgressMessage "Configuration validated successfully" -Type "Success"
    
    # Create output directory if it doesn't exist
    if (-not (Test-Path -Path $OutputPath)) {
        New-Item -ItemType Directory -Path $OutputPath | Out-Null
        Write-ProgressMessage "Created output directory: $OutputPath" -Type "Success"
    }
    
    # Generate template name if not provided
    if ([string]::IsNullOrWhiteSpace($TemplateName)) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $TemplateName = "SiteTemplate_$timestamp"
    }
    
    $templatePath = Join-Path $OutputPath "$TemplateName.pnp"
    $logPath = Join-Path $OutputPath "$TemplateName.log"
    
    # Start transcript
    Start-Transcript -Path $logPath
    
    # Connect to SharePoint
    Write-ProgressMessage "Connecting to SharePoint site: $SourceSiteUrl" -Type "Info"
    
    Connect-SharePoint -SiteUrl $SourceSiteUrl -ConfigFilePath $ConfigFile -ClientIdParam $ClientId -TenantParam $Tenant
    
    # Get site information
    Write-ProgressMessage "Gathering site information..." -Type "Info"
    $siteInfo = Get-SiteInfo
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Site Information" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Title:         $($siteInfo.Title)" -ForegroundColor White
    Write-Host "  Description:   $($siteInfo.Description)" -ForegroundColor White
    Write-Host "  URL:           $($siteInfo.Url)" -ForegroundColor White
    Write-Host "  Site Type:     $($siteInfo.SiteType)" -ForegroundColor Yellow
    Write-Host "  Template ID:   $($siteInfo.TemplateId)" -ForegroundColor Gray
    Write-Host "  Created:       $($siteInfo.Created)" -ForegroundColor White
    Write-Host "  Last Modified: $($siteInfo.LastModified)" -ForegroundColor White
    Write-Host "  Lists/Libs:    $($siteInfo.ListCount)" -ForegroundColor White
    Write-Host "  Pages:         $($siteInfo.PageCount)" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    # Validate selective export parameters
    if ($IncludeLists -and $ExcludeLists) {
        Write-Host "ERROR: Cannot specify both -IncludeLists and -ExcludeLists" -ForegroundColor Red
        Write-Host "Choose one approach: whitelist (IncludeLists) or blacklist (ExcludeLists)" -ForegroundColor Yellow
        Stop-Transcript
        Disconnect-PnPOnline
        exit 1
    }
    
    # Get all lists for filtering logic
    $allLists = Get-PnPList | Where-Object { -not $_.Hidden }
    $listsToExport = $allLists
    
    # Apply list filtering
    if ($IncludeLists) {
        $listsToExport = $allLists | Where-Object { $IncludeLists -contains $_.Title }
        $notFoundLists = $IncludeLists | Where-Object { $_ -notin $listsToExport.Title }
        
        if ($notFoundLists) {
            Write-ProgressMessage "WARNING: The following lists were not found: $($notFoundLists -join ', ')" -Type "Warning"
        }
        
        Write-ProgressMessage "Filtered to $($listsToExport.Count) lists: $($listsToExport.Title -join ', ')" -Type "Info"
    }
    elseif ($ExcludeLists) {
        $listsToExport = $allLists | Where-Object { $ExcludeLists -notcontains $_.Title }
        Write-ProgressMessage "Excluding $($ExcludeLists.Count) lists, exporting $($listsToExport.Count)" -Type "Info"
    }
    
    # Preview mode - show what would be exported
    if ($Preview) {
        Write-Host ""
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
        Write-Host "  PREVIEW MODE - No Template Will Be Created" -ForegroundColor Magenta
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
        Write-Host ""
        
        Write-Host "Export Configuration:" -ForegroundColor Cyan
        Write-Host "  Template Name:  $TemplateName" -ForegroundColor White
        Write-Host "  Include Content: $(if ($IncludeContent -and -not $StructureOnly) { 'Yes' } else { 'No' })" -ForegroundColor White
        Write-Host "  Structure Only:  $(if ($StructureOnly) { 'Yes' } else { 'No' })" -ForegroundColor White
        Write-Host "  Exclude Pages:   $(if ($ExcludePages) { 'Yes' } else { 'No' })" -ForegroundColor White
        Write-Host "  Row Limit:       $ContentRowLimit" -ForegroundColor White
        Write-Host ""
        
        Write-Host "Lists/Libraries to Export ($($listsToExport.Count)):" -ForegroundColor Cyan
        foreach ($list in $listsToExport | Sort-Object Title) {
            $baseType = switch ($list.BaseType) {
                "GenericList" { "List" }
                "DocumentLibrary" { "Library" }
                default { $list.BaseType }
            }
            $contentIndicator = if ($list.ItemCount -gt 0) { " ($($list.ItemCount) items)" } else { " (empty)" }
            Write-Host "  • $($list.Title) [$baseType]$contentIndicator" -ForegroundColor White
        }
        Write-Host ""
        
        if (-not $ExcludePages) {
            $pages = Get-PnPListItem -List "Site Pages" -PageSize 500 | Where-Object { $_["FileLeafRef"] -like "*.aspx" }
            Write-Host "Pages to Export ($($pages.Count)):" -ForegroundColor Cyan
            foreach ($page in $pages | Select-Object -First 20) {
                Write-Host "  • $($page["FileLeafRef"])" -ForegroundColor White
            }
            if ($pages.Count -gt 20) {
                Write-Host "  ... and $($pages.Count - 20) more pages" -ForegroundColor Gray
            }
            Write-Host ""
        }
        
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
        Write-Host "Preview complete. Use without -Preview to perform export." -ForegroundColor Magenta
        Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
        Write-Host ""
        
        Stop-Transcript
        Disconnect-PnPOnline
        exit 0
    }
    
    # Build handlers parameter
    $handlersParam = @{}
    $handlersToInclude = @()
    
    # Default handlers to include
    $defaultHandlers = @('Lists', 'Fields', 'ContentTypes', 'CustomActions', 'Features', 'Navigation', 'SiteSettings')
    
    if (-not $ExcludePages) {
        $defaultHandlers += 'Pages'
    }
    else {
        Write-ProgressMessage "Pages will be excluded from export" -Type "Warning"
    }
    
    if ($ExcludeHandlers -and $ExcludeHandlers.Count -gt 0) {
        Write-ProgressMessage "Excluding handlers: $($ExcludeHandlers -join ', ')" -Type "Warning"
        $handlersToInclude = $defaultHandlers | Where-Object { $ExcludeHandlers -notcontains $_ }
        $handlersParam['Handlers'] = $handlersToInclude -join ','
    }
    else {
        if ($ExcludePages) {
            $handlersParam['Handlers'] = $defaultHandlers -join ','
        }
        else {
            $handlersParam['Handlers'] = 'All'
        }
    }
    
    # Export site template
    Write-ProgressMessage "Exporting site template..." -Type "Info"
    Write-ProgressMessage "Handlers: $($handlersParam['Handlers'])" -Type "Info"
    
    $exportParams = @{
        Out = $templatePath
    }
    $exportParams += $handlersParam
    
    # If we're filtering lists, we need to export them individually
    if ($IncludeLists -or $ExcludeLists) {
        Write-ProgressMessage "Using selective list export..." -Type "Info"
        
        # First export base template without lists
        $baseHandlers = $handlersParam['Handlers'] -replace ',?Lists,?', ''
        $exportParams['Handlers'] = $baseHandlers
        Get-PnPSiteTemplate @exportParams
        
        # Then add selected lists one by one
        foreach ($list in $listsToExport) {
            try {
                Write-ProgressMessage "Adding list schema: $($list.Title)" -Type "Info"
                Add-PnPListFoldersToSiteTemplate -Path $templatePath -List $list.Title -Recursive -ErrorAction Continue
            }
            catch {
                Write-ProgressMessage "Failed to add list '$($list.Title)': $($_.Exception.Message)" -Type "Error"
            }
        }
    }
    else {
        Get-PnPSiteTemplate @exportParams
    }
    
    Write-ProgressMessage "Template structure exported successfully" -Type "Success"
    
    # Add content if requested (and not StructureOnly)
    if ($IncludeContent -and -not $StructureOnly) {
        Write-ProgressMessage "Adding content data to template (RowLimit: $ContentRowLimit)..." -Type "Info"
        
        # Filter to lists with content
        $listsWithContent = $listsToExport | Where-Object { $_.ItemCount -gt 0 }
        
        if ($listsWithContent.Count -eq 0) {
            Write-ProgressMessage "No lists with content found to export" -Type "Warning"
        }
        else {
            Write-ProgressMessage "Found $($listsWithContent.Count) lists with content" -Type "Info"
            
            foreach ($list in $listsWithContent) {
                try {
                    $listTitle = $list.Title
                    $itemCount = $list.ItemCount
                    
                    if ($itemCount -gt $ContentRowLimit) {
                        Write-ProgressMessage "List '$listTitle' has $itemCount items. Only first $ContentRowLimit will be exported." -Type "Warning"
                    }
                    
                    Write-ProgressMessage "Exporting content from: $listTitle ($itemCount items)" -Type "Info"
                    
                    $query = "<View><RowLimit>$ContentRowLimit</RowLimit></View>"
                    Add-PnPDataRowsToSiteTemplate -Path $templatePath -List $listTitle -Query $query -ErrorAction Continue
                    
                    Write-ProgressMessage "  └─ Completed: $listTitle" -Type "Success"
                }
                catch {
                    Write-ProgressMessage "  └─ Failed to export content from '$listTitle': $($_.Exception.Message)" -Type "Error"
                }
            }
        }
    }
    elseif ($StructureOnly) {
        Write-ProgressMessage "Structure-only export - no content data will be included" -Type "Warning"
    }
    
    # Get file size
    $fileInfo = Get-Item $templatePath
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  Export Completed Successfully!" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  Template File: $templatePath" -ForegroundColor White
    Write-Host "  File Size:     $fileSizeMB MB" -ForegroundColor White
    Write-Host "  Log File:      $logPath" -ForegroundColor White
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    
    Write-Host "⚠ IMPORTANT for Import:" -ForegroundColor Yellow
    Write-Host "  Source site type: $($siteInfo.SiteType)" -ForegroundColor Yellow
    Write-Host "  Create target site with the SAME type to avoid errors!" -ForegroundColor Yellow
    Write-Host ""
    
    Write-Host "📋 Next Steps:" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  For SAME-TENANT migrations:" -ForegroundColor White
    Write-Host "    .\Import-SharePointSiteTemplate.ps1 ``" -ForegroundColor Gray
    Write-Host "      -TargetSiteUrl 'https://tenant.sharepoint.com/sites/target' ``" -ForegroundColor Gray
    Write-Host "      -TemplatePath '$templatePath'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "  For CROSS-TENANT migrations:" -ForegroundColor White
    Write-Host "    1. Generate user mapping template:" -ForegroundColor Gray
    Write-Host "       .\New-UserMappingTemplate.ps1 -TemplatePath '$templatePath'" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    2. Edit user-mapping-template.csv with target tenant emails" -ForegroundColor Gray
    Write-Host ""
    Write-Host "    3. Import with user mapping:" -ForegroundColor Gray
    Write-Host "       .\Import-SharePointSiteTemplate.ps1 ``" -ForegroundColor Gray
    Write-Host "         -TargetSiteUrl 'https://targettenant.sharepoint.com/sites/target' ``" -ForegroundColor Gray
    Write-Host "         -TemplatePath '$templatePath' ``" -ForegroundColor Gray
    Write-Host "         -UserMappingFile 'user-mapping-template.csv' ``" -ForegroundColor Gray
    Write-Host "         -IgnoreDuplicateDataRowErrors" -ForegroundColor Gray
    Write-Host ""
    
    Write-ProgressMessage "Template ready for import to target site" -Type "Success"
}
catch {
    Write-ProgressMessage "Export failed: $($_.Exception.Message)" -Type "Error"
    Write-Host ""
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    throw
}
finally {
    Stop-Transcript
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

#endregion
