<#
.SYNOPSIS
    Permanently removes SharePoint sites from the recycle bin (deleted sites).

.DESCRIPTION
    This script connects to SharePoint tenant admin center and permanently deletes
    sites that are in the recycle bin. Use with caution - this action is irreversible.

.PARAMETER SiteUrl
    The URL of the site to permanently delete. Must be in the recycle bin.

.PARAMETER ListAll
    Lists all sites currently in the recycle bin without deleting.

.PARAMETER DeleteAll
    Permanently deletes ALL sites in the recycle bin. Requires confirmation.

.PARAMETER TenantAdminUrl
    SharePoint Admin Center URL. Auto-detected from app-config.json if not provided.

.PARAMETER ConfigFile
    Path to app-config.json file. Default: app-config.json in script directory.

.PARAMETER Force
    Skip confirmation prompts. Use with extreme caution.

.EXAMPLE
    .\Remove-DeletedSharePointSite.ps1 -ListAll

.EXAMPLE
    .\Remove-DeletedSharePointSite.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/OldSite"

.EXAMPLE
    .\Remove-DeletedSharePointSite.ps1 -DeleteAll -Force

.NOTES
    Author: IT Support
    Date: February 3, 2026
    Requires: PnP.PowerShell module
    
    WARNING: This permanently deletes sites. They cannot be recovered.
#>

[CmdletBinding(DefaultParameterSetName = 'List')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'DeleteOne')]
    [ValidatePattern('^https://[^/]+\.sharepoint\.com/.*$')]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'List')]
    [switch]$ListAll,

    [Parameter(Mandatory = $true, ParameterSetName = 'DeleteAll')]
    [switch]$DeleteAll,

    [Parameter(Mandatory = $false)]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [string]$ConfigFile = "app-config.json",

    [Parameter(Mandatory = $false)]
    [switch]$Force
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

function Connect-AdminCenter {
    param(
        [string]$AdminUrl,
        [string]$ConfigFilePath
    )
    
    # First validate the configuration file
    if (-not (Test-ConfigurationFile -ConfigFilePath $ConfigFilePath)) {
        throw "Configuration validation failed. Cannot proceed with connection."
    }
    
    $configPath = Join-Path $PSScriptRoot $ConfigFilePath
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ Loading configuration..." -ForegroundColor Cyan
    $config = Get-Content $configPath | ConvertFrom-Json
    
    # Determine admin URL if not provided
    if (-not $AdminUrl) {
        if ($config.tenantDomain) {
            $tenantName = $config.tenantDomain.Split('.')[0]
            $AdminUrl = "https://$tenantName-admin.sharepoint.com"
        } else {
            throw "Cannot determine admin URL. Please provide -TenantAdminUrl parameter or add tenantDomain to config file."
        }
    }
    
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ Connecting to: $AdminUrl" -ForegroundColor Cyan
    
    # Try certificate authentication first
    if ($config.clientId -and $config.certificateThumbprint -and $config.tenantId) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ Using certificate authentication" -ForegroundColor Cyan
        
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $config.certificateThumbprint }
        if (-not $cert) {
            throw "Certificate not found with thumbprint: $($config.certificateThumbprint)"
        }
        
        Connect-PnPOnline -Url $AdminUrl `
            -ClientId $config.clientId `
            -Thumbprint $config.certificateThumbprint `
            -Tenant $config.tenantId
        
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ✓ Connected using certificate" -ForegroundColor Green
    }
    elseif ($config.clientId -and $config.clientSecret) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ Using client secret authentication" -ForegroundColor Cyan
        
        Connect-PnPOnline -Url $AdminUrl `
            -ClientId $config.clientId `
            -ClientSecret $config.clientSecret
        
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ✓ Connected using client secret" -ForegroundColor Green
    }
    else {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ Using interactive authentication" -ForegroundColor Cyan
        Connect-PnPOnline -Url $AdminUrl -Interactive
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ✓ Connected interactively" -ForegroundColor Green
    }
}

function Get-DeletedSites {
    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ Retrieving deleted sites..." -ForegroundColor Cyan
    
    $deletedSites = Get-PnPTenantDeletedSite
    
    if ($deletedSites.Count -eq 0) {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ✓ No sites in recycle bin" -ForegroundColor Green
        return @()
    }
    
    return $deletedSites
}

function Show-DeletedSites {
    param([array]$Sites)
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Deleted Sites (Recycle Bin)" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Sites.Count -eq 0) {
        Write-Host "  No sites in recycle bin" -ForegroundColor Green
    } else {
        Write-Host "  Found $($Sites.Count) deleted site(s):" -ForegroundColor Yellow
        Write-Host ""
        
        foreach ($site in $Sites) {
            Write-Host "  • $($site.Url)" -ForegroundColor White
            Write-Host "    Title:        $($site.Title)" -ForegroundColor Gray
            Write-Host "    Deleted:      $($site.DeletionTime)" -ForegroundColor Gray
            Write-Host "    Days in bin:  $((New-TimeSpan -Start $site.DeletionTime -End (Get-Date)).Days)" -ForegroundColor Gray
            Write-Host ""
        }
    }
    
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
}

function Remove-SitePermanently {
    param(
        [string]$Url,
        [switch]$SkipConfirmation
    )
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ⚠ WARNING: PERMANENT DELETION" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host ""
    Write-Host "  Site: $Url" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  This action will:" -ForegroundColor Yellow
    Write-Host "    • PERMANENTLY delete the site" -ForegroundColor White
    Write-Host "    • Remove ALL content and data" -ForegroundColor White
    Write-Host "    • Cannot be undone or recovered" -ForegroundColor White
    Write-Host ""
    
    if (-not $SkipConfirmation) {
        $confirmation = Read-Host "Type 'DELETE' to confirm permanent deletion"
        if ($confirmation -ne 'DELETE') {
            Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ⚠ Operation cancelled" -ForegroundColor Yellow
            return $false
        }
    }
    
    try {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ Permanently deleting site..." -ForegroundColor Cyan
        
        Remove-PnPTenantDeletedSite -Identity $Url -Force
        
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ✓ Site permanently deleted" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ✗ Failed to delete site: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

#endregion

#region Main Script

try {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  SharePoint Deleted Site Cleanup" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    # Ensure PnP module is available
    Ensure-PnPModule
    
    # Connect to admin center
    Connect-AdminCenter -AdminUrl $TenantAdminUrl -ConfigFilePath $ConfigFile
    
    # Get deleted sites
    $deletedSites = Get-DeletedSites
    
    switch ($PSCmdlet.ParameterSetName) {
        'List' {
            Show-DeletedSites -Sites $deletedSites
        }
        
        'DeleteOne' {
            # Check if site exists in recycle bin
            $site = $deletedSites | Where-Object { $_.Url -eq $SiteUrl }
            
            if (-not $site) {
                Write-Host ""
                Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
                Write-Host "  ✗ SITE NOT FOUND IN RECYCLE BIN" -ForegroundColor Red
                Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
                Write-Host ""
                Write-Host "  Site: $SiteUrl" -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Possible reasons:" -ForegroundColor Cyan
                Write-Host "    • Site is not deleted (still active)" -ForegroundColor White
                Write-Host "    • Site was already permanently deleted" -ForegroundColor White
                Write-Host "    • Site URL is incorrect" -ForegroundColor White
                Write-Host ""
                Write-Host "  Use -ListAll to see sites in recycle bin:" -ForegroundColor Cyan
                Write-Host "    .\Remove-DeletedSharePointSite.ps1 -ListAll" -ForegroundColor White
                Write-Host ""
                
                throw "Site not found in recycle bin"
            }
            
            $success = Remove-SitePermanently -Url $SiteUrl -SkipConfirmation:$Force
            
            if ($success) {
                Write-Host ""
                Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
                Write-Host "  ✓ Cleanup Completed" -ForegroundColor Green
                Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
                Write-Host ""
            }
        }
        
        'DeleteAll' {
            if ($deletedSites.Count -eq 0) {
                Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ℹ No sites to delete" -ForegroundColor Cyan
                return
            }
            
            Show-DeletedSites -Sites $deletedSites
            
            Write-Host ""
            Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
            Write-Host "  ⚠ WARNING: BULK PERMANENT DELETION" -ForegroundColor Red
            Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
            Write-Host ""
            Write-Host "  About to permanently delete $($deletedSites.Count) site(s)" -ForegroundColor Yellow
            Write-Host "  This action CANNOT be undone!" -ForegroundColor Yellow
            Write-Host ""
            
            if (-not $Force) {
                $confirmation = Read-Host "Type 'DELETE ALL' to confirm"
                if ($confirmation -ne 'DELETE ALL') {
                    Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ⚠ Operation cancelled" -ForegroundColor Yellow
                    return
                }
            }
            
            $successCount = 0
            $failCount = 0
            
            foreach ($site in $deletedSites) {
                $success = Remove-SitePermanently -Url $site.Url -SkipConfirmation
                if ($success) {
                    $successCount++
                } else {
                    $failCount++
                }
            }
            
            Write-Host ""
            Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
            Write-Host "  Bulk Deletion Completed" -ForegroundColor Green
            Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
            Write-Host "  Successfully deleted: $successCount" -ForegroundColor Green
            if ($failCount -gt 0) {
                Write-Host "  Failed:              $failCount" -ForegroundColor Red
            }
            Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
            Write-Host ""
        }
    }
}
catch {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host "  ✗ Operation Failed" -ForegroundColor Red
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Red
    Write-Host ""
    Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Host ""
    throw
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}

#endregion
