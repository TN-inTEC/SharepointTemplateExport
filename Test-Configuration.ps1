<#
.SYNOPSIS
    Validates SharePoint template configuration files

.DESCRIPTION
    Tests configuration files for completeness and validity before running
    export/import operations. Checks for required fields, certificate availability,
    and provides helpful guidance for fixing issues.

.PARAMETER ConfigFile
    Path to the configuration file to validate. Default: app-config.json

.EXAMPLE
    .\Test-Configuration.ps1

.EXAMPLE
    .\Test-Configuration.ps1 -ConfigFile "app-config-target.json"

.NOTES
    Author: IT Support
    Date: February 3, 2026
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$ConfigFile = "app-config.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Write-ColorMessage {
    param(
        [string]$Message,
        [string]$Type = "Info"
    )
    
    $color = switch ($Type) {
        "Success" { "Green" }
        "Warning" { "Yellow" }
        "Error" { "Red" }
        "Info" { "Cyan" }
        default { "White" }
    }
    
    Write-Host $Message -ForegroundColor $color
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
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  Configuration Validation" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Testing: $ConfigFilePath" -ForegroundColor White
    Write-Host ""
    
    # Test 1: File exists
    Write-Host "Test 1: Configuration file exists..." -NoNewline
    if (-not (Test-Path $fullPath)) {
        Write-Host " ✗ FAIL" -ForegroundColor Red
        Write-Host ""
        Write-Host "ERROR: Configuration file not found" -ForegroundColor Red
        Write-Host "Expected location: $fullPath" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Quick fix:" -ForegroundColor Cyan
        Write-Host "  Copy-Item app-config.sample.json $ConfigFilePath" -ForegroundColor Gray
        Write-Host ""
        return $false
    }
    Write-Host " ✓ PASS" -ForegroundColor Green
    
    # Test 2: Valid JSON
    Write-Host "Test 2: Valid JSON format..." -NoNewline
    try {
        $config = Get-Content $fullPath -Raw | ConvertFrom-Json
        Write-Host " ✓ PASS" -ForegroundColor Green
    }
    catch {
        Write-Host " ✗ FAIL" -ForegroundColor Red
        Write-Host ""
        Write-Host "ERROR: Invalid JSON format" -ForegroundColor Red
        Write-Host "Error: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host ""
        return $false
    }
    
    # Test 3: Required fields present
    Write-Host "Test 3: Required fields present..." -NoNewline
    $missingFields = @()
    if (-not $config.tenantId) { $missingFields += "tenantId" }
    if (-not $config.clientId) { $missingFields += "clientId" }
    if (-not $config.tenantDomain) { $missingFields += "tenantDomain" }
    
    if ($missingFields.Count -gt 0) {
        Write-Host " ✗ FAIL" -ForegroundColor Red
        Write-Host ""
        Write-Host "ERROR: Missing required fields" -ForegroundColor Red
        Write-Host "Missing: $($missingFields -join ', ')" -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Required fields:" -ForegroundColor Cyan
        Write-Host "  - tenantId: Azure AD tenant ID (GUID)" -ForegroundColor Gray
        Write-Host "  - clientId: App registration client ID (GUID)" -ForegroundColor Gray
        Write-Host "  - tenantDomain: Tenant domain (e.g., contoso.onmicrosoft.com)" -ForegroundColor Gray
        Write-Host ""
        return $false
    }
    Write-Host " ✓ PASS" -ForegroundColor Green
    
    # Test 4: Authentication method configured
    Write-Host "Test 4: Authentication method configured..." -NoNewline
    $hasAuth = $false
    $authMethod = ""
    if ($config.certificateThumbprint) { 
        $hasAuth = $true 
        $authMethod = "Certificate"
    }
    if ($config.clientSecret) { 
        $hasAuth = $true 
        if ($authMethod) { $authMethod += " + " }
        $authMethod += "Client Secret"
    }
    
    if (-not $hasAuth) {
        Write-Host " ✗ FAIL" -ForegroundColor Red
        Write-Host ""
        Write-Host "ERROR: No authentication method configured" -ForegroundColor Red
        Write-Host "You must provide either:" -ForegroundColor Yellow
        Write-Host "  - certificateThumbprint (recommended)" -ForegroundColor Gray
        Write-Host "  - clientSecret (fallback)" -ForegroundColor Gray
        Write-Host ""
        return $false
    }
    Write-Host " ✓ PASS ($authMethod)" -ForegroundColor Green
    
    # Test 5: Certificate exists (if specified)
    if ($config.certificateThumbprint) {
        Write-Host "Test 5: Certificate available..." -NoNewline
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $config.certificateThumbprint }
        if (-not $cert) {
            Write-Host " ✗ FAIL" -ForegroundColor Red
            Write-Host ""
            Write-Host "ERROR: Certificate not found" -ForegroundColor Red
            Write-Host "Thumbprint: $($config.certificateThumbprint)" -ForegroundColor Yellow
            Write-Host ""
            Write-Host "Available certificates:" -ForegroundColor Cyan
            $allCerts = Get-ChildItem Cert:\CurrentUser\My
            if ($allCerts.Count -eq 0) {
                Write-Host "  (None found)" -ForegroundColor Gray
            } else {
                $allCerts | Select-Object Subject, Thumbprint, NotAfter | Format-Table -AutoSize
            }
            Write-Host ""
            Write-Host "To generate a certificate, see MANUAL-APP-REGISTRATION.md" -ForegroundColor Cyan
            Write-Host ""
            return $false
        }
        
        # Check certificate expiration
        $daysUntilExpiry = ($cert.NotAfter - (Get-Date)).Days
        if ($daysUntilExpiry -lt 0) {
            Write-Host " ✗ FAIL (Expired)" -ForegroundColor Red
            Write-Host ""
            Write-Host "ERROR: Certificate has expired" -ForegroundColor Red
            Write-Host "Expired on: $($cert.NotAfter)" -ForegroundColor Yellow
            Write-Host ""
            return $false
        }
        elseif ($daysUntilExpiry -lt 30) {
            Write-Host " ⚠ PASS (Expires soon)" -ForegroundColor Yellow
            Write-Host "    WARNING: Certificate expires in $daysUntilExpiry days ($($cert.NotAfter))" -ForegroundColor Yellow
        }
        else {
            Write-Host " ✓ PASS (Expires: $($cert.NotAfter))" -ForegroundColor Green
        }
        
        # Test 6: Certificate has private key
        Write-Host "Test 6: Certificate has private key..." -NoNewline
        if (-not $cert.HasPrivateKey) {
            Write-Host " ✗ FAIL" -ForegroundColor Red
            Write-Host ""
            Write-Host "ERROR: Certificate does not have a private key" -ForegroundColor Red
            Write-Host "This certificate cannot be used for authentication." -ForegroundColor Yellow
            Write-Host ""
            return $false
        }
        Write-Host " ✓ PASS" -ForegroundColor Green
    }
    
    # Test 7: GUID format validation
    Write-Host "Test 7: Valid GUID formats..." -NoNewline
    try {
        [void][System.Guid]::Parse($config.tenantId)
        [void][System.Guid]::Parse($config.clientId)
        Write-Host " ✓ PASS" -ForegroundColor Green
    }
    catch {
        Write-Host " ✗ FAIL" -ForegroundColor Red
        Write-Host ""
        Write-Host "ERROR: Invalid GUID format" -ForegroundColor Red
        Write-Host "tenantId and clientId must be valid GUIDs" -ForegroundColor Yellow
        Write-Host ""
        return $false
    }
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host "  ✓ All tests passed!" -ForegroundColor Green
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Green
    Write-Host ""
    Write-Host "Configuration Summary:" -ForegroundColor Cyan
    Write-Host "  Tenant Domain: $($config.tenantDomain)" -ForegroundColor White
    Write-Host "  Tenant ID:     $($config.tenantId)" -ForegroundColor Gray
    Write-Host "  Client ID:     $($config.clientId)" -ForegroundColor Gray
    Write-Host "  Auth Method:   $authMethod" -ForegroundColor White
    Write-Host ""
    Write-Host "This configuration is ready to use with:" -ForegroundColor Green
    Write-Host "  .\Export-SharePointSiteTemplate.ps1 -ConfigFile '$ConfigFilePath'" -ForegroundColor Gray
    Write-Host "  .\Import-SharePointSiteTemplate.ps1 -ConfigFile '$ConfigFilePath'" -ForegroundColor Gray
    Write-Host "  .\Remove-DeletedSharePointSite.ps1 -ConfigFile '$ConfigFilePath'" -ForegroundColor Gray
    Write-Host ""
    
    return $true
}

# Main execution
try {
    $result = Test-ConfigurationFile -ConfigFilePath $ConfigFile
    
    if (-not $result) {
        Write-Host "Configuration validation failed." -ForegroundColor Red
        Write-Host "See CONFIG-README.md and MANUAL-APP-REGISTRATION.md for setup guidance." -ForegroundColor Yellow
        Write-Host ""
        exit 1
    }
    
    exit 0
}
catch {
    Write-Host ""
    Write-Host "ERROR: Validation failed with exception" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Yellow
    Write-Host ""
    exit 1
}
