<#
.SYNOPSIS
    Validates SharePoint template configuration files

.DESCRIPTION
    Tests configuration files for completeness and validity before running
    export/import operations. Checks for required fields, certificate availability,
    and provides helpful guidance for fixing issues.
    
    Supports validating multiple configuration files for cross-tenant migrations.

.PARAMETER ConfigFile
    Path to the configuration file to validate. Default: app-config.json

.PARAMETER SourceConfigFile
    Path to the source tenant configuration file for cross-tenant validation.
    If provided with -TargetConfigFile, validates both configurations.

.PARAMETER TargetConfigFile
    Path to the target tenant configuration file for cross-tenant validation.
    Must be used with -SourceConfigFile.

.EXAMPLE
    .\Test-Configuration.ps1
    
    Validates the default app-config.json

.EXAMPLE
    .\Test-Configuration.ps1 -ConfigFile "app-config-target.json"
    
    Validates a specific configuration file

.EXAMPLE
    .\Test-Configuration.ps1 -SourceConfigFile "app-config-source.json" -TargetConfigFile "app-config-target.json"
    
    Validates both source and target configurations for cross-tenant migration

.NOTES
    Author: IT Support
    Date: February 4, 2026
#>

[CmdletBinding(DefaultParameterSetName = "Single")]
param(
    [Parameter(Mandatory = $false, ParameterSetName = "Single")]
    [string]$ConfigFile = "app-config.json",
    
    [Parameter(Mandatory = $true, ParameterSetName = "CrossTenant")]
    [string]$SourceConfigFile,
    
    [Parameter(Mandatory = $true, ParameterSetName = "CrossTenant")]
    [string]$TargetConfigFile
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
        [string]$ConfigFilePath,
        [string]$Label = ""
    )
    
    $fullPath = if ([System.IO.Path]::IsPathRooted($ConfigFilePath)) {
        $ConfigFilePath
    } else {
        Join-Path $PSScriptRoot $ConfigFilePath
    }
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    if ($Label) {
        Write-Host "  Configuration Validation - $Label" -ForegroundColor Cyan
    } else {
        Write-Host "  Configuration Validation" -ForegroundColor Cyan
    }
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
        return @{ Success = $false }
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
        return @{ Success = $false }
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
        return @{ Success = $false }
    }
    Write-Host " ✓ PASS" -ForegroundColor Green
    
    # Test 4: Authentication method configured
    Write-Host "Test 4: Authentication method configured..." -NoNewline
    $hasAuth = $false
    $authMethod = ""
    if ($config.PSObject.Properties['certificateThumbprint'] -and $config.certificateThumbprint) { 
        $hasAuth = $true 
        $authMethod = "Certificate"
    }
    if ($config.PSObject.Properties['clientSecret'] -and $config.clientSecret) { 
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
        return @{ Success = $false }
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
            return @{ Success = $false }
        }
        
        # Check certificate expiration
        $daysUntilExpiry = ($cert.NotAfter - (Get-Date)).Days
        if ($daysUntilExpiry -lt 0) {
            Write-Host " ✗ FAIL (Expired)" -ForegroundColor Red
            Write-Host ""
            Write-Host "ERROR: Certificate has expired" -ForegroundColor Red
            Write-Host "Expired on: $($cert.NotAfter)" -ForegroundColor Yellow
            Write-Host ""
            return @{ Success = $false }
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
            return @{ Success = $false }
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
        return @{ Success = $false }
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
    
    if (-not $Label) {
        Write-Host "This configuration is ready to use with:" -ForegroundColor Green
        Write-Host "  .\Export-SharePointSiteTemplate.ps1 -ConfigFile '$ConfigFilePath'" -ForegroundColor Gray
        Write-Host "  .\Import-SharePointSiteTemplate.ps1 -ConfigFile '$ConfigFilePath'" -ForegroundColor Gray
        Write-Host "  .\Remove-DeletedSharePointSite.ps1 -ConfigFile '$ConfigFilePath'" -ForegroundColor Gray
        Write-Host ""
    }
    
    return @{
        Success = $true
        TenantDomain = $config.tenantDomain
        TenantId = $config.tenantId
        ClientId = $config.clientId
        AuthMethod = $authMethod
    }
}

# Main execution
try {
    if ($PSCmdlet.ParameterSetName -eq "CrossTenant") {
        # Validate both source and target configurations
        Write-Host ""
        Write-Host "╔═══════════════════════════════════════════════════════╗" -ForegroundColor Magenta
        Write-Host "║     Cross-Tenant Migration Configuration Test        ║" -ForegroundColor Magenta
        Write-Host "╚═══════════════════════════════════════════════════════╝" -ForegroundColor Magenta
        Write-Host ""
        
        $sourceResult = Test-ConfigurationFile -ConfigFilePath $SourceConfigFile -Label "SOURCE TENANT"
        
        if (-not $sourceResult.Success) {
            Write-Host ""
            Write-Host "❌ Source configuration validation failed." -ForegroundColor Red
            Write-Host "Fix the source configuration before proceeding." -ForegroundColor Yellow
            Write-Host ""
            exit 1
        }
        
        $targetResult = Test-ConfigurationFile -ConfigFilePath $TargetConfigFile -Label "TARGET TENANT"
        
        if (-not $targetResult.Success) {
            Write-Host ""
            Write-Host "❌ Target configuration validation failed." -ForegroundColor Red
            Write-Host "Fix the target configuration before proceeding." -ForegroundColor Yellow
            Write-Host ""
            exit 1
        }
        
        # Display cross-tenant summary
        Write-Host ""
        Write-Host "╔═══════════════════════════════════════════════════════╗" -ForegroundColor Green
        Write-Host "║   ✓ Both Configurations Valid - Ready for Migration  ║" -ForegroundColor Green
        Write-Host "╚═══════════════════════════════════════════════════════╝" -ForegroundColor Green
        Write-Host ""
        Write-Host "Cross-Tenant Migration Summary:" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "  SOURCE TENANT:" -ForegroundColor Yellow
        Write-Host "    Domain:      $($sourceResult.TenantDomain)" -ForegroundColor White
        Write-Host "    Tenant ID:   $($sourceResult.TenantId)" -ForegroundColor Gray
        Write-Host "    Auth:        $($sourceResult.AuthMethod)" -ForegroundColor White
        Write-Host ""
        Write-Host "  TARGET TENANT:" -ForegroundColor Yellow
        Write-Host "    Domain:      $($targetResult.TenantDomain)" -ForegroundColor White
        Write-Host "    Tenant ID:   $($targetResult.TenantId)" -ForegroundColor Gray
        Write-Host "    Auth:        $($targetResult.AuthMethod)" -ForegroundColor White
        Write-Host ""
        
        # Warn if same tenant
        if ($sourceResult.TenantId -eq $targetResult.TenantId) {
            Write-Host "  ⚠️  WARNING: Source and Target are the SAME tenant!" -ForegroundColor Yellow
            Write-Host "      This is a same-tenant migration." -ForegroundColor Yellow
            Write-Host ""
        }
        
        Write-Host "Ready to use with:" -ForegroundColor Green
        Write-Host "  # 1. Export from source" -ForegroundColor Gray
        Write-Host "  .\Export-SharePointSiteTemplate.ps1 -ConfigFile '$SourceConfigFile' -IncludeContent" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  # 2. Generate user mapping (for cross-tenant)" -ForegroundColor Gray
        Write-Host "  .\New-UserMappingTemplate.ps1 -TemplatePath 'template.pnp'" -ForegroundColor Gray
        Write-Host ""
        Write-Host "  # 3. Import to target" -ForegroundColor Gray
        Write-Host "  .\Import-SharePointSiteTemplate.ps1 -ConfigFile '$TargetConfigFile' -TemplatePath 'template.pnp' -UserMappingFile 'user-mapping.csv'" -ForegroundColor Gray
        Write-Host ""
        
    } else {
        # Single configuration validation
        $result = Test-ConfigurationFile -ConfigFilePath $ConfigFile
        
        if (-not $result.Success) {
            Write-Host "Configuration validation failed." -ForegroundColor Red
            Write-Host "See CONFIG-README.md and MANUAL-APP-REGISTRATION.md for setup guidance." -ForegroundColor Yellow
            Write-Host ""
            exit 1
        }
        
        Write-Host "💡 TIP: For cross-tenant migrations, use:" -ForegroundColor Cyan
        Write-Host "  .\Test-Configuration.ps1 -SourceConfigFile 'app-config-source.json' -TargetConfigFile 'app-config-target.json'" -ForegroundColor Gray
        Write-Host ""
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
