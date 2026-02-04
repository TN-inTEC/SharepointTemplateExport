<#
.SYNOPSIS
    Registers an Azure AD application for SharePoint site template operations.

.DESCRIPTION
    This script creates and configures an Azure AD application registration with the necessary
    permissions for SharePoint site template export and import operations. The app will be
    configured with appropriate API permissions for SharePoint Online access.

.PARAMETER AppName
    Name for the Azure AD application. Default: "SharePoint Site Template Manager"

.PARAMETER Tenant
    The tenant name (e.g., contoso) or full domain (contoso.onmicrosoft.com).

.PARAMETER AddCertificate
    If specified, creates and uploads a self-signed certificate for certificate-based authentication.
    This is more secure than client secrets.

.PARAMETER CertificateYears
    Number of years the certificate should be valid. Default: 2

.PARAMETER AddClientSecret
    If specified, creates a client secret for the application.

.PARAMETER SecretYears
    Number of years the client secret should be valid. Default: 2

.PARAMETER ExportPath
    Path where app registration details will be saved. Default: C:\PSReports\AppRegistrations

.PARAMETER GrantAdminConsent
    Automatically grant admin consent for the requested permissions (requires Global Admin).

.EXAMPLE
    .\Register-SharePointApp.ps1 -Tenant "contoso" -AddCertificate

.EXAMPLE
    .\Register-SharePointApp.ps1 -Tenant "contoso.onmicrosoft.com" -AddClientSecret -GrantAdminConsent

.EXAMPLE
    .\Register-SharePointApp.ps1 -AppName "SP Template Tool" -Tenant "contoso" -AddCertificate -CertificateYears 3

.NOTES
    Author: IT Support
    Date: February 3, 2026
    Requires: Microsoft.Graph.Applications module and appropriate Azure AD permissions
    
    Required Permissions:
    - Application.ReadWrite.All (to create app registration)
    - AppRoleAssignment.ReadWrite.All (to grant admin consent)
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$AppName = "SharePoint Site Template Manager",
    
    [Parameter(Mandatory = $true)]
    [string]$Tenant,
    
    [Parameter(Mandatory = $false)]
    [switch]$AddCertificate,
    
    [Parameter(Mandatory = $false)]
    [int]$CertificateYears = 2,
    
    [Parameter(Mandatory = $false)]
    [switch]$AddClientSecret,
    
    [Parameter(Mandatory = $false)]
    [int]$SecretYears = 2,
    
    [Parameter(Mandatory = $false)]
    [string]$ExportPath = "C:\PSReports\AppRegistrations",
    
    [Parameter(Mandatory = $false)]
    [switch]$GrantAdminConsent
)

# Ensure tenant is in correct format
if ($Tenant -notlike "*.onmicrosoft.com") {
    $TenantDomain = "$Tenant.onmicrosoft.com"
} else {
    $TenantDomain = $Tenant
    $Tenant = $Tenant.Split('.')[0]
}

$ErrorActionPreference = 'Stop'

# Create output directory
if (-not (Test-Path $ExportPath)) {
    New-Item -Path $ExportPath -ItemType Directory -Force | Out-Null
}

$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$outputFile = Join-Path $ExportPath "AppRegistration_$($AppName -replace '[^\w]','')_$timestamp.txt"

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "SharePoint App Registration Script" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Check for required module
Write-Host "Checking for Microsoft.Graph modules..." -ForegroundColor Yellow
$requiredModules = @(
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Applications'
)

foreach ($module in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $module)) {
        Write-Host "Installing $module..." -ForegroundColor Yellow
        try {
            Install-Module -Name $module -Scope CurrentUser -Force -AllowClobber
            Write-Host "✓ $module installed successfully" -ForegroundColor Green
        }
        catch {
            Write-Host "✗ Failed to install $module. Please run: Install-Module $module" -ForegroundColor Red
            exit 1
        }
    }
}

# Connect to Microsoft Graph
Write-Host ""
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
Write-Host "  Note: This requires administrative permissions in your tenant." -ForegroundColor Gray
Write-Host ""

# Check if already connected
$existingConnection = Get-MgContext -ErrorAction SilentlyContinue
if ($existingConnection) {
    Write-Host "✓ Already connected to Microsoft Graph" -ForegroundColor Green
    Write-Host "  Account: $($existingConnection.Account)" -ForegroundColor Gray
    Write-Host "  Tenant: $($existingConnection.TenantId)" -ForegroundColor Gray
    
    # Verify scopes
    $hasRequiredScopes = $existingConnection.Scopes -contains "Application.ReadWrite.All"
    if (-not $hasRequiredScopes) {
        Write-Host "  Warning: Current connection may not have required permissions. Reconnecting..." -ForegroundColor Yellow
        Disconnect-MgGraph -ErrorAction SilentlyContinue
        $existingConnection = $null
    }
}

if (-not $existingConnection) {
    # Disable WAM (Web Account Manager) to use traditional browser flow
    $env:AZURE_IDENTITY_DISABLE_WAM = "true"
    
    # Try multiple authentication methods
    $connected = $false
    $authMethods = @(
        @{Name="Device Code"; Params=@{TenantId=$TenantDomain; Scopes=@("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"); UseDeviceAuthentication=$true; ContextScope="Process"}},
        @{Name="Interactive Browser"; Params=@{TenantId=$TenantDomain; Scopes=@("Application.ReadWrite.All", "AppRoleAssignment.ReadWrite.All"); ContextScope="Process"}}
    )
    
    foreach ($method in $authMethods) {
        try {
            Write-Host "  Trying $($method.Name) authentication..." -ForegroundColor Yellow
            if ($method.Name -like "*Device Code*") {
                Write-Host "  (You'll receive a code to enter at https://microsoft.com/devicelogin)" -ForegroundColor DarkGray
            }
            
            Connect-MgGraph @($method.Params) -NoWelcome -ErrorAction Stop
            $connected = $true
            Write-Host "✓ Connected to Microsoft Graph using $($method.Name)" -ForegroundColor Green
            break
        }
        catch {
            $errorMsg = $_.Exception.Message
            Write-Host "  ✗ $($method.Name) failed" -ForegroundColor Red
            
            # Check for specific error codes
            if ($errorMsg -like "*53003*" -or $errorMsg -like "*conditional access*") {
                Write-Host ""
                Write-Host "  ERROR 53003: Conditional Access Policy Blocking Authentication" -ForegroundColor Red
                Write-Host ""
                Write-Host "  This error means your organization's security policies are blocking this app." -ForegroundColor Yellow
                Write-Host ""
                Write-Host "  Required Actions (contact your IT Admin):" -ForegroundColor White
                Write-Host "  1. Your device may need to be registered with Azure AD" -ForegroundColor Cyan
                Write-Host "  2. Conditional Access policies may need adjustment for:" -ForegroundColor Cyan
                Write-Host "     - App: Microsoft Graph Command Line Tools" -ForegroundColor Gray
                Write-Host "     - App ID: 14d82eec-204b-4c2f-b7e8-296a70dab67e" -ForegroundColor Gray
                Write-Host "  3. Your device may need to be marked as compliant/trusted" -ForegroundColor Cyan
                Write-Host ""
                Write-Host "  Alternative Solutions:" -ForegroundColor White
                Write-Host "  A. Use Azure Cloud Shell (has no device restrictions)" -ForegroundColor Cyan
                Write-Host "  B. Use a registered/compliant device" -ForegroundColor Cyan
                Write-Host "  C. Ask admin to create exception for Graph Command Line Tools app" -ForegroundColor Cyan
                Write-Host "  D. Use Azure Portal to create app registration manually" -ForegroundColor Cyan
                Write-Host ""
                exit 1
            } else {
                Write-Host "  Details: $errorMsg" -ForegroundColor DarkGray
            }
            Write-Host ""
        }
    }
    
    if (-not $connected) {
        Write-Host ""
        Write-Host "✗ Failed to connect to Microsoft Graph with all authentication methods" -ForegroundColor Red
        Write-Host ""
        Write-Host "Common Issues:" -ForegroundColor Yellow
        Write-Host "1. Conditional Access Policies (Error 53003)" -ForegroundColor White
        Write-Host "   - Device not registered/compliant with organization policies" -ForegroundColor Gray
        Write-Host "   - Solution: Contact IT admin or use Azure Cloud Shell" -ForegroundColor Gray
        Write-Host ""
        Write-Host "2. Insufficient Permissions" -ForegroundColor White
        Write-Host "   - You need Global Admin or Application Admin role" -ForegroundColor Gray
        Write-Host "   - Solution: Request appropriate permissions from IT admin" -ForegroundColor Gray
        Write-Host ""
        Write-Host "Alternative: Create App Registration via Azure Portal" -ForegroundColor Cyan
        Write-Host "  https://portal.azure.com/#view/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/~/RegisteredApps" -ForegroundColor Cyan
        Write-Host ""
        exit 1
    }
}

try {
    Write-Host ""
    Write-Host "Creating Azure AD application..." -ForegroundColor Yellow
    
    # Define required resource access (API permissions)
    # SharePoint Online API
    $sharepointResourceId = "00000003-0000-0ff1-ce00-000000000000" # Office 365 SharePoint Online
    
    # Permissions needed for PnP template operations
    $requiredResourceAccess = @{
        ResourceAppId = $sharepointResourceId
        ResourceAccess = @(
            @{
                # Sites.FullControl.All - Full control of all site collections
                Id = "678536fe-1083-478a-9c59-b99265e6b0d3"
                Type = "Role" # Application permission
            },
            @{
                # Sites.Read.All - Read items in all site collections
                Id = "332a536c-c7ef-4017-ab91-336970924f0d"
                Type = "Role" # Application permission
            },
            @{
                # Sites.Manage.All - Create, edit, and delete items and lists in all site collections
                Id = "0c0bf378-bf22-4481-8f81-9e89a9b4960a"
                Type = "Role" # Application permission
            }
        )
    }
    
    # Create the application registration
    $appParams = @{
        DisplayName = $AppName
        SignInAudience = "AzureADMyOrg"
        RequiredResourceAccess = @($requiredResourceAccess)
        Web = @{
            RedirectUris = @("http://localhost")
        }
    }
    
    $app = New-MgApplication -BodyParameter $appParams
    Write-Host "✓ Application created: $($app.DisplayName)" -ForegroundColor Green
    Write-Host "  Application (Client) ID: $($app.AppId)" -ForegroundColor Cyan
    Write-Host "  Object ID: $($app.Id)" -ForegroundColor Cyan
    
    # Create service principal for the app
    Write-Host ""
    Write-Host "Creating service principal..." -ForegroundColor Yellow
    $sp = New-MgServicePrincipal -AppId $app.AppId
    Write-Host "✓ Service principal created" -ForegroundColor Green
    
    # Store app details
    $appDetails = @{
        AppName = $app.DisplayName
        ApplicationId = $app.AppId
        ObjectId = $app.Id
        TenantId = $TenantDomain
        CreatedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    }
    
    # Handle certificate authentication
    if ($AddCertificate) {
        Write-Host ""
        Write-Host "Creating self-signed certificate..." -ForegroundColor Yellow
        
        $certName = "CN=$($AppName -replace '[^\w\s]','')"
        $certPath = Join-Path $ExportPath "$($AppName -replace '[^\w]','')_Certificate.pfx"
        $cerPath = Join-Path $ExportPath "$($AppName -replace '[^\w]','')_Certificate.cer"
        
        # Generate a secure password for the certificate
        $certPassword = -join ((48..57) + (65..90) + (97..122) | Get-Random -Count 20 | ForEach-Object {[char]$_})
        $securePassword = ConvertTo-SecureString -String $certPassword -Force -AsPlainText
        
        # Create the certificate
        $cert = New-SelfSignedCertificate -Subject $certName `
            -CertStoreLocation "Cert:\CurrentUser\My" `
            -KeyExportPolicy Exportable `
            -KeySpec Signature `
            -KeyLength 2048 `
            -KeyAlgorithm RSA `
            -HashAlgorithm SHA256 `
            -NotAfter (Get-Date).AddYears($CertificateYears)
        
        # Export certificate
        Export-PfxCertificate -Cert $cert -FilePath $certPath -Password $securePassword | Out-Null
        Export-Certificate -Cert $cert -FilePath $cerPath | Out-Null
        
        Write-Host "✓ Certificate created" -ForegroundColor Green
        Write-Host "  Thumbprint: $($cert.Thumbprint)" -ForegroundColor Cyan
        Write-Host "  Expires: $($cert.NotAfter.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
        
        # Upload certificate to app registration
        Write-Host ""
        Write-Host "Uploading certificate to app registration..." -ForegroundColor Yellow
        
        $certBase64 = [Convert]::ToBase64String($cert.GetRawCertData())
        
        Update-MgApplication -ApplicationId $app.Id -KeyCredentials @(
            @{
                Type = "AsymmetricX509Cert"
                Usage = "Verify"
                Key = [System.Convert]::FromBase64String($certBase64)
            }
        )
        
        Write-Host "✓ Certificate uploaded to application" -ForegroundColor Green
        
        $appDetails.CertificateThumbprint = $cert.Thumbprint
        $appDetails.CertificateExpiry = $cert.NotAfter.ToString('yyyy-MM-dd')
        $appDetails.CertificatePath = $certPath
        $appDetails.CertificatePassword = $certPassword
        $appDetails.PublicCertPath = $cerPath
    }
    
    # Handle client secret
    if ($AddClientSecret) {
        Write-Host ""
        Write-Host "Creating client secret..." -ForegroundColor Yellow
        
        $secretParams = @{
            PasswordCredential = @{
                DisplayName = "Auto-generated secret"
                EndDateTime = (Get-Date).AddYears($SecretYears)
            }
        }
        
        $secret = Add-MgApplicationPassword -ApplicationId $app.Id -BodyParameter $secretParams
        
        Write-Host "✓ Client secret created" -ForegroundColor Green
        Write-Host "  Secret expires: $($secret.EndDateTime.ToString('yyyy-MM-dd'))" -ForegroundColor Cyan
        
        $appDetails.ClientSecret = $secret.SecretText
        $appDetails.SecretExpiry = $secret.EndDateTime.ToString('yyyy-MM-dd')
    }
    
    # Grant admin consent if requested
    if ($GrantAdminConsent) {
        Write-Host ""
        Write-Host "Granting admin consent for API permissions..." -ForegroundColor Yellow
        Write-Host "  (This may take a moment...)" -ForegroundColor Gray
        
        # Wait a moment for the service principal to be fully provisioned
        Start-Sleep -Seconds 5
        
        try {
            # Get SharePoint service principal
            $sharepointSP = Get-MgServicePrincipal -Filter "appId eq '$sharepointResourceId'"
            
            # Grant each permission
            foreach ($permission in $requiredResourceAccess.ResourceAccess) {
                $grantParams = @{
                    PrincipalId = $sp.Id
                    ResourceId = $sharepointSP.Id
                    AppRoleId = $permission.Id
                }
                
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $sp.Id -BodyParameter $grantParams | Out-Null
            }
            
            Write-Host "✓ Admin consent granted" -ForegroundColor Green
            $appDetails.AdminConsentGranted = $true
        }
        catch {
            Write-Host "⚠ Failed to grant admin consent automatically: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Host "  You'll need to grant consent manually in the Azure Portal" -ForegroundColor Yellow
            $appDetails.AdminConsentGranted = $false
        }
    }
    else {
        $appDetails.AdminConsentGranted = $false
    }
    
    # Save app details to file
    Write-Host ""
    Write-Host "Saving app registration details..." -ForegroundColor Yellow
    
    $output = @"
========================================
SharePoint App Registration Details
========================================
Created: $($appDetails.CreatedDate)

Application Name: $($appDetails.AppName)
Application (Client) ID: $($appDetails.ApplicationId)
Object ID: $($appDetails.ObjectId)
Tenant ID: $($appDetails.TenantId)

"@
    
    if ($appDetails.CertificateThumbprint) {
        $output += @"
Certificate Authentication:
------------------------------------------
Certificate Thumbprint: $($appDetails.CertificateThumbprint)
Certificate Expiry: $($appDetails.CertificateExpiry)
PFX Certificate Path: $($appDetails.CertificatePath)
Certificate Password: $($appDetails.CertificatePassword)
Public Certificate Path: $($appDetails.PublicCertPath)

IMPORTANT: Store the certificate password securely!

"@
    }
    
    if ($appDetails.ClientSecret) {
        $output += @"
Client Secret Authentication:
------------------------------------------
Client Secret: $($appDetails.ClientSecret)
Secret Expiry: $($appDetails.SecretExpiry)

IMPORTANT: Store the client secret securely! You won't be able to retrieve it again.

"@
    }
    
    $output += @"
API Permissions Configured:
------------------------------------------
• Sites.FullControl.All (Application)
• Sites.Read.All (Application)
• Sites.Manage.All (Application)

Admin Consent Status: $($appDetails.AdminConsentGranted)

"@
    
    if (-not $appDetails.AdminConsentGranted) {
        $output += @"
Manual Admin Consent Required:
------------------------------------------
1. Go to Azure Portal > Azure Active Directory > App registrations
2. Find and select "$($appDetails.AppName)"
3. Click "API permissions"
4. Click "Grant admin consent for $Tenant"

"@
    }
    
    $output += @"

Usage in PnP PowerShell Scripts:
------------------------------------------
"@
    
    if ($appDetails.CertificateThumbprint) {
        $output += @"
# Using Certificate Authentication:
Connect-PnPOnline -Url "https://$Tenant.sharepoint.com/sites/YourSite" ``
    -ClientId "$($appDetails.ApplicationId)" ``
    -Tenant "$($appDetails.TenantId)" ``
    -Thumbprint "$($appDetails.CertificateThumbprint)"

"@
    }
    
    if ($appDetails.ClientSecret) {
        $output += @"
# Using Client Secret Authentication:
`$secureSecret = ConvertTo-SecureString -String "$($appDetails.ClientSecret)" -AsPlainText -Force
Connect-PnPOnline -Url "https://$Tenant.sharepoint.com/sites/YourSite" ``
    -ClientId "$($appDetails.ApplicationId)" ``
    -Tenant "$($appDetails.TenantId)" ``
    -ClientSecret `$secureSecret

"@
    }
    
    $output += @"

Next Steps:
------------------------------------------
1. If not done automatically, grant admin consent in Azure Portal
2. Update your Export and Import scripts with the ClientId
3. Use either certificate or client secret for authentication
4. Test the connection using the PnP PowerShell commands above

For more information:
https://pnp.github.io/powershell/articles/authentication.html
========================================
"@
    
    $output | Out-File -FilePath $outputFile -Encoding UTF8
    Write-Host "✓ Details saved to: $outputFile" -ForegroundColor Green
    
    # Display summary
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host "App Registration Complete!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "Application (Client) ID: $($appDetails.ApplicationId)" -ForegroundColor Cyan
    Write-Host "Tenant: $($appDetails.TenantId)" -ForegroundColor Cyan
    
    if ($appDetails.CertificateThumbprint) {
        Write-Host "Certificate Thumbprint: $($appDetails.CertificateThumbprint)" -ForegroundColor Cyan
    }
    
    Write-Host ""
    Write-Host "IMPORTANT:" -ForegroundColor Yellow
    Write-Host "• Full details saved to: $outputFile" -ForegroundColor White
    
    if ($appDetails.ClientSecret) {
        Write-Host "• Client secret is shown only once - save it securely!" -ForegroundColor White
    }
    
    if ($appDetails.CertificatePassword) {
        Write-Host "• Certificate password is saved in the details file" -ForegroundColor White
    }
    
    if (-not $appDetails.AdminConsentGranted) {
        Write-Host "• Grant admin consent in Azure Portal before using the app" -ForegroundColor White
    }
    
    Write-Host ""
    Write-Host "Use this Client ID in your Export and Import scripts:" -ForegroundColor Yellow
    Write-Host "  -ClientId `"$($appDetails.ApplicationId)`"" -ForegroundColor Cyan
    Write-Host ""
    
}
catch {
    Write-Host ""
    Write-Host "✗ Error occurred: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "Stack Trace:" -ForegroundColor Gray
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
    exit 1
}
finally {
    # Disconnect from Microsoft Graph
    Disconnect-MgGraph | Out-Null
}
