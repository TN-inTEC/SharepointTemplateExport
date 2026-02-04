<#
.SYNOPSIS
    Generates a user mapping template CSV from a SharePoint site template export.

.DESCRIPTION
    Analyzes a PnP site template (.pnp or .xml) or connects to a live SharePoint site
    to extract all unique user references. Creates a CSV template for mapping users
    during cross-tenant migrations.

.PARAMETER TemplatePath
    Path to the exported PnP template file (.pnp or .xml) to analyze.

.PARAMETER SiteUrl
    URL of a SharePoint site to scan for users (alternative to template file).

.PARAMETER OutputPath
    Path where the user mapping template CSV will be saved.
    Default: user-mapping-template.csv in the current directory.

.PARAMETER IncludeSystemAccounts
    Include system accounts (SharePoint App, System Account) in the output.

.PARAMETER ConfigFile
    Path to app-config.json file for authentication (when using -SiteUrl).

.EXAMPLE
    .\New-UserMappingTemplate.ps1 -TemplatePath "C:\PSReports\SiteTemplates\SLM_Academy.pnp"

.EXAMPLE
    .\New-UserMappingTemplate.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/Academy" -OutputPath "C:\Temp\users.csv"

.NOTES
    Author: IT Support
    Date: February 4, 2026
    Requires: PnP.PowerShell module (when using -SiteUrl)
#>

[CmdletBinding(DefaultParameterSetName = 'FromTemplate')]
param(
    [Parameter(Mandatory = $true, ParameterSetName = 'FromTemplate')]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Template file not found: $_"
        }
        return $true
    })]
    [string]$TemplatePath,

    [Parameter(Mandatory = $true, ParameterSetName = 'FromSite')]
    [ValidatePattern('^https://[^/]+\.sharepoint\.com/.*$')]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "user-mapping-template.csv",

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSystemAccounts,

    [Parameter(Mandatory = $false, ParameterSetName = 'FromSite')]
    [string]$ConfigFile = "app-config.json"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Helper Functions

function Write-ProgressMessage {
    param(
        [string]$Message,
        [ValidateSet("Info", "Success", "Warning", "Error")]
        [string]$Type = "Info"
    )
    
    $colors = @{
        "Info"    = "Cyan"
        "Success" = "Green"
        "Warning" = "Yellow"
        "Error"   = "Red"
    }
    
    $prefix = @{
        "Info"    = "[INFO]"
        "Success" = "[SUCCESS]"
        "Warning" = "[WARNING]"
        "Error"   = "[ERROR]"
    }
    
    Write-Host "$($prefix[$Type]) $Message" -ForegroundColor $colors[$Type]
}

function Get-UsersFromTemplate {
    param(
        [string]$TemplatePath
    )
    
    $users = @{}
    $extension = [System.IO.Path]::GetExtension($TemplatePath).ToLower()
    
    if ($extension -eq ".pnp") {
        # PnP files are ZIP archives containing XML
        Write-ProgressMessage "Extracting PnP template (ZIP archive)..." -Type "Info"
        
        $tempFolder = Join-Path $env:TEMP "PnPTemplateExtract_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
        New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
        
        try {
            # Extract the PnP file
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($TemplatePath, $tempFolder)
            
            # Look for manifest XML
            $manifestPath = Get-ChildItem -Path $tempFolder -Filter "*.xml" -Recurse | Select-Object -First 1
            
            if ($manifestPath) {
                Write-ProgressMessage "Analyzing manifest: $($manifestPath.Name)" -Type "Info"
                [xml]$templateXml = Get-Content $manifestPath.FullName -Raw
                $users = Extract-UsersFromXml -TemplateXml $templateXml
            }
            else {
                Write-ProgressMessage "No XML manifest found in PnP file" -Type "Warning"
            }
        }
        finally {
            # Cleanup temp folder
            if (Test-Path $tempFolder) {
                Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
    elseif ($extension -eq ".xml") {
        Write-ProgressMessage "Analyzing XML template..." -Type "Info"
        [xml]$templateXml = Get-Content $TemplatePath -Raw
        $users = Extract-UsersFromXml -TemplateXml $templateXml
    }
    else {
        throw "Unsupported template file type: $extension. Expected .pnp or .xml"
    }
    
    return $users
}

function Extract-UsersFromXml {
    param(
        [xml]$TemplateXml
    )
    
    $users = @{}
    
    # Extract from Security section
    if ($TemplateXml.Provisioning.Templates.ProvisioningTemplate.Security) {
        Write-ProgressMessage "Scanning security settings..." -Type "Info"
        
        $security = $TemplateXml.Provisioning.Templates.ProvisioningTemplate.Security
        
        # Site Groups and Members
        foreach ($group in $security.AdditionalAdministrators.User) {
            if ($group.Name) {
                $users[$group.Name] = @{
                    Email = $group.Name
                    DisplayName = $group.Name
                    Context = "Site Administrator"
                }
            }
        }
        
        foreach ($group in $security.SiteGroups.SiteGroup) {
            foreach ($member in $group.Members.User) {
                if ($member.Name) {
                    $users[$member.Name] = @{
                        Email = $member.Name
                        DisplayName = $member.Name
                        Context = "Site Group: $($group.Title)"
                    }
                }
            }
        }
    }
    
    # Extract from Lists section (Created By, Modified By, User fields)
    if ($TemplateXml.Provisioning.Templates.ProvisioningTemplate.Lists) {
        Write-ProgressMessage "Scanning list items and metadata..." -Type "Info"
        
        foreach ($list in $TemplateXml.Provisioning.Templates.ProvisioningTemplate.Lists.ListInstance) {
            foreach ($item in $list.DataRows.DataRow) {
                foreach ($value in $item.DataValue) {
                    # Check for user fields (Person or User columns typically contain email or claims)
                    if ($value.FieldName -match "(Author|Editor|Created.*By|Modified.*By|Assigned.*To|Owner|Manager)" -or
                        $value.'#text' -match '@') {
                        
                        $userText = $value.'#text'
                        
                        # Extract email from various formats
                        if ($userText -match '([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})') {
                            $email = $matches[1]
                            if (-not $users.ContainsKey($email)) {
                                $users[$email] = @{
                                    Email = $email
                                    DisplayName = $email.Split('@')[0]
                                    Context = "List: $($list.Title), Field: $($value.FieldName)"
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    
    # Extract from Files section (Created By, Modified By)
    if ($TemplateXml.Provisioning.Templates.ProvisioningTemplate.Files) {
        Write-ProgressMessage "Scanning files and documents..." -Type "Info"
        
        foreach ($file in $TemplateXml.Provisioning.Templates.ProvisioningTemplate.Files.File) {
            # Check properties
            foreach ($prop in $file.Properties.Property) {
                if ($prop.Key -match "(Author|Editor|Created.*By|Modified.*By)" -and $prop.Value -match '@') {
                    if ($prop.Value -match '([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})') {
                        $email = $matches[1]
                        if (-not $users.ContainsKey($email)) {
                            $users[$email] = @{
                                Email = $email
                                DisplayName = $email.Split('@')[0]
                                Context = "File: $($file.Src)"
                            }
                        }
                    }
                }
            }
        }
    }
    
    return $users
}

function Get-UsersFromSite {
    param(
        [string]$SiteUrl,
        [string]$ConfigFile
    )
    
    Write-ProgressMessage "Connecting to SharePoint site..." -Type "Info"
    
    # Source the connection function from Import script
    $importScriptPath = Join-Path $PSScriptRoot "Import-SharePointSiteTemplate.ps1"
    if (Test-Path $importScriptPath) {
        . $importScriptPath
    }
    
    Connect-SharePoint -SiteUrl $SiteUrl -ConfigFilePath $ConfigFile
    
    $users = @{}
    
    try {
        Write-ProgressMessage "Scanning site users..." -Type "Info"
        
        # Get all site users
        $siteUsers = Get-PnPUser
        foreach ($user in $siteUsers) {
            if ($user.Email -and $user.Email -match '@') {
                $users[$user.Email] = @{
                    Email = $user.Email
                    DisplayName = $user.Title
                    Context = "Site User"
                }
            }
        }
        
        Write-ProgressMessage "Scanning site groups..." -Type "Info"
        
        # Get users from site groups
        $groups = Get-PnPGroup
        foreach ($group in $groups) {
            $groupUsers = Get-PnPGroupMember -Identity $group.Title
            foreach ($user in $groupUsers) {
                if ($user.Email -and $user.Email -match '@') {
                    $users[$user.Email] = @{
                        Email = $user.Email
                        DisplayName = $user.Title
                        Context = "Group: $($group.Title)"
                    }
                }
            }
        }
        
        Write-ProgressMessage "Scanning lists for user fields..." -Type "Info"
        
        # Get users from list items (Author, Editor, custom user fields)
        $lists = Get-PnPList | Where-Object { -not $_.Hidden }
        foreach ($list in $lists) {
            $items = Get-PnPListItem -List $list.Title -PageSize 500 -ErrorAction SilentlyContinue
            
            foreach ($item in $items) {
                # Check Author and Editor
                if ($item.FieldValues.Author -and $item.FieldValues.Author.Email) {
                    $email = $item.FieldValues.Author.Email
                    if (-not $users.ContainsKey($email)) {
                        $users[$email] = @{
                            Email = $email
                            DisplayName = $item.FieldValues.Author.LookupValue
                            Context = "List: $($list.Title) (Author)"
                        }
                    }
                }
                
                if ($item.FieldValues.Editor -and $item.FieldValues.Editor.Email) {
                    $email = $item.FieldValues.Editor.Email
                    if (-not $users.ContainsKey($email)) {
                        $users[$email] = @{
                            Email = $email
                            DisplayName = $item.FieldValues.Editor.LookupValue
                            Context = "List: $($list.Title) (Editor)"
                        }
                    }
                }
                
                # Check for other user fields
                foreach ($key in $item.FieldValues.Keys) {
                    $value = $item.FieldValues[$key]
                    if ($value -is [Microsoft.SharePoint.Client.FieldUserValue]) {
                        if ($value.Email) {
                            if (-not $users.ContainsKey($value.Email)) {
                                $users[$value.Email] = @{
                                    Email = $value.Email
                                    DisplayName = $value.LookupValue
                                    Context = "List: $($list.Title), Field: $key"
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    finally {
        Disconnect-PnPOnline
        Write-ProgressMessage "Disconnected from SharePoint" -Type "Info"
    }
    
    return $users
}

function Export-UserMappingTemplate {
    param(
        [hashtable]$Users,
        [string]$OutputPath,
        [bool]$IncludeSystemAccounts
    )
    
    $filteredUsers = $Users.GetEnumerator() | Where-Object {
        $email = $_.Value.Email
        
        # Filter system accounts if requested
        if (-not $IncludeSystemAccounts) {
            if ($email -match '(sharepoint|system)@' -or 
                $email -match 'app@sharepoint' -or
                $_.Value.DisplayName -match 'System Account|SharePoint App') {
                return $false
            }
        }
        
        return $true
    }
    
    # Create CSV data
    $csvData = @()
    $csvData += "SourceUser,TargetUser,SourceDisplayName,TargetDisplayName,Notes"
    
    foreach ($user in ($filteredUsers | Sort-Object { $_.Value.Email })) {
        $email = $user.Value.Email
        $displayName = $user.Value.DisplayName
        $context = $user.Value.Context
        
        # Pre-populate TargetUser with same email (user can edit if different)
        $csvData += "$email,$email,$displayName,$displayName,Found in: $context"
    }
    
    # Save to file
    $csvData | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
    
    Write-ProgressMessage "User mapping template created: $OutputPath" -Type "Success"
    Write-ProgressMessage "Total users: $($filteredUsers.Count)" -Type "Info"
}

#endregion

#region Main Script

try {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host " SharePoint User Mapping Template Generator" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    
    $users = @{}
    
    if ($PSCmdlet.ParameterSetName -eq 'FromTemplate') {
        Write-ProgressMessage "Processing template file: $TemplatePath" -Type "Info"
        $users = Get-UsersFromTemplate -TemplatePath $TemplatePath
    }
    else {
        Write-ProgressMessage "Scanning live SharePoint site: $SiteUrl" -Type "Info"
        
        # Ensure PnP module
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            throw "PnP.PowerShell module not found. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
        }
        
        $users = Get-UsersFromSite -SiteUrl $SiteUrl -ConfigFile $ConfigFile
    }
    
    if ($users.Count -eq 0) {
        Write-ProgressMessage "No users found in the specified source" -Type "Warning"
        Write-Host ""
        exit 0
    }
    
    Write-ProgressMessage "Found $($users.Count) unique user(s)" -Type "Success"
    
    # Export to CSV
    Export-UserMappingTemplate -Users $users -OutputPath $OutputPath -IncludeSystemAccounts $IncludeSystemAccounts
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Green
    Write-Host " Next Steps:" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Green
    Write-Host ""
    Write-Host "1. Review the generated file: $OutputPath" -ForegroundColor White
    Write-Host "2. Update the 'TargetUser' column with target tenant emails" -ForegroundColor White
    Write-Host "3. Update 'TargetDisplayName' if names are different" -ForegroundColor White
    Write-Host "4. Leave 'TargetUser' empty for users that should not be mapped" -ForegroundColor White
    Write-Host "5. Use the mapping file with Import script:" -ForegroundColor White
    Write-Host "   -UserMappingFile '$OutputPath'" -ForegroundColor Gray
    Write-Host ""
}
catch {
    Write-ProgressMessage "Error: $($_.Exception.Message)" -Type "Error"
    Write-Host ""
    exit 1
}

#endregion
