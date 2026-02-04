<#
.SYNOPSIS
    Inspects and analyzes a SharePoint PnP template file (.pnp) to show its contents.

.DESCRIPTION
    This script extracts and analyzes a PnP template file to provide detailed information
    about what the template contains, including lists, libraries, pages, users, content types,
    and site structure. Helps make informed decisions before importing templates.

.PARAMETER TemplatePath
    Path to the PnP template file (.pnp) to inspect.

.PARAMETER OutputFormat
    Output format for the inspection report. Options: Console (default), JSON, CSV, HTML.

.PARAMETER OutputPath
    Path where the inspection report will be saved (for JSON, CSV, or HTML formats).

.PARAMETER ShowUsers
    Include detailed user analysis in the report.

.PARAMETER ShowContent
    Include content statistics (item counts, file sizes) in the report.

.PARAMETER CompareTo
    Optional path to another template file for comparison.

.PARAMETER Detailed
    Show detailed information including field definitions, content type details, etc.

.EXAMPLE
    .\Get-TemplateContent.ps1 -TemplatePath "C:\Templates\Site.pnp"

.EXAMPLE
    .\Get-TemplateContent.ps1 -TemplatePath "Site.pnp" -OutputFormat JSON -OutputPath "manifest.json"

.EXAMPLE
    .\Get-TemplateContent.ps1 -TemplatePath "Old.pnp" -CompareTo "New.pnp"

.EXAMPLE
    .\Get-TemplateContent.ps1 -TemplatePath "Site.pnp" -Detailed -ShowUsers -ShowContent

.NOTES
    Author: IT Support
    Date: February 4, 2026
    Version: 1.0
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_)) {
            throw "Template file not found: $_"
        }
        if ($_ -notmatch '\.(pnp|xml)$') {
            throw "File must be a .pnp or .xml template file"
        }
        return $true
    })]
    [string]$TemplatePath,

    [Parameter(Mandatory = $false)]
    [ValidateSet("Console", "JSON", "CSV", "HTML")]
    [string]$OutputFormat = "Console",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$ShowUsers,

    [Parameter(Mandatory = $false)]
    [switch]$ShowContent,

    [Parameter(Mandatory = $false)]
    [ValidateScript({
        if ($_ -and -not (Test-Path $_)) {
            throw "Comparison template file not found: $_"
        }
        return $true
    })]
    [string]$CompareTo,

    [Parameter(Mandatory = $false)]
    [switch]$Detailed
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

function Extract-PnPTemplate {
    param([string]$TemplatePath)
    
    $extension = [System.IO.Path]::GetExtension($TemplatePath).ToLower()
    $tempFolder = Join-Path $env:TEMP "PnPInspection_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    New-Item -ItemType Directory -Path $tempFolder -Force | Out-Null
    
    try {
        if ($extension -eq ".pnp") {
            Write-ProgressMessage "Extracting PnP template (ZIP archive)..." -Type "Info"
            Add-Type -AssemblyName System.IO.Compression.FileSystem
            [System.IO.Compression.ZipFile]::ExtractToDirectory($TemplatePath, $tempFolder)
            
            # Find the XML manifest
            $xmlFile = Get-ChildItem -Path $tempFolder -Filter "*.xml" -Recurse | Select-Object -First 1
            
            if (-not $xmlFile) {
                throw "No XML manifest found in PnP template"
            }
            
            [xml]$xml = Get-Content $xmlFile.FullName -Raw
            return @{
                Xml = $xml
                TempFolder = $tempFolder
                ManifestPath = $xmlFile.FullName
            }
        }
        elseif ($extension -eq ".xml") {
            Write-ProgressMessage "Loading XML template..." -Type "Info"
            [xml]$xml = Get-Content $TemplatePath -Raw
            return @{
                Xml = $xml
                TempFolder = $null
                ManifestPath = $TemplatePath
            }
        }
    }
    catch {
        if ($tempFolder -and (Test-Path $tempFolder)) {
            Remove-Item $tempFolder -Recurse -Force -ErrorAction SilentlyContinue
        }
        throw
    }
}

function Analyze-Template {
    param(
        [xml]$TemplateXml,
        [bool]$IncludeUsers,
        [bool]$IncludeContent,
        [bool]$DetailedAnalysis
    )
    
    $analysis = @{
        FileName = $null
        FileSize = $null
        SiteInfo = @{}
        Lists = @()
        Libraries = @()
        Pages = @()
        Users = @()
        ContentTypes = @()
        SiteColumns = @()
        Features = @{
            SiteCollection = @()
            Web = @()
        }
        Security = @{
            Administrators = @()
            Groups = @()
        }
        Navigation = @{
            TopNav = @()
            QuickLaunch = @()
        }
        Statistics = @{
            TotalLists = 0
            TotalLibraries = 0
            TotalPages = 0
            TotalUsers = 0
            TotalItems = 0
        }
    }
    
    $template = $TemplateXml.Provisioning.Templates.ProvisioningTemplate
    
    # Site Information
    if ($template) {
        $analysis.SiteInfo = @{
            ID = $template.ID
            Version = $template.Version
            BaseSiteTemplate = $template.BaseSiteTemplate
            Scope = $TemplateXml.Provisioning.Templates.ProvisioningTemplate.GetAttribute("Scope")
        }
    }
    
    # Lists and Libraries
    if ($template.Lists) {
        foreach ($list in $template.Lists.ListInstance) {
            $itemCount = 0
            if ($list.DataRows) {
                $itemCount = @($list.DataRows.DataRow).Count
            }
            
            $listInfo = @{
                Title = $list.Title
                Description = $list.Description
                TemplateType = $list.TemplateType
                Url = $list.Url
                ItemCount = $itemCount
                EnableVersioning = $list.EnableVersioning
                EnableFolderCreation = $list.EnableFolderCreation
                Hidden = $list.Hidden
                ContentTypesEnabled = $list.ContentTypesEnabled
            }
            
            # Determine if it's a library or list
            $libraryTypes = @(101, 109, 119, 851) # Document, Picture, Web Page, Asset Library
            if ($list.TemplateType -in $libraryTypes) {
                $analysis.Libraries += $listInfo
                $analysis.Statistics.TotalLibraries++
            }
            else {
                $analysis.Lists += $listInfo
                $analysis.Statistics.TotalLists++
            }
            
            $analysis.Statistics.TotalItems += $itemCount
        }
    }
    
    # Pages
    if ($template.Files) {
        foreach ($file in $template.Files.File) {
            if ($file.Src -match '\.(aspx|html)$') {
                $analysis.Pages += @{
                    Name = [System.IO.Path]::GetFileName($file.Src)
                    Path = $file.Src
                    Level = $file.Level
                    Overwrite = $file.Overwrite
                }
                $analysis.Statistics.TotalPages++
            }
        }
    }
    
    # Users
    if ($IncludeUsers) {
        $users = @{}
        
        # From Security
        if ($template.Security) {
            foreach ($admin in $template.Security.AdditionalAdministrators.User) {
                if ($admin.Name) {
                    if (-not $users.ContainsKey($admin.Name)) {
                        $users[$admin.Name] = @{
                            Email = $admin.Name
                            Roles = @("Site Administrator")
                            ItemsCreated = 0
                        }
                    }
                }
            }
            
            foreach ($group in $template.Security.SiteGroups.SiteGroup) {
                foreach ($member in $group.Members.User) {
                    if ($member.Name) {
                        if (-not $users.ContainsKey($member.Name)) {
                            $users[$member.Name] = @{
                                Email = $member.Name
                                Roles = @()
                                ItemsCreated = 0
                            }
                        }
                        $users[$member.Name].Roles += "Group: $($group.Title)"
                    }
                }
            }
        }
        
        # From List Items (count items created)
        if ($IncludeContent -and $template.Lists) {
            foreach ($list in $template.Lists.ListInstance) {
                foreach ($item in $list.DataRows.DataRow) {
                    foreach ($value in $item.DataValue) {
                        if ($value.FieldName -match "(Author|Created.*By)" -and $value.'#text' -match '@') {
                            if ($value.'#text' -match '([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})') {
                                $email = $matches[1]
                                if (-not $users.ContainsKey($email)) {
                                    $users[$email] = @{
                                        Email = $email
                                        Roles = @("Content Creator")
                                        ItemsCreated = 0
                                    }
                                }
                                $users[$email].ItemsCreated++
                            }
                        }
                    }
                }
            }
        }
        
        $analysis.Users = $users.Values | Sort-Object -Property Email
        $analysis.Statistics.TotalUsers = $users.Count
    }
    
    # Content Types
    if ($template.ContentTypes -and $DetailedAnalysis) {
        foreach ($ct in $template.ContentTypes.ContentType) {
            $analysis.ContentTypes += @{
                ID = $ct.ID
                Name = $ct.Name
                Group = $ct.Group
                Description = $ct.Description
            }
        }
    }
    
    # Site Columns
    if ($template.SiteFields -and $DetailedAnalysis) {
        foreach ($field in $template.SiteFields.Field) {
            $analysis.SiteColumns += @{
                ID = $field.ID
                Name = $field.Name
                DisplayName = $field.DisplayName
                Type = $field.Type
                Group = $field.Group
            }
        }
    }
    
    # Features
    if ($template.Features -and $DetailedAnalysis) {
        foreach ($feature in $template.Features.SiteFeatures.Feature) {
            $analysis.Features.SiteCollection += @{
                ID = $feature.ID
                Name = $feature.Description
            }
        }
        
        foreach ($feature in $template.Features.WebFeatures.Feature) {
            $analysis.Features.Web += @{
                ID = $feature.ID
                Name = $feature.Description
            }
        }
    }
    
    return $analysis
}

function Format-ConsoleOutput {
    param([hashtable]$Analysis, [string]$TemplateFile)
    
    $fileInfo = Get-Item $TemplateFile
    $fileSizeMB = [math]::Round($fileInfo.Length / 1MB, 2)
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  SharePoint Template Analysis Report" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "File: " -NoNewline -ForegroundColor White
    Write-Host $fileInfo.Name -ForegroundColor Yellow
    Write-Host "Size: " -NoNewline -ForegroundColor White
    Write-Host "$fileSizeMB MB" -ForegroundColor Yellow
    Write-Host "Path: " -NoNewline -ForegroundColor White
    Write-Host $fileInfo.FullName -ForegroundColor Gray
    Write-Host ""
    
    # Site Information
    if ($Analysis.SiteInfo.BaseSiteTemplate) {
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host " SITE INFORMATION" -ForegroundColor Cyan
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host "  Base Template: " -NoNewline -ForegroundColor White
        Write-Host $Analysis.SiteInfo.BaseSiteTemplate -ForegroundColor Yellow
        if ($Analysis.SiteInfo.Version) {
            Write-Host "  Version: " -NoNewline -ForegroundColor White
            Write-Host $Analysis.SiteInfo.Version -ForegroundColor Yellow
        }
        Write-Host ""
    }
    
    # Statistics Summary
    Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host " SUMMARY STATISTICS" -ForegroundColor Cyan
    Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
    Write-Host "  Lists:         " -NoNewline -ForegroundColor White
    Write-Host $Analysis.Statistics.TotalLists -ForegroundColor Green
    Write-Host "  Libraries:     " -NoNewline -ForegroundColor White
    Write-Host $Analysis.Statistics.TotalLibraries -ForegroundColor Green
    Write-Host "  Pages:         " -NoNewline -ForegroundColor White
    Write-Host $Analysis.Statistics.TotalPages -ForegroundColor Green
    Write-Host "  Total Items:   " -NoNewline -ForegroundColor White
    Write-Host $Analysis.Statistics.TotalItems -ForegroundColor Green
    if ($Analysis.Statistics.TotalUsers -gt 0) {
        Write-Host "  Users:         " -NoNewline -ForegroundColor White
        Write-Host $Analysis.Statistics.TotalUsers -ForegroundColor Green
    }
    Write-Host ""
    
    # Lists
    if ($Analysis.Lists.Count -gt 0) {
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host " LISTS ($($Analysis.Lists.Count))" -ForegroundColor Cyan
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        
        $Analysis.Lists | Sort-Object Title | ForEach-Object {
            Write-Host "  • " -NoNewline -ForegroundColor Yellow
            Write-Host $_.Title -NoNewline -ForegroundColor White
            Write-Host " ($($_.ItemCount) items)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    # Libraries
    if ($Analysis.Libraries.Count -gt 0) {
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host " LIBRARIES ($($Analysis.Libraries.Count))" -ForegroundColor Cyan
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        
        $Analysis.Libraries | Sort-Object Title | ForEach-Object {
            Write-Host "  • " -NoNewline -ForegroundColor Yellow
            Write-Host $_.Title -NoNewline -ForegroundColor White
            Write-Host " ($($_.ItemCount) items)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    # Pages
    if ($Analysis.Pages.Count -gt 0) {
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host " SITE PAGES ($($Analysis.Pages.Count))" -ForegroundColor Cyan
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        
        $Analysis.Pages | Sort-Object Name | ForEach-Object {
            Write-Host "  • " -NoNewline -ForegroundColor Yellow
            Write-Host $_.Name -ForegroundColor White
        }
        Write-Host ""
    }
    
    # Users
    if ($Analysis.Users.Count -gt 0) {
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host " USERS REFERENCED ($($Analysis.Users.Count))" -ForegroundColor Cyan
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        
        $topUsers = $Analysis.Users | Sort-Object -Property ItemsCreated -Descending | Select-Object -First 10
        
        foreach ($user in $topUsers) {
            Write-Host "  • " -NoNewline -ForegroundColor Yellow
            Write-Host $user.Email -NoNewline -ForegroundColor White
            
            if ($user.ItemsCreated -gt 0) {
                Write-Host " ($($user.ItemsCreated) items)" -NoNewline -ForegroundColor Gray
            }
            
            if ($user.Roles.Count -gt 0) {
                Write-Host " - $($user.Roles -join ', ')" -ForegroundColor Cyan
            }
            else {
                Write-Host ""
            }
        }
        
        if ($Analysis.Users.Count -gt 10) {
            Write-Host "  ... and $($Analysis.Users.Count - 10) more users" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    # Content Types
    if ($Analysis.ContentTypes.Count -gt 0) {
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        Write-Host " CONTENT TYPES ($($Analysis.ContentTypes.Count))" -ForegroundColor Cyan
        Write-Host "─────────────────────────────────────────────────────" -ForegroundColor Gray
        
        $Analysis.ContentTypes | Sort-Object Name | ForEach-Object {
            Write-Host "  • " -NoNewline -ForegroundColor Yellow
            Write-Host $_.Name -NoNewline -ForegroundColor White
            if ($_.Group) {
                Write-Host " ($($_.Group))" -ForegroundColor Gray
            }
            else {
                Write-Host ""
            }
        }
        Write-Host ""
    }
    
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
}

function Export-AnalysisToJSON {
    param(
        [hashtable]$Analysis,
        [string]$OutputPath
    )
    
    $json = $Analysis | ConvertTo-Json -Depth 10
    $json | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
    
    Write-ProgressMessage "Analysis exported to JSON: $OutputPath" -Type "Success"
}

function Compare-TemplateFiles {
    param(
        [hashtable]$SourceAnalysis,
        [hashtable]$TargetAnalysis,
        [string]$SourceFile,
        [string]$TargetFile
    )
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
    Write-Host "  Template Comparison" -ForegroundColor Magenta
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
    Write-Host ""
    Write-Host "Source: " -NoNewline -ForegroundColor White
    Write-Host (Split-Path $SourceFile -Leaf) -ForegroundColor Yellow
    Write-Host "Target: " -NoNewline -ForegroundColor White
    Write-Host (Split-Path $TargetFile -Leaf) -ForegroundColor Yellow
    Write-Host ""
    
    # Compare Lists
    $sourceLists = $SourceAnalysis.Lists.Title
    $targetLists = $TargetAnalysis.Lists.Title
    
    $listsAdded = $targetLists | Where-Object { $_ -notin $sourceLists }
    $listsRemoved = $sourceLists | Where-Object { $_ -notin $targetLists }
    $listsCommon = $sourceLists | Where-Object { $_ -in $targetLists }
    
    Write-Host "LISTS" -ForegroundColor Cyan
    if ($listsAdded) {
        Write-Host "  Added: " -NoNewline -ForegroundColor Green
        Write-Host ($listsAdded -join ', ')
    }
    if ($listsRemoved) {
        Write-Host "  Removed: " -NoNewline -ForegroundColor Red
        Write-Host ($listsRemoved -join ', ')
    }
    if ($listsCommon) {
        Write-Host "  Unchanged: " -NoNewline -ForegroundColor Gray
        Write-Host "$($listsCommon.Count) lists"
    }
    Write-Host ""
    
    # Compare Users
    if ($SourceAnalysis.Users.Count -gt 0 -or $TargetAnalysis.Users.Count -gt 0) {
        $sourceUsers = $SourceAnalysis.Users.Email
        $targetUsers = $TargetAnalysis.Users.Email
        
        $usersAdded = $targetUsers | Where-Object { $_ -notin $sourceUsers }
        $usersRemoved = $sourceUsers | Where-Object { $_ -notin $targetUsers }
        
        Write-Host "USERS" -ForegroundColor Cyan
        if ($usersAdded) {
            Write-Host "  Added: " -NoNewline -ForegroundColor Green
            Write-Host ($usersAdded -join ', ')
        }
        if ($usersRemoved) {
            Write-Host "  Removed: " -NoNewline -ForegroundColor Red
            Write-Host ($usersRemoved -join ', ')
        }
        Write-Host ""
    }
    
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Magenta
    Write-Host ""
}

#endregion

#region Main Script

try {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host " SharePoint Template Inspector" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    # Extract and parse template
    $templateData = Extract-PnPTemplate -TemplatePath $TemplatePath
    
    Write-ProgressMessage "Analyzing template structure..." -Type "Info"
    $analysis = Analyze-Template -TemplateXml $templateData.Xml `
        -IncludeUsers:$ShowUsers `
        -IncludeContent:$ShowContent `
        -DetailedAnalysis:$Detailed
    
    # If comparing, analyze second template
    if ($CompareTo) {
        Write-ProgressMessage "Analyzing comparison template..." -Type "Info"
        $compareData = Extract-PnPTemplate -TemplatePath $CompareTo
        $compareAnalysis = Analyze-Template -TemplateXml $compareData.Xml `
            -IncludeUsers:$ShowUsers `
            -IncludeContent:$ShowContent `
            -DetailedAnalysis:$Detailed
    }
    
    # Output based on format
    switch ($OutputFormat) {
        "Console" {
            Format-ConsoleOutput -Analysis $analysis -TemplateFile $TemplatePath
            
            if ($CompareTo) {
                Compare-TemplateFiles -SourceAnalysis $analysis `
                    -TargetAnalysis $compareAnalysis `
                    -SourceFile $TemplatePath `
                    -TargetFile $CompareTo
            }
        }
        "JSON" {
            if (-not $OutputPath) {
                $OutputPath = "template-analysis-$(Get-Date -Format 'yyyyMMdd_HHmmss').json"
            }
            Export-AnalysisToJSON -Analysis $analysis -OutputPath $OutputPath
        }
        "CSV" {
            Write-ProgressMessage "CSV export not yet implemented. Use JSON format instead." -Type "Warning"
        }
        "HTML" {
            Write-ProgressMessage "HTML export not yet implemented. Use JSON format instead." -Type "Warning"
        }
    }
    
    Write-ProgressMessage "Template inspection complete" -Type "Success"
}
catch {
    Write-ProgressMessage "Error: $($_.Exception.Message)" -Type "Error"
    Write-Host ""
    Write-Host $_.ScriptStackTrace -ForegroundColor Red
    Write-Host ""
    exit 1
}
finally {
    # Cleanup temp folders
    if ($templateData.TempFolder -and (Test-Path $templateData.TempFolder)) {
        Remove-Item $templateData.TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    if ($compareData.TempFolder -and (Test-Path $compareData.TempFolder)) {
        Remove-Item $compareData.TempFolder -Recurse -Force -ErrorAction SilentlyContinue
    }
}

#endregion
