<#
.SYNOPSIS
    Compares two PnP site templates and shows the differences.

.DESCRIPTION
    This script extracts and analyzes two .pnp template files, then compares their contents
    to identify differences in lists, libraries, pages, users, content types, and other components.
    Useful for validating template changes or comparing source and target configurations.

.PARAMETER Template1Path
    Path to the first .pnp template file.

.PARAMETER Template2Path
    Path to the second .pnp template file.

.PARAMETER OutputFormat
    Output format for the comparison. Options: Console (default), JSON, HTML, CSV

.PARAMETER OutputPath
    Path to save the comparison report. If not specified, outputs to console only.

.PARAMETER CompareComponents
    Specific components to compare. Default is 'All'.
    Options: All, Lists, Pages, Users, ContentTypes, Fields, Features, Security, Navigation

.EXAMPLE
    .\Compare-Templates.ps1 -Template1Path "C:\Templates\before.pnp" -Template2Path "C:\Templates\after.pnp"
    
    Compare two templates and display differences in the console.

.EXAMPLE
    .\Compare-Templates.ps1 -Template1Path "C:\Templates\source.pnp" -Template2Path "C:\Templates\target.pnp" `
        -OutputFormat JSON -OutputPath "C:\Reports\comparison.json"
    
    Compare templates and save results as JSON.

.EXAMPLE
    .\Compare-Templates.ps1 -Template1Path "C:\Templates\v1.pnp" -Template2Path "C:\Templates\v2.pnp" `
        -CompareComponents Lists,Pages
    
    Compare only lists and pages between two templates.

.NOTES
    Author: IT Support
    Date: February 4, 2026
    Requires: System.IO.Compression.FileSystem for ZIP handling
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_)) { throw "Template file not found: $_" }
        if ($_ -notmatch '\.pnp$') { throw "File must be a .pnp template file" }
        return $true
    })]
    [string]$Template1Path,

    [Parameter(Mandatory = $true)]
    [ValidateScript({
        if (-not (Test-Path $_)) { throw "Template file not found: $_" }
        if ($_ -notmatch '\.pnp$') { throw "File must be a .pnp template file" }
        return $true
    })]
    [string]$Template2Path,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Console', 'JSON', 'HTML', 'CSV')]
    [string]$OutputFormat = 'Console',

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'Lists', 'Pages', 'Users', 'ContentTypes', 'Fields', 'Features', 'Security', 'Navigation')]
    [string[]]$CompareComponents = @('All')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Functions

function Extract-PnPTemplate {
    param([string]$TemplatePath)
    
    try {
        # Load ZIP assembly
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        
        # Create temp extraction folder
        $extractPath = Join-Path ([System.IO.Path]::GetTempPath()) "PnPTemplate_$([Guid]::NewGuid())"
        
        # Ensure temp folder doesn't exist
        if (Test-Path $extractPath) {
            Remove-Item $extractPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        
        New-Item -ItemType Directory -Path $extractPath -Force | Out-Null
        
        # Extract ZIP
        [System.IO.Compression.ZipFile]::ExtractToDirectory($TemplatePath, $extractPath)
        
        # Find the XML manifest - look for Provisioning root element (actual template)
        # Newer .pnp files have the template in Files subfolder, older ones have it at root
        $xmlFile = Get-ChildItem -Path $extractPath -Filter "*.xml" -Recurse | 
            Where-Object { $_.Name -ne '[Content_Types].xml' } |
            ForEach-Object {
                try {
                    $testXml = [xml](Get-Content $_.FullName -Raw -ErrorAction Stop)
                    if ($testXml.DocumentElement.LocalName -eq 'Provisioning') {
                        return $_
                    }
                } catch {
                    # Skip files that aren't valid XML or we can't read
                }
            } | Select-Object -First 1
        
        if (-not $xmlFile) {
            throw "No Provisioning template XML found in PnP package"
        }
        
        # Load XML
        [xml]$xmlContent = Get-Content -Path $xmlFile.FullName -Raw
        
        return @{
            ExtractPath = $extractPath
            XmlFile = $xmlFile.FullName
            Xml = $xmlContent
        }
    }
    catch {
        throw "Failed to extract template: $($_.Exception.Message)"
    }
}

function Get-TemplateListsAndLibraries {
    param([xml]$Xml)
    
    $lists = @()
    $listInstances = $Xml.SelectNodes("//pnp:ListInstance", $nsmgr)
    
    foreach ($list in $listInstances) {
        $lists += [PSCustomObject]@{
            Title = $list.Title
            Url = $list.Url
            TemplateType = $list.TemplateType
            Description = $list.Description
            OnQuickLaunch = $list.OnQuickLaunch
            ContentTypesEnabled = $list.ContentTypesEnabled
        }
    }
    
    return $lists
}

function Get-TemplatePages {
    param([xml]$Xml)
    
    $pages = @()
    $pageNodes = $Xml.SelectNodes("//pnp:ClientSidePages/pnp:ClientSidePage", $nsmgr)
    
    foreach ($page in $pageNodes) {
        $pages += [PSCustomObject]@{
            Name = $page.Name
            Title = $page.Title
            Layout = $page.PageLayoutType
            PromoteAsNewsArticle = $page.PromoteAsNewsArticle
        }
    }
    
    return $pages
}

function Get-TemplateUsers {
    param([xml]$Xml)
    
    $users = @()
    $userPattern = 'i:0#\.f\|membership\|([^"]+@[^"]+)'
    
    $xmlString = $Xml.OuterXml
    $matches = [regex]::Matches($xmlString, $userPattern)
    
    foreach ($match in $matches) {
        $email = $match.Groups[1].Value
        if ($users -notcontains $email) {
            $users += $email
        }
    }
    
    return $users | Sort-Object
}

function Get-TemplateContentTypes {
    param([xml]$Xml)
    
    $contentTypes = @()
    $ctNodes = $Xml.SelectNodes("//pnp:ContentType", $nsmgr)
    
    foreach ($ct in $ctNodes) {
        $contentTypes += [PSCustomObject]@{
            Id = $ct.ID
            Name = $ct.Name
            Description = $ct.Description
            Group = $ct.Group
        }
    }
    
    return $contentTypes
}

function Compare-Collections {
    param(
        [string]$ComponentName,
        [object[]]$Collection1,
        [object[]]$Collection2,
        [string]$KeyProperty = "Title"
    )
    
    $items1 = @($Collection1 | ForEach-Object { $_.$KeyProperty })
    $items2 = @($Collection2 | ForEach-Object { $_.$KeyProperty })
    
    $onlyIn1 = $items1 | Where-Object { $items2 -notcontains $_ }
    $onlyIn2 = $items2 | Where-Object { $items1 -notcontains $_ }
    $inBoth = $items1 | Where-Object { $items2 -contains $_ }
    
    return [PSCustomObject]@{
        Component = $ComponentName
        OnlyInTemplate1 = $onlyIn1
        OnlyInTemplate2 = $onlyIn2
        InBoth = $inBoth
        TotalTemplate1 = $items1.Count
        TotalTemplate2 = $items2.Count
        Differences = ($onlyIn1.Count + $onlyIn2.Count)
    }
}

function Format-ConsoleComparison {
    param([object[]]$ComparisonResults)
    
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  TEMPLATE COMPARISON RESULTS" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    foreach ($result in $ComparisonResults) {
        Write-Host "Component: $($result.Component)" -ForegroundColor Yellow
        Write-Host "  Template 1: $($result.TotalTemplate1) items" -ForegroundColor White
        Write-Host "  Template 2: $($result.TotalTemplate2) items" -ForegroundColor White
        Write-Host "  Total Differences: $($result.Differences)" -ForegroundColor $(if ($result.Differences -gt 0) { "Red" } else { "Green" })
        Write-Host ""
        
        if ($result.OnlyInTemplate1.Count -gt 0) {
            Write-Host "  Only in Template 1 ($($result.OnlyInTemplate1.Count)):" -ForegroundColor Magenta
            foreach ($item in $result.OnlyInTemplate1) {
                Write-Host "    - $item" -ForegroundColor Gray
            }
            Write-Host ""
        }
        
        if ($result.OnlyInTemplate2.Count -gt 0) {
            Write-Host "  Only in Template 2 ($($result.OnlyInTemplate2.Count)):" -ForegroundColor Cyan
            foreach ($item in $result.OnlyInTemplate2) {
                Write-Host "    + $item" -ForegroundColor Gray
            }
            Write-Host ""
        }
        
        if ($result.InBoth.Count -gt 0 -and $result.InBoth.Count -le 10) {
            Write-Host "  In Both Templates ($($result.InBoth.Count)):" -ForegroundColor Green
            foreach ($item in $result.InBoth) {
                Write-Host "    • $item" -ForegroundColor DarkGray
            }
            Write-Host ""
        }
        elseif ($result.InBoth.Count -gt 10) {
            Write-Host "  In Both Templates: $($result.InBoth.Count) items" -ForegroundColor Green
            Write-Host ""
        }
    }
    
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
}

function Export-ComparisonToJSON {
    param(
        [object[]]$ComparisonResults,
        [string]$OutputPath
    )
    
    $jsonOutput = @{
        GeneratedAt = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        Template1 = $Template1Path
        Template2 = $Template2Path
        Comparisons = $ComparisonResults
    } | ConvertTo-Json -Depth 10
    
    $jsonOutput | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "Comparison exported to: $OutputPath" -ForegroundColor Green
}

function Export-ComparisonToCSV {
    param(
        [object[]]$ComparisonResults,
        [string]$OutputPath
    )
    
    $csvData = @()
    
    foreach ($result in $ComparisonResults) {
        $csvData += [PSCustomObject]@{
            Component = $result.Component
            Template1Count = $result.TotalTemplate1
            Template2Count = $result.TotalTemplate2
            Differences = $result.Differences
            OnlyInTemplate1 = ($result.OnlyInTemplate1 -join '; ')
            OnlyInTemplate2 = ($result.OnlyInTemplate2 -join '; ')
            InBoth = ($result.InBoth -join '; ')
        }
    }
    
    $csvData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host "Comparison exported to: $OutputPath" -ForegroundColor Green
}

function Export-ComparisonToHTML {
    param(
        [object[]]$ComparisonResults,
        [string]$OutputPath
    )
    
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>Template Comparison Report</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background-color: #f5f5f5; }
        h1 { color: #0078d4; border-bottom: 3px solid #0078d4; padding-bottom: 10px; }
        .metadata { background: white; padding: 15px; margin-bottom: 20px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .component { background: white; padding: 20px; margin-bottom: 20px; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .component-header { font-size: 1.3em; color: #0078d4; margin-bottom: 15px; font-weight: bold; }
        .stats { display: flex; gap: 20px; margin-bottom: 15px; }
        .stat { flex: 1; padding: 10px; background: #f0f0f0; border-radius: 3px; text-align: center; }
        .stat-value { font-size: 2em; font-weight: bold; color: #0078d4; }
        .stat-label { font-size: 0.9em; color: #666; }
        .diff-section { margin-top: 15px; }
        .diff-header { font-weight: bold; margin-bottom: 5px; padding: 5px; background: #f0f0f0; border-radius: 3px; }
        .diff-removed { color: #d13438; }
        .diff-added { color: #107c10; }
        .diff-same { color: #666; }
        ul { list-style-type: none; padding-left: 20px; }
        li { padding: 3px 0; }
        .no-diff { color: #107c10; font-weight: bold; padding: 10px; background: #e7f5e7; border-radius: 3px; }
    </style>
</head>
<body>
    <h1>SharePoint Template Comparison Report</h1>
    <div class="metadata">
        <p><strong>Generated:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")</p>
        <p><strong>Template 1:</strong> $([System.IO.Path]::GetFileName($Template1Path))</p>
        <p><strong>Template 2:</strong> $([System.IO.Path]::GetFileName($Template2Path))</p>
    </div>
"@
    
    foreach ($result in $ComparisonResults) {
        $diffClass = if ($result.Differences -eq 0) { "no-diff" } else { "" }
        
        $html += @"
    <div class="component">
        <div class="component-header">$($result.Component)</div>
        <div class="stats">
            <div class="stat">
                <div class="stat-value">$($result.TotalTemplate1)</div>
                <div class="stat-label">Template 1</div>
            </div>
            <div class="stat">
                <div class="stat-value">$($result.TotalTemplate2)</div>
                <div class="stat-label">Template 2</div>
            </div>
            <div class="stat">
                <div class="stat-value" style="color: $(if ($result.Differences -gt 0) { '#d13438' } else { '#107c10' })">$($result.Differences)</div>
                <div class="stat-label">Differences</div>
            </div>
        </div>
"@
        
        if ($result.Differences -eq 0) {
            $html += "        <div class='no-diff'>✓ No differences - templates match for this component</div>`n"
        }
        else {
            if ($result.OnlyInTemplate1.Count -gt 0) {
                $html += "        <div class='diff-section'>`n"
                $html += "            <div class='diff-header diff-removed'>Only in Template 1 ($($result.OnlyInTemplate1.Count)):</div>`n"
                $html += "            <ul>`n"
                foreach ($item in $result.OnlyInTemplate1) {
                    $html += "                <li class='diff-removed'>- $([System.Web.HttpUtility]::HtmlEncode($item))</li>`n"
                }
                $html += "            </ul>`n"
                $html += "        </div>`n"
            }
            
            if ($result.OnlyInTemplate2.Count -gt 0) {
                $html += "        <div class='diff-section'>`n"
                $html += "            <div class='diff-header diff-added'>Only in Template 2 ($($result.OnlyInTemplate2.Count)):</div>`n"
                $html += "            <ul>`n"
                foreach ($item in $result.OnlyInTemplate2) {
                    $html += "                <li class='diff-added'>+ $([System.Web.HttpUtility]::HtmlEncode($item))</li>`n"
                }
                $html += "            </ul>`n"
                $html += "        </div>`n"
            }
        }
        
        $html += "    </div>`n"
    }
    
    $html += @"
</body>
</html>
"@
    
    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "Comparison exported to: $OutputPath" -ForegroundColor Green
}

#endregion

#region Main Script

try {
    Write-Host ""
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host "  PnP Template Comparison Tool" -ForegroundColor Cyan
    Write-Host "═══════════════════════════════════════════════════════════════════════" -ForegroundColor Cyan
    Write-Host ""
    
    # Extract and analyze template 1
    Write-Host "Extracting Template 1..." -ForegroundColor Yellow
    $template1 = Extract-PnPTemplate -TemplatePath $Template1Path
    
    # Create namespace manager for XPath queries
    $nsmgr = New-Object System.Xml.XmlNamespaceManager($template1.Xml.NameTable)
    $nsmgr.AddNamespace("pnp", "http://schemas.dev.office.com/PnP/2015/12/ProvisioningSchema")
    
    # Extract and analyze template 2
    Write-Host "Extracting Template 2..." -ForegroundColor Yellow
    $template2 = Extract-PnPTemplate -TemplatePath $Template2Path
    
    Write-Host "Analyzing templates..." -ForegroundColor Yellow
    Write-Host ""
    
    $comparisonResults = @()
    
    # Compare Lists and Libraries
    if ($CompareComponents -contains 'All' -or $CompareComponents -contains 'Lists') {
        $lists1 = Get-TemplateListsAndLibraries -Xml $template1.Xml
        $lists2 = Get-TemplateListsAndLibraries -Xml $template2.Xml
        $comparisonResults += Compare-Collections -ComponentName "Lists & Libraries" -Collection1 $lists1 -Collection2 $lists2 -KeyProperty "Title"
    }
    
    # Compare Pages
    if ($CompareComponents -contains 'All' -or $CompareComponents -contains 'Pages') {
        $pages1 = Get-TemplatePages -Xml $template1.Xml
        $pages2 = Get-TemplatePages -Xml $template2.Xml
        $comparisonResults += Compare-Collections -ComponentName "Pages" -Collection1 $pages1 -Collection2 $pages2 -KeyProperty "Name"
    }
    
    # Compare Users
    if ($CompareComponents -contains 'All' -or $CompareComponents -contains 'Users') {
        $users1 = Get-TemplateUsers -Xml $template1.Xml
        $users2 = Get-TemplateUsers -Xml $template2.Xml
        $usersObj1 = $users1 | ForEach-Object { [PSCustomObject]@{ Email = $_ } }
        $usersObj2 = $users2 | ForEach-Object { [PSCustomObject]@{ Email = $_ } }
        $comparisonResults += Compare-Collections -ComponentName "Users" -Collection1 $usersObj1 -Collection2 $usersObj2 -KeyProperty "Email"
    }
    
    # Compare Content Types
    if ($CompareComponents -contains 'All' -or $CompareComponents -contains 'ContentTypes') {
        $ct1 = Get-TemplateContentTypes -Xml $template1.Xml
        $ct2 = Get-TemplateContentTypes -Xml $template2.Xml
        $comparisonResults += Compare-Collections -ComponentName "Content Types" -Collection1 $ct1 -Collection2 $ct2 -KeyProperty "Name"
    }
    
    # Output results
    switch ($OutputFormat) {
        'Console' {
            Format-ConsoleComparison -ComparisonResults $comparisonResults
        }
        'JSON' {
            if (-not $OutputPath) {
                $OutputPath = Join-Path (Get-Location) "template-comparison.json"
            }
            Export-ComparisonToJSON -ComparisonResults $comparisonResults -OutputPath $OutputPath
            Format-ConsoleComparison -ComparisonResults $comparisonResults
        }
        'CSV' {
            if (-not $OutputPath) {
                $OutputPath = Join-Path (Get-Location) "template-comparison.csv"
            }
            Export-ComparisonToCSV -ComparisonResults $comparisonResults -OutputPath $OutputPath
            Format-ConsoleComparison -ComparisonResults $comparisonResults
        }
        'HTML' {
            if (-not $OutputPath) {
                $OutputPath = Join-Path (Get-Location) "template-comparison.html"
            }
            Export-ComparisonToHTML -ComparisonResults $comparisonResults -OutputPath $OutputPath
            Format-ConsoleComparison -ComparisonResults $comparisonResults
        }
    }
    
    # Cleanup temp folders
    Remove-Item -Path $template1.ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path $template2.ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    
    Write-Host "Comparison complete!" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    exit 1
}
finally {
    # Ensure cleanup
    if ($template1 -and $template1.ExtractPath -and (Test-Path $template1.ExtractPath)) {
        Remove-Item -Path $template1.ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    if ($template2 -and $template2.ExtractPath -and (Test-Path $template2.ExtractPath)) {
        Remove-Item -Path $template2.ExtractPath -Recurse -Force -ErrorAction SilentlyContinue
    }
}

#endregion
