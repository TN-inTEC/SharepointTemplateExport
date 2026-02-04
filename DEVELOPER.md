# Developer Guide

## Overview

This document provides guidelines, standards, and best practices for contributing to the SharePoint Template Export/Import toolset. Following these guidelines ensures code quality, maintainability, and consistency across the project.

## Table of Contents

- [Development Philosophy](#development-philosophy)
- [Before You Start](#before-you-start)
- [Code Standards](#code-standards)
- [Feature Development Workflow](#feature-development-workflow)
- [Testing Requirements](#testing-requirements)
- [Documentation Standards](#documentation-standards)
- [Commit Guidelines](#commit-guidelines)
- [Pull Request Process](#pull-request-process)
- [Common Patterns](#common-patterns)
- [Security Guidelines](#security-guidelines)

---

## Development Philosophy

### Core Principles

1. **MSP-Focused**: Solutions should be practical for managed service providers and IT operations teams
2. **Production-Ready**: Code must be robust, with proper error handling and logging
3. **Self-Documenting**: Code should be clear; use functions with descriptive names
4. **User-Friendly**: Provide helpful error messages and guidance
5. **Secure by Default**: Never commit secrets; use certificate-based auth
6. **Backwards Compatible**: Avoid breaking changes to existing scripts

### Target Audience

- IT administrators and MSP technicians
- PowerShell skill level: Intermediate
- SharePoint knowledge: Basic to intermediate
- Azure AD familiarity: Basic

---

## Before You Start

### Prerequisites

1. **Review existing code** to understand patterns and conventions
2. **Read all documentation**:
   - [README.md](README.md) - Main documentation
   - [CONFIG-README.md](CONFIG-README.md) - Configuration guide
   - [MANUAL-APP-REGISTRATION.md](MANUAL-APP-REGISTRATION.md) - Setup guide
3. **Test existing functionality** in your environment
4. **Check open issues** for related work or discussions

### Development Environment

**Required:**
- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1+ or PowerShell 7+
- PnP.PowerShell module (2.12+)
- SharePoint Online tenant (for testing)
- Azure AD app registration with certificate

**Recommended:**
- VS Code with PowerShell extension
- Git for version control
- Test tenant or isolated SharePoint site collection

---

## Code Standards

### PowerShell Style Guide

#### Script Structure

All PowerShell scripts should follow this structure:

```powershell
<#
.SYNOPSIS
    Brief one-line description

.DESCRIPTION
    Detailed description of what the script does
    Include use cases and important notes

.PARAMETER ParameterName
    Description of the parameter

.EXAMPLE
    .\Script.ps1 -ParameterName "Value"
    Description of what this example does

.NOTES
    Author: [Your Name/Team]
    Date: YYYY-MM-DD
    Requires: PnP.PowerShell module
#>

[CmdletBinding(SupportsShouldProcess = $true)]  # If script makes changes
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ParameterName,
    
    [Parameter(Mandatory = $false)]
    [switch]$OptionalSwitch
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Functions

function Verb-Noun {
    <#
    .SYNOPSIS
        Brief function description
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Parameter
    )
    
    # Function implementation
}

#endregion

#region Main Script

try {
    # Main script logic
}
catch {
    Write-ProgressMessage "Error: $($_.Exception.Message)" -Type "Error"
    throw
}
finally {
    # Cleanup
}

#endregion
```

#### Naming Conventions

- **Functions**: Use approved PowerShell verbs (Get-, Set-, New-, Remove-, etc.)
- **Variables**: Use `$camelCase` for local variables, `$PascalCase` for parameters
- **Constants**: Use `$UPPER_CASE` for constants
- **Booleans**: Prefix with `is`, `has`, `should` (e.g., `$isValid`, `$hasContent`)

**Approved verbs:**
```powershell
# Get approved verbs
Get-Verb | Where-Object { $_.Group -in @('Common', 'Data', 'Lifecycle') }
```

#### Error Handling

**Always use try-catch blocks** for operations that can fail:

```powershell
try {
    Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Thumbprint $Thumbprint
}
catch {
    Write-ProgressMessage "Failed to connect: $($_.Exception.Message)" -Type "Error"
    Write-Host "Possible causes:" -ForegroundColor Yellow
    Write-Host "  - Invalid credentials" -ForegroundColor Gray
    Write-Host "  - Network connectivity issues" -ForegroundColor Gray
    throw
}
```

**Provide actionable error messages:**

```powershell
# ❌ BAD
throw "Error occurred"

# ✅ GOOD
throw "Configuration file not found: $ConfigFile. Run: Copy-Item app-config.sample.json app-config.json"
```

#### Parameter Validation

Use PowerShell's built-in validation attributes:

```powershell
[Parameter(Mandatory = $true)]
[ValidatePattern('^https://[^/]+\.sharepoint\.com/.*$')]
[string]$SiteUrl,

[Parameter(Mandatory = $false)]
[ValidateRange(1, 10000)]
[int]$RowLimit = 5000,

[Parameter(Mandatory = $false)]
[ValidateSet("TeamSite", "CommunicationSite")]
[string]$SiteType = "TeamSite",

[Parameter(Mandatory = $false)]
[ValidateScript({
    if (-not (Test-Path $_)) {
        throw "File not found: $_"
    }
    return $true
})]
[string]$FilePath
```

#### Output and Logging

Use the standard `Write-ProgressMessage` function:

```powershell
function Write-ProgressMessage {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Success", "Warning", "Error")]
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

# Usage
Write-ProgressMessage "Connecting to SharePoint..." -Type "Info"
Write-ProgressMessage "Connection successful" -Type "Success"
Write-ProgressMessage "Template size exceeds 100MB" -Type "Warning"
Write-ProgressMessage "Authentication failed" -Type "Error"
```

#### Comments and Documentation

```powershell
# ✅ GOOD: Explain WHY, not WHAT
# PnP templates use ACS auth with client secrets (not modern auth)
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -ClientSecret $ClientSecret

# ✅ GOOD: Document complex logic
# Extract email from claims format: i:0#.f|membership|user@domain.com
if ($claimsIdentity -match '\|([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})') {
    $email = $matches[1]
}

# ❌ BAD: Stating the obvious
# Get the site URL
$url = $SiteUrl
```

---

## Feature Development Workflow

### 1. Feature Request Process

**Before starting development:**

1. **Check existing issues** - Search for related feature requests or discussions
2. **Create an issue** (if none exists) with:
   - Clear description of the feature
   - Use case / business justification
   - Expected behavior
   - Example usage
3. **Get approval** - Discuss with maintainers before significant work

**Issue Template:**
```markdown
## Feature Request: [Feature Name]

### Description
Brief description of the feature

### Use Case
Who needs this and why? What problem does it solve?

### Proposed Solution
How should this work?

### Example Usage
```powershell
# Example of how the feature would be used
.\Script.ps1 -NewParameter "Value"
```

### Alternatives Considered
What other approaches were considered?

### Impact
- Breaking changes? (Yes/No)
- New dependencies? (Yes/No)
- Affects existing scripts? (Yes/No)
```

### 2. Development Process

**Step-by-step workflow:**

1. **Create a feature branch**
   ```powershell
   git checkout -b feature/user-mapping-enhancements
   ```

2. **Follow the design pattern**
   - Use existing functions as templates
   - Maintain consistent error handling
   - Add comprehensive parameter validation

3. **Implement incrementally**
   - Start with core functionality
   - Add error handling
   - Add logging and progress messages
   - Add parameter validation

4. **Self-review checklist:**
   - [ ] Code follows style guide
   - [ ] Error handling implemented
   - [ ] Parameters validated
   - [ ] Progress messages added
   - [ ] Functions documented
   - [ ] Examples in help text
   - [ ] No hardcoded values
   - [ ] No secrets in code

### 3. Code Refinement Process

**For improvements to existing code:**

1. **Document current behavior** - Understand what currently exists
2. **Identify specific improvements** - Be clear about what will change
3. **Ensure backwards compatibility** - Don't break existing usage
4. **Update tests** - Reflect the refinements
5. **Update documentation** - Keep docs in sync with code

**Backwards Compatibility:**
```powershell
# ✅ GOOD: Add optional parameter (backwards compatible)
[Parameter(Mandatory = $false)]
[string]$NewFeature

# ✅ GOOD: Deprecate gracefully with warning
if ($OldParameter) {
    Write-Warning "The -OldParameter is deprecated. Use -NewParameter instead."
    $NewParameter = $OldParameter
}

# ❌ BAD: Remove or rename parameters without notice
# This breaks existing scripts using the tool
```

---

## Testing Requirements

### Testing Standards

**All new features and refinements must include:**

1. **Manual testing** in a test environment
2. **Test scenarios** documented
3. **Edge cases** considered
4. **Error conditions** tested

### Test Environment Setup

**Minimum test setup:**
- Test SharePoint site (non-production)
- Test users (at least 3)
- Sample content (lists, libraries, items)
- Valid Azure AD app registration

### Test Documentation

Create test scenarios in this format:

```markdown
## Test Scenario: [Feature Name]

**Objective**: Verify [what you're testing]

**Prerequisites**:
- Test site URL: https://test.sharepoint.com/sites/TestSite
- Test users: user1@test.com, user2@test.com

**Steps**:
1. Action 1
2. Action 2
3. Action 3

**Expected Results**:
- ✅ Result 1
- ✅ Result 2
- ✅ Result 3

**Actual Results**:
[Document what actually happened]

**Pass/Fail**: [Pass/Fail with notes]
```

### Testing Checklist

Before submitting code:

- [ ] **Happy path** - Feature works as expected
- [ ] **Invalid input** - Proper error messages
- [ ] **Missing files** - Handles gracefully
- [ ] **Authentication failures** - Clear guidance
- [ ] **Large data sets** - Performance acceptable
- [ ] **Empty/null values** - No crashes
- [ ] **Special characters** - Handled correctly
- [ ] **Existing functionality** - Not broken

---

## Documentation Standards

### Required Documentation

**For every new feature:**

1. **Update README.md** with:
   - Feature description
   - Usage examples
   - Parameters reference
   - Common scenarios

2. **Update inline help** (`Get-Help` content)
   - Synopsis
   - Description
   - All parameters
   - At least 2 examples

3. **Create/update test guide** if applicable

4. **Add to quick reference** if user-facing feature

### Documentation Style

**Use consistent formatting:**

```markdown
## Feature Name

Brief description of the feature.

### Usage

```powershell
.\Script.ps1 -Parameter "Value"
```

### Parameters

- `-Parameter` (Required): Description of parameter

### Examples

#### Example 1: Basic Usage
```powershell
.\Script.ps1 -Parameter "Value"
```
This does X and Y.

#### Example 2: Advanced Usage
```powershell
.\Script.ps1 -Parameter "Value" -Option1 -Option2
```
This does X, Y, and Z.

### Troubleshooting

**Issue**: Error message
**Solution**: How to fix it
```

---

## Commit Guidelines

### Commit Message Format

Use [Conventional Commits](https://www.conventionalcommits.org/) format:

```
<type>(<scope>): <subject>

<body>

<footer>
```

**Types:**
- `feat:` - New feature
- `fix:` - Bug fix
- `docs:` - Documentation only
- `refactor:` - Code refactoring (no functional changes)
- `perf:` - Performance improvement
- `test:` - Adding or updating tests
- `chore:` - Maintenance tasks

**Examples:**

```
feat(user-mapping): Add bulk user validation function

- Add Test-BulkUserMapping function
- Validate up to 500 users in parallel
- Improve performance for large migrations

Closes #42
```

```
fix(import): Handle null user references in people picker fields

- Check for null before accessing user properties
- Add defensive null checks in Apply-UserMappingToTemplate
- Prevent NullReferenceException during import

Fixes #58
```

```
docs(readme): Add troubleshooting section for certificate errors

- Document common certificate-related errors
- Provide solutions for expired certificates
- Add links to certificate renewal process
```

### Commit Best Practices

- **One logical change per commit** - Don't mix unrelated changes
- **Commit often** - Small, focused commits are better than large ones
- **Write clear messages** - Future you will thank present you
- **Reference issues** - Use `Closes #123` or `Fixes #456`

---

## Pull Request Process

### Before Submitting

- [ ] Code follows style guide
- [ ] All tests pass
- [ ] Documentation updated
- [ ] Commit messages follow guidelines
- [ ] No merge conflicts
- [ ] Self-reviewed the changes

### PR Template

```markdown
## Description
Brief description of the changes

## Related Issue
Closes #[issue number]

## Type of Change
- [ ] Bug fix (non-breaking change which fixes an issue)
- [ ] New feature (non-breaking change which adds functionality)
- [ ] Breaking change (fix or feature that would cause existing functionality to not work as expected)
- [ ] Documentation update

## Testing Performed
- [ ] Manual testing in test environment
- [ ] Edge cases tested
- [ ] Error conditions verified
- [ ] Existing functionality validated

### Test Environment
- PowerShell Version: [e.g., 7.4]
- PnP.PowerShell Version: [e.g., 2.12.0]
- Tenant Type: [e.g., Commercial, GCC]

### Test Results
[Describe what was tested and results]

## Documentation
- [ ] README.md updated
- [ ] Inline help updated
- [ ] Examples provided
- [ ] Test guide updated (if applicable)

## Checklist
- [ ] My code follows the style guidelines of this project
- [ ] I have performed a self-review of my own code
- [ ] I have commented my code, particularly in hard-to-understand areas
- [ ] I have made corresponding changes to the documentation
- [ ] My changes generate no new warnings
- [ ] I have tested my changes in a test environment
- [ ] No secrets or sensitive data in commits

## Screenshots (if applicable)
[Add screenshots of new features or UI changes]

## Additional Notes
[Any additional information that reviewers should know]
```

### Review Process

**All pull requests require:**
1. Code review by at least one maintainer
2. Documentation review
3. Testing validation
4. No unresolved comments before merge

---

## Common Patterns

### Authentication Pattern

All scripts should use this consistent authentication approach:

```powershell
function Connect-SharePoint {
    param(
        [string]$SiteUrl,
        [string]$ConfigFilePath,
        [string]$ClientIdParam,
        [string]$TenantParam
    )
    
    $configPath = Join-Path $PSScriptRoot $ConfigFilePath
    if (Test-Path $configPath) {
        try {
            $config = Get-Content $configPath | ConvertFrom-Json
            
            # Try certificate auth first (modern auth)
            if ($config.clientId -and $config.certificateThumbprint -and $config.tenantId) {
                Connect-PnPOnline -Url $SiteUrl -ClientId $config.clientId `
                    -Thumbprint $config.certificateThumbprint -Tenant $config.tenantId `
                    -WarningAction SilentlyContinue
                return
            }
            # Fall back to client secret (ACS)
            elseif ($config.clientId -and $config.clientSecret) {
                Connect-PnPOnline -Url $SiteUrl -ClientId $config.clientId `
                    -ClientSecret $config.clientSecret -WarningAction SilentlyContinue
                return
            }
        }
        catch {
            Write-ProgressMessage "App registration auth failed: $($_.Exception.Message)" -Type "Warning"
        }
    }
    
    # Fall back to interactive
    Connect-PnPOnline -Url $SiteUrl -Interactive -WarningAction SilentlyContinue
}
```

### Configuration Validation Pattern

```powershell
function Test-ConfigurationFile {
    param([string]$ConfigFilePath)
    
    $fullPath = if ([System.IO.Path]::IsPathRooted($ConfigFilePath)) {
        $ConfigFilePath
    } else {
        Join-Path $PSScriptRoot $ConfigFilePath
    }
    
    if (-not (Test-Path $fullPath)) {
        Write-Host "ERROR: Configuration file not found: $ConfigFilePath" -ForegroundColor Red
        # Provide helpful guidance
        return $false
    }
    
    try {
        $config = Get-Content $fullPath -Raw | ConvertFrom-Json
        
        # Validate required fields
        $missingFields = @()
        if (-not $config.tenantId) { $missingFields += "tenantId" }
        if (-not $config.clientId) { $missingFields += "clientId" }
        
        if ($missingFields.Count -gt 0) {
            Write-Host "ERROR: Missing required fields: $($missingFields -join ', ')" -ForegroundColor Red
            return $false
        }
        
        return $true
    }
    catch {
        Write-Host "ERROR: Invalid JSON in configuration file" -ForegroundColor Red
        return $false
    }
}
```

### Progress Reporting Pattern

```powershell
$startTime = Get-Date

Write-ProgressMessage "Starting operation..." -Type "Info"

# Perform work
# ...

$endTime = Get-Date
$duration = $endTime - $startTime

Write-ProgressMessage "Operation completed in $($duration.TotalMinutes.ToString('0.00')) minutes" -Type "Success"
```

---

## Security Guidelines

### Never Commit Secrets

**Forbidden in commits:**
- ❌ Passwords
- ❌ Client secrets
- ❌ Certificates (private keys)
- ❌ Personal Access Tokens
- ❌ API keys
- ❌ Tenant IDs (in public repos)
- ❌ Real user emails/data

**Allowed:**
- ✅ Certificate thumbprints (public identifier)
- ✅ Example/placeholder values
- ✅ Sample configuration templates
- ✅ Public documentation

### Configuration File Security

```powershell
# ✅ GOOD: Use configuration files (git-ignored)
$config = Get-Content "app-config.json" | ConvertFrom-Json
Connect-PnPOnline -ClientId $config.clientId -Thumbprint $config.certificateThumbprint

# ❌ BAD: Hardcoded secrets
Connect-PnPOnline -ClientId "abc123" -Thumbprint "def456"
```

### Input Validation

**Always validate user input:**

```powershell
# URL validation
if ($SiteUrl -notmatch '^https://[^/]+\.sharepoint\.com/.*$') {
    throw "Invalid SharePoint URL format"
}

# File path validation
if (-not (Test-Path $FilePath)) {
    throw "File not found: $FilePath"
}

# Prevent path traversal
$safePath = [System.IO.Path]::GetFullPath($FilePath)
if (-not $safePath.StartsWith($AllowedBasePath)) {
    throw "Path traversal detected"
}
```

### Secure Credential Handling

```powershell
# ✅ GOOD: Use SecureString for sensitive data
$securePassword = Read-Host "Enter password" -AsSecureString

# ✅ GOOD: Use certificate-based auth
Connect-PnPOnline -ClientId $clientId -Thumbprint $thumbprint

# ❌ BAD: Plain text passwords
$password = "MyPassword123"
```

---

## Versioning and Releases

### Version Numbering

Follow [Semantic Versioning](https://semver.org/): `MAJOR.MINOR.PATCH`

- **MAJOR**: Breaking changes
- **MINOR**: New features (backwards compatible)
- **PATCH**: Bug fixes (backwards compatible)

### Release Checklist

- [ ] All tests pass
- [ ] Documentation complete
- [ ] CHANGELOG.md updated
- [ ] Version numbers updated
- [ ] Tag created: `git tag v1.2.0`
- [ ] Release notes written

---

## Getting Help

### Resources

- **Documentation**: Start with [README.md](README.md)
- **Issues**: Check existing issues for similar problems
- **Discussions**: Use GitHub Discussions for questions
- **Examples**: Review existing scripts for patterns

### Asking Questions

**Good question format:**

```markdown
## Question: [Brief Summary]

### Context
What are you trying to accomplish?

### What I've Tried
- Step 1
- Step 2

### Environment
- PowerShell Version: 7.4
- PnP.PowerShell Version: 2.12.0
- Operating System: Windows 11

### Error Message (if applicable)
```
[Paste error message]
```

### Expected Behavior
What should happen?

### Actual Behavior
What actually happens?
```

---

## Code Review Guidelines

### For Reviewers

**What to look for:**
- ✅ Code follows style guide
- ✅ Proper error handling
- ✅ Clear variable names
- ✅ Helpful error messages
- ✅ Documentation updated
- ✅ No security issues
- ✅ No hardcoded values
- ✅ Backwards compatible (or documented breaking change)

**Providing feedback:**
- Be constructive and respectful
- Explain the "why" behind suggestions
- Provide code examples when possible
- Distinguish between "must fix" and "nice to have"

### For Contributors

**Receiving feedback:**
- Feedback is about the code, not you personally
- Ask questions if something is unclear
- Don't take it personally
- Learn from the feedback

---

## Contributing Checklist

Before submitting any code:

- [ ] Read and understand this developer guide
- [ ] Code follows PowerShell style guide
- [ ] Proper error handling implemented
- [ ] Parameters validated
- [ ] Progress messages added
- [ ] Functions documented with help text
- [ ] Examples provided
- [ ] No hardcoded values or secrets
- [ ] Tested in test environment
- [ ] Documentation updated
- [ ] Commit messages follow guidelines
- [ ] Self-reviewed code
- [ ] Ready for code review

---

## Quick Reference

### File Naming Conventions

- Scripts: `Verb-Noun.ps1` (e.g., `Export-SharePointSiteTemplate.ps1`)
- Documentation: `UPPERCASE-NAME.md` (e.g., `README.md`, `DEVELOPER.md`)
- Config files: `lowercase-with-dashes.json` (e.g., `app-config.json`)
- Sample files: `name.sample.ext` (e.g., `user-mapping.sample.csv`)

### Useful Commands

```powershell
# Validate PowerShell script
Test-ScriptFileInfo -Path .\Script.ps1

# Get help on a script
Get-Help .\Script.ps1 -Full

# Check code style with PSScriptAnalyzer
Invoke-ScriptAnalyzer -Path .\Script.ps1

# Format code (manually or with formatter)
# Use consistent indentation (4 spaces)

# Test in WhatIf mode
.\Script.ps1 -WhatIf
```

---

## Contact and Support

- **Maintainers**: [List maintainers here]
- **Issues**: GitHub Issues
- **Discussions**: GitHub Discussions
- **Email**: [Support email if applicable]

---

**Document Version**: 1.0  
**Last Updated**: February 4, 2026  
**Maintained By**: Project Maintainers

---

Thank you for contributing to this project! Your efforts help improve SharePoint migration capabilities for MSPs and IT teams everywhere. 🚀
