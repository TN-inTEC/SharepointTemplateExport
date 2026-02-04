# Configuration Files

This directory contains configuration file templates for authenticating with SharePoint.

## Quick Setup

1. **Copy the sample file:**
   ```powershell
   Copy-Item app-config.sample.json app-config.json
   ```

2. **Edit `app-config.json` with your Azure AD app credentials:**
   - `tenantId`: Your Azure AD tenant ID (GUID)
   - `clientId`: Your app registration client ID (GUID)
   - `certificateThumbprint`: Your certificate thumbprint (40 hex characters)
   - `tenantDomain`: Your tenant domain (e.g., contoso.onmicrosoft.com)
   - `clientSecret`: (Optional) Client secret for fallback authentication

3. **For cross-tenant scenarios**, create separate config files:
   - `app-config-source.json` - For the source tenant
   - `app-config-target.json` - For the target tenant

## Security Notes

⚠️ **IMPORTANT**: Never commit actual config files to source control!

- `app-config*.json` files (except examples) are ignored by git
- Keep your credentials secure
- Rotate certificates before expiration
- Use certificate authentication (more secure than client secrets)

## Required Values

### Finding Your Tenant ID
```powershell
# Option 1: Azure Portal
# Go to Azure Active Directory → Properties → Tenant ID

# Option 2: PowerShell
Connect-AzAccount
Get-AzTenant
```

### Finding Your Client ID
```powershell
# Azure Portal: 
# Azure Active Directory → App registrations → Your App → Application (client) ID
```

### Finding Your Certificate Thumbprint
```powershell
# List certificates in your personal store
Get-ChildItem Cert:\CurrentUser\My | Select-Object Subject, Thumbprint, NotAfter
```

## Example Files

- `app-config.sample.json` - Template with placeholder values
- `app-config-source.example.json` - Example for source tenant
- `app-config-target.example.json` - Example for target tenant

## Troubleshooting

If you see configuration errors when running scripts:
1. Verify all required fields are present
2. Check certificate exists: `Get-ChildItem Cert:\CurrentUser\My`
3. Ensure thumbprint matches exactly (no spaces)
4. Verify tenantId and clientId are valid GUIDs
5. See MANUAL-APP-REGISTRATION.md for detailed setup

## Authentication Priority

The scripts try authentication methods in this order:
1. **Certificate** (recommended) - Modern auth, most secure
2. **Client Secret** (fallback) - Requires ACS enabled
3. **Interactive** (last resort) - May be blocked by Conditional Access
