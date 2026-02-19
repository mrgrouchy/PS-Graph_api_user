<#
.SYNOPSIS
    Manage Microsoft Graph API delegated permission scopes on the
    "Microsoft Graph Command Line Tools" service principal in Entra ID.

.DESCRIPTION
    Allows viewing, adding, and removing delegated OAuth2 permission grants
    (scopes) on the Microsoft Graph Command Line Tools enterprise app
    (AppId: 14d82eec-204b-4c2f-b7e8-296a70dab67e).

    Add/Remove operations merge or subtract from the existing scope list so
    no grants are accidentally wiped.

.PARAMETER Action
    View   - List all current delegated permission grants and their scopes.
    Add    - Add one or more scopes to the existing grant (creates the grant
             if it does not yet exist).
    Remove - Remove one or more scopes from the existing grant.

.PARAMETER Scopes
    Space- or comma-separated list of scope names to add or remove.
    Example: "User.Read", "Mail.Read,Calendars.Read", "User.Read Mail.Read"

.PARAMETER ConsentType
    AllPrincipals  (default) - Admin consent applied to all users.
    Principal                - Consent scoped to a specific user (requires
                               -PrincipalId).

.PARAMETER PrincipalId
    Object ID of the user when ConsentType is Principal.

.EXAMPLE
    # View all current grants
    .\main.ps1 -Action View

.EXAMPLE
    # Add scopes (admin consent for all users)
    .\main.ps1 -Action Add -Scopes "User.Read","Mail.Read"

.EXAMPLE
    # Remove a specific scope
    .\main.ps1 -Action Remove -Scopes "Mail.Read"
#>

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter(Mandatory)]
    [ValidateSet('View', 'Add', 'Remove')]
    [string] $Action,

    [Parameter()]
    [string[]] $Scopes,

    [Parameter()]
    [ValidateSet('AllPrincipals', 'Principal')]
    [string] $ConsentType = 'AllPrincipals',

    [Parameter()]
    [string] $PrincipalId
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# If PrincipalId is supplied without an explicit ConsentType, treat it as a user-specific grant
if ($PrincipalId -and -not $PSBoundParameters.ContainsKey('ConsentType')) {
    $ConsentType = 'Principal'
}

# ── Constants ──────────────────────────────────────────────────────────────────
$GraphCommandLineToolsAppId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$MicrosoftGraphAppId        = '00000003-0000-0000-c000-000000000000'

# ── Helper: Ensure Microsoft.Graph module is available ─────────────────────────
function Assert-GraphModule {
    if (-not (Get-Module -ListAvailable -Name 'Microsoft.Graph.Applications')) {
        Write-Host 'Microsoft.Graph module not found. Installing...' -ForegroundColor Yellow
        Install-Module Microsoft.Graph -Scope CurrentUser -Repository PSGallery -Force
    }
    Import-Module Microsoft.Graph.Applications -ErrorAction Stop
    Import-Module Microsoft.Graph.Users        -ErrorAction Stop
}

# ── Helper: Normalise a scope string/array into a sorted unique string set ─────
function ConvertTo-ScopeSet ([string] $scopeString) {
    return @($scopeString -split '\s+' |
             Where-Object { $_ -ne '' } |
             Select-Object -Unique |
             Sort-Object)
}

function Format-ScopeSet ([string[]] $set) {
    return ($set -join ' ')
}

# ── Helper: Resolve a user object ID to a UPN (falls back to the raw ID) ───────
function Resolve-PrincipalUpn ([string] $ObjectId) {
    try {
        $user = Get-MgUser -UserId $ObjectId -Property 'userPrincipalName' -ErrorAction Stop
        return $user.UserPrincipalName
    }
    catch {
        return $ObjectId   # not a user (e.g. group/SP), return raw ID
    }
}

# ── Helper: Find the Microsoft Graph Command Line Tools service principal ───────
function Get-ToolsServicePrincipal {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$GraphCommandLineToolsAppId'" -ErrorAction Stop
    if (-not $sp) {
        throw "Could not find the 'Microsoft Graph Command Line Tools' service principal. " +
              "Ensure it exists in your tenant (run Connect-MgGraph at least once)."
    }
    return $sp
}

# ── Helper: Find the Microsoft Graph resource service principal ─────────────────
function Get-GraphResourceServicePrincipal {
    $sp = Get-MgServicePrincipal -Filter "appId eq '$MicrosoftGraphAppId'" -ErrorAction Stop
    if (-not $sp) {
        throw "Could not find the Microsoft Graph resource service principal."
    }
    return $sp
}

# ── Helper: Get the OAuth2 permission grant for the given combination ──────────
function Get-PermissionGrant {
    param (
        [string] $ClientSpId,
        [string] $ResourceSpId,
        [string] $ConsentType,
        [string] $PrincipalId
    )

    # Raw REST call — bypasses SDK property-mapping issues entirely.
    $allGrants = @()
    $uri = "v1.0/servicePrincipals/$ClientSpId/oauth2PermissionGrants"
    do {
        $response   = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        $allGrants += $response.value
        $uri        = if ($response.ContainsKey('@odata.nextLink')) { $response.'@odata.nextLink' } else { $null }
    } while ($uri)

    Write-Verbose "Grant lookup: $($allGrants.Count) grants found for SP $ClientSpId"
    foreach ($g in $allGrants) {
        Write-Verbose "  id=$($g.id)  resourceId=$($g.resourceId)  consentType=$($g.consentType)  principalId=$($g.principalId)  scope=$($g.scope)"
    }

    $match = $allGrants | Where-Object {
        $_.resourceId  -eq $ResourceSpId -and
        $_.consentType -eq $ConsentType  -and
        ($ConsentType -ne 'Principal' -or $_.principalId -eq $PrincipalId)
    } | Select-Object -First 1

    if (-not $match) { return $null }

    # Normalise to PascalCase PSCustomObject so all callers keep working unchanged
    return [PSCustomObject]@{
        Id          = $match.id
        ConsentType = $match.consentType
        PrincipalId = $match.principalId
        ResourceId  = $match.resourceId
        Scope       = $match.scope
    }
}

# ── ACTION: View ───────────────────────────────────────────────────────────────
function Invoke-ViewAction {
    Write-Host "`nFetching service principals..." -ForegroundColor Cyan
    $toolsSp    = Get-ToolsServicePrincipal
    $graphResSp = Get-GraphResourceServicePrincipal

    Write-Host "Service principal : $($toolsSp.DisplayName)" -ForegroundColor Cyan
    Write-Host "Object ID         : $($toolsSp.Id)"
    Write-Host "App ID            : $($toolsSp.AppId)"
    Write-Host ""

    $rawGrants = @()
    $uri = "v1.0/servicePrincipals/$($toolsSp.Id)/oauth2PermissionGrants"
    do {
        $response   = Invoke-MgGraphRequest -Uri $uri -Method GET -ErrorAction Stop
        $rawGrants += $response.value
        $uri        = if ($response.ContainsKey('@odata.nextLink')) { $response.'@odata.nextLink' } else { $null }
    } while ($uri)

    $grants = $rawGrants |
              Where-Object {
                  $_.resourceId -eq $graphResSp.Id -and
                  ($ConsentType -eq 'AllPrincipals' -or $_.consentType -eq $ConsentType) -and
                  ($ConsentType -ne 'Principal' -or [string]::IsNullOrEmpty($PrincipalId) -or $_.principalId -eq $PrincipalId)
              } |
              ForEach-Object {
                  [PSCustomObject]@{
                      Id          = $_.id
                      ConsentType = $_.consentType
                      PrincipalId = $_.principalId
                      ResourceId  = $_.resourceId
                      Scope       = $_.scope
                  }
              }

    if (-not $grants) {
        Write-Host "No OAuth2 permission grants found for this service principal." -ForegroundColor Yellow
        return
    }

    foreach ($grant in $grants) {
        Write-Host "─────────────────────────────────────────────────────────────" -ForegroundColor DarkGray
        Write-Host "Grant ID      : $($grant.Id)"
        Write-Host "Consent Type  : $($grant.ConsentType)"
        if ($grant.ConsentType -eq 'Principal') {
            $upn = Resolve-PrincipalUpn -ObjectId $grant.PrincipalId
            Write-Host "Principal     : $upn"
            if ($upn -ne $grant.PrincipalId) {
                Write-Host "Principal ID  : $($grant.PrincipalId)" -ForegroundColor DarkGray
            }
        }
        Write-Host ""
        $scopeSet = ConvertTo-ScopeSet -scopeString $grant.Scope
        if ($scopeSet.Count -eq 0) {
            Write-Host "  (no scopes)" -ForegroundColor Yellow
        } else {
            Write-Host "Scopes ($($scopeSet.Count)):" -ForegroundColor Green
            $scopeSet | ForEach-Object { Write-Host "  · $_" }
        }
        Write-Host ""
    }
}

# ── ACTION: Add ────────────────────────────────────────────────────────────────
function Invoke-AddAction {
    param ([string[]] $NewScopes)

    if (-not $NewScopes -or $NewScopes.Count -eq 0) {
        throw "You must provide at least one scope with -Scopes when using -Action Add."
    }

    # Normalise input (support comma-separated values inside each element)
    $newScopeSet = @(($NewScopes -join ' ' -replace ',', ' ') -split '\s+' |
                     Where-Object { $_ -ne '' } |
                     Select-Object -Unique |
                     Sort-Object)

    Write-Host "`nFetching service principals..." -ForegroundColor Cyan
    $toolsSp    = Get-ToolsServicePrincipal
    $graphResSp = Get-GraphResourceServicePrincipal

    $grant = Get-PermissionGrant -ClientSpId $toolsSp.Id `
                                  -ResourceSpId $graphResSp.Id `
                                  -ConsentType $ConsentType `
                                  -PrincipalId $PrincipalId

    if ($grant) {
        # Merge: existing + new, deduped and sorted
        $existingSet = ConvertTo-ScopeSet -scopeString $grant.Scope
        $mergedSet   = @($existingSet + $newScopeSet | Select-Object -Unique | Sort-Object)
        $addedScopes = @($newScopeSet | Where-Object { $existingSet -notcontains $_ })

        if ($addedScopes.Count -eq 0) {
            Write-Host "All specified scopes already exist in the grant. No changes made." -ForegroundColor Yellow
            Write-Host "Current scopes: $($existingSet -join ', ')"
            return
        }

        $mergedString = Format-ScopeSet -set $mergedSet

        Write-Host ""
        Write-Host "Existing grant  : $($grant.Id)"
        Write-Host "Existing scopes : $($existingSet -join ', ')" -ForegroundColor DarkGray
        Write-Host "Adding scopes   : $($addedScopes -join ', ')" -ForegroundColor Green
        Write-Host "Merged scopes   : $($mergedSet -join ', ')" -ForegroundColor Cyan

        if ($PSCmdlet.ShouldProcess($grant.Id, "Update OAuth2 permission grant scope")) {
            Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id `
                -BodyParameter @{ scope = $mergedString } -ErrorAction Stop
            Write-Host "`nGrant updated successfully." -ForegroundColor Green
        }
    }
    else {
        # No existing grant – create a new one
        $scopeString = Format-ScopeSet -set $newScopeSet

        $body = @{
            clientId    = $toolsSp.Id
            consentType = $ConsentType
            resourceId  = $graphResSp.Id
            scope       = $scopeString
        }

        if ($ConsentType -eq 'Principal') {
            if (-not $PrincipalId) {
                throw "-PrincipalId is required when ConsentType is 'Principal'."
            }
            $body['principalId'] = $PrincipalId
        }

        Write-Host ""
        Write-Host "No existing grant found. Creating new grant." -ForegroundColor Yellow
        Write-Host "Scopes to add   : $($newScopeSet -join ', ')" -ForegroundColor Green

        if ($PSCmdlet.ShouldProcess("New OAuth2PermissionGrant", "Create")) {
            $newGrant = New-MgOauth2PermissionGrant -BodyParameter $body -ErrorAction Stop
            Write-Host "`nGrant created successfully. ID: $($newGrant.Id)" -ForegroundColor Green
        }
    }
}

# ── ACTION: Remove ─────────────────────────────────────────────────────────────
function Invoke-RemoveAction {
    param ([string[]] $RemoveScopes)

    if (-not $RemoveScopes -or $RemoveScopes.Count -eq 0) {
        throw "You must provide at least one scope with -Scopes when using -Action Remove."
    }

    # Normalise input
    $removeScopeSet = @(($RemoveScopes -join ' ' -replace ',', ' ') -split '\s+' |
                        Where-Object { $_ -ne '' } |
                        Select-Object -Unique |
                        Sort-Object)

    Write-Host "`nFetching service principals..." -ForegroundColor Cyan
    $toolsSp    = Get-ToolsServicePrincipal
    $graphResSp = Get-GraphResourceServicePrincipal

    $grant = Get-PermissionGrant -ClientSpId $toolsSp.Id `
                                  -ResourceSpId $graphResSp.Id `
                                  -ConsentType $ConsentType `
                                  -PrincipalId $PrincipalId

    if (-not $grant) {
        Write-Host "No OAuth2 permission grant found for the specified combination. Nothing to remove." -ForegroundColor Yellow
        return
    }

    $existingSet  = ConvertTo-ScopeSet -scopeString $grant.Scope
    $notFound         = @($removeScopeSet | Where-Object { $existingSet -notcontains $_ })
    $remainingSet     = @($existingSet    | Where-Object { $removeScopeSet -notcontains $_ })
    $actuallyRemoving = @($removeScopeSet | Where-Object { $existingSet -contains $_ })

    Write-Host ""
    Write-Host "Existing grant   : $($grant.Id)"
    Write-Host "Existing scopes  : $($existingSet -join ', ')" -ForegroundColor DarkGray

    if ($notFound.Count -gt 0) {
        Write-Host "Not found (skip) : $($notFound -join ', ')" -ForegroundColor Yellow
    }

    if ($actuallyRemoving.Count -eq 0) {
        Write-Host "None of the specified scopes exist in the grant. No changes made." -ForegroundColor Yellow
        return
    }

    Write-Host "Removing scopes  : $($actuallyRemoving -join ', ')" -ForegroundColor Red

    if ($remainingSet.Count -eq 0) {
        Write-Host "Remaining scopes : (none – grant will be deleted)" -ForegroundColor Yellow

        if ($PSCmdlet.ShouldProcess($grant.Id, "Delete OAuth2 permission grant (no scopes remaining)")) {
            Remove-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id -ErrorAction Stop
            Write-Host "`nGrant deleted (no scopes remained)." -ForegroundColor Green
        }
    }
    else {
        $remainingString = Format-ScopeSet -set $remainingSet
        Write-Host "Remaining scopes : $($remainingSet -join ', ')" -ForegroundColor Cyan

        if ($PSCmdlet.ShouldProcess($grant.Id, "Update OAuth2 permission grant scope")) {
            Update-MgOauth2PermissionGrant -OAuth2PermissionGrantId $grant.Id `
                -BodyParameter @{ scope = $remainingString } -ErrorAction Stop
            Write-Host "`nGrant updated successfully." -ForegroundColor Green
        }
    }
}

# ── Main ───────────────────────────────────────────────────────────────────────

Assert-GraphModule

# Verify we are connected; prompt if not
try {
    $ctx = Get-MgContext -ErrorAction Stop
    if (-not $ctx) { throw }
}
catch {
    Write-Host "Not connected to Microsoft Graph. Connecting..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes 'Application.Read.All', 'DelegatedPermissionGrant.ReadWrite.All', 'User.Read.All' -NoWelcome
}

switch ($Action) {
    'View'   { Invoke-ViewAction }
    'Add'    { Invoke-AddAction   -NewScopes    $Scopes }
    'Remove' { Invoke-RemoveAction -RemoveScopes $Scopes }
}
