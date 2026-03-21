<#
.SYNOPSIS
    Syncs user profile photos from Microsoft Graph to Dynamics 365 Dataverse.

.DESCRIPTION
    For each active systemuser with an Azure AD Object ID, this script:
    1. Fetches their profile photo from Microsoft Graph
    2. Writes it to the systemuser entityimage field in Dataverse
    
    Uses device code flow for interactive authentication (no app registration required).
    Run this script whenever you add new users or want to refresh profile pictures.

.PARAMETER OrgUrl
    The Dataverse organization URL (e.g., https://yourorg.crm.dynamics.com)

.PARAMETER TenantId
    The Azure AD Tenant ID

.EXAMPLE
    .\SyncPhotos.ps1
    .\SyncPhotos.ps1 -OrgUrl "https://myorg.crm.dynamics.com" -TenantId "your-tenant-id"
#>
param(
    [string]$OrgUrl   = "https://mauriciomaster.crm.dynamics.com",
    [string]$TenantId = "48ac8550-da32-403e-9d2c-d280efe32983"
)

$ErrorActionPreference = "Stop"

# First-party client IDs for device code flow (no app registration needed)
$DataverseClientId = "51f81489-12ee-4a9e-aaae-a2591f45987d"   # Azure PowerShell
$GraphClientId     = "14d82eec-204b-4c2f-b7e8-296a70dab67e"   # Microsoft Graph CLI

function Get-DeviceCodeToken {
    param(
        [string]$TenantId,
        [string]$ClientId,
        [string]$Resource
    )

    $body = "client_id=$ClientId&resource=$Resource&grant_type=device_code"
    $dc   = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/devicecode" `
                              -Method Post -Body $body -ContentType "application/x-www-form-urlencoded"

    Write-Host "`n$($dc.message)" -ForegroundColor Yellow
    Write-Host "Waiting for authentication..." -ForegroundColor Cyan

    $pollBody = "client_id=$ClientId&grant_type=device_code&code=$($dc.device_code)"

    while ($true) {
        Start-Sleep -Seconds 5
        try {
            $tok = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantId/oauth2/token" `
                                     -Method Post -Body $pollBody -ContentType "application/x-www-form-urlencoded" `
                                     -ErrorAction Stop
            return $tok.access_token
        } catch {
            $err = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
            if ($err.error -eq "authorization_pending") { continue }
            throw $_
        }
    }
}

# ── Authentication ───────────────────────────────────────────────────────────
Write-Host "`n=== Step 1/2: Authenticate to Dataverse ===" -ForegroundColor Green
$dvToken = Get-DeviceCodeToken -TenantId $TenantId -ClientId $DataverseClientId -Resource $OrgUrl
Write-Host "Dataverse token acquired.`n" -ForegroundColor Green

Write-Host "=== Step 2/2: Authenticate to Microsoft Graph ===" -ForegroundColor Green
$graphToken = Get-DeviceCodeToken -TenantId $TenantId -ClientId $GraphClientId -Resource "https://graph.microsoft.com"
Write-Host "Graph token acquired.`n" -ForegroundColor Green

# ── Headers ──────────────────────────────────────────────────────────────────
$dvHeaders = @{
    Authorization      = "Bearer $dvToken"
    "OData-MaxVersion" = "4.0"
    "OData-Version"    = "4.0"
    Accept             = "application/json"
}
$graphHeaders = @{ Authorization = "Bearer $graphToken" }

$base = "$OrgUrl/api/data/v9.2"

# ── Fetch active users ───────────────────────────────────────────────────────
Write-Host "=== Fetching active users with AAD Object IDs... ===" -ForegroundColor Green
$filter   = "isdisabled eq false and azureactivedirectoryobjectid ne null"
$select   = "systemuserid,fullname,azureactivedirectoryobjectid"
$usersUrl = "$base/systemusers?`$select=$select&`$filter=$filter"
$users    = (Invoke-RestMethod -Uri $usersUrl -Headers $dvHeaders).value
Write-Host "Found $($users.Count) active users.`n" -ForegroundColor Cyan

# ── Sync photos ──────────────────────────────────────────────────────────────
Write-Host "=== Syncing photos... ===`n" -ForegroundColor Green
$synced  = 0
$noPhoto = 0
$errors  = 0

foreach ($u in $users) {
    $name  = $u.fullname
    $aadId = $u.azureactivedirectoryobjectid
    $sysId = $u.systemuserid

    try {
        $photoResp  = Invoke-WebRequest -Uri "https://graph.microsoft.com/v1.0/users/$aadId/photo/`$value" `
                                        -Headers $graphHeaders -ErrorAction Stop
        $photoBytes = [byte[]]$photoResp.Content
        $photoB64   = [Convert]::ToBase64String($photoBytes)

        $patchBody    = @{ entityimage = $photoB64 } | ConvertTo-Json -Compress
        $patchHeaders = $dvHeaders.Clone()
        $patchHeaders["Content-Type"] = "application/json"

        Invoke-RestMethod -Uri "$base/systemusers($sysId)" -Method Patch `
                          -Headers $patchHeaders -Body $patchBody

        $kb = [Math]::Round($photoBytes.Length / 1024, 1)
        Write-Host "  SYNCED: $name ($kb KB)" -ForegroundColor Green
        $synced++
    } catch {
        $status = $_.Exception.Response.StatusCode
        if ($status -eq 'NotFound' -or [int]$status -eq 404) {
            Write-Host "  SKIP:  $name (no photo in Graph)" -ForegroundColor DarkGray
            $noPhoto++
        } else {
            Write-Host "  ERROR: $name - $($_.Exception.Message)" -ForegroundColor Red
            $errors++
        }
    }
}

# ── Summary ──────────────────────────────────────────────────────────────────
Write-Host "`n=== Summary ===" -ForegroundColor Green
Write-Host "  Synced:   $synced"  -ForegroundColor Green
Write-Host "  No photo: $noPhoto" -ForegroundColor DarkGray
$errColor = if ($errors -gt 0) { "Red" } else { "Green" }
Write-Host "  Errors:   $errors"  -ForegroundColor $errColor
Write-Host ""
