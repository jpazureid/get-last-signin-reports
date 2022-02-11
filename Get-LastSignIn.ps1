Param(
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$CertificateThumbprint,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$TenantId,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$ClientId,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$Outfile = "$env:USERPROFILE\Desktop\lastSignIns.csv",
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$EnbaleLastSignInActivityDetail = $true
)

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Error "Microsoft.Graph module does not exist. Please run `"Install-Module -Name Microsoft.Graph`" command as local administrator"
    return;
}

try {
    Write-Host "Disconnect Graph..." -BackgroundColor "Black" -ForegroundColor "Green" 
    Disconnect-Graph -ErrorAction SilentlyContinue | Out-Null    
}
catch {}

try {
    Write-Host "Connecting Graph..." -BackgroundColor "Black" -ForegroundColor "Green" 
    if ("" -eq $CertificateThumbprint -or "" -eq $ClientId) {
        Write-Host "Client credentail is not provided. Connect-Graph as Administrator account..." -ForegroundColor Yellow
        if ($TenantId) {
            Connect-Graph -Scopes "User.Read.All, AuditLog.Read.All" -TenantId $TenantId
        }
        else {
            Connect-Graph -Scopes "Directory.Read.All, AuditLog.Read.All"
        }
    }
    else {
        Connect-Graph -ClientId $ClientId -CertificateThumbprint $CertificateThumbprint -TenantId $TenantId -ErrorAction Stop
    }
}
catch {
    Write-Error $_.Exception
    return;
}

# Use Beta API
Select-MgProfile -Name beta

try {
    # Get all users with ID, UPN and SignInActivity
    # Azure AD Premium lisence is required to complete this action
    Write-Host "Reading all users data... This operation might take longer..." -BackgroundColor "Black" -ForegroundColor "Green" 
    $users = Get-MgUser -All -Property id, userPrincipalName, signInActivity

    if ($EnbaleLastSignInActivityDetail) {
        Write-Host "Reading all signIn Activity data... This operation might take longer..." -BackgroundColor "Black" -ForegroundColor "Green" 
        $users | ForEach-Object {
            $activity = $_.SignInActivity
            $lastSignInRequestId = $activity.lastSignInRequestId
            $lastNonInteractiveSignInRequestId = $activity.lastNonInteractiveSignInRequestId

            if (($null -eq $lastSignInRequestId) -And ($null -eq $lastNonInteractiveSignInRequestId)) {
                return;
            }

            try {
                $lastSignInEvent = Get-MgAuditLogSignIn -SignInId $lastSignInRequestId -ErrorAction Stop
                $_ | Add-Member -MemberType NoteProperty -Name "LastSignInEvent" -value $lastSignInEvent

                $filter = "signInEventTypes/any(t: t eq 'nonInteractiveUser') and Id eq '" + $lastNonInteractiveSignInRequestId + "'"
                $lastNonInteractiveSignInEvent = Get-MgAuditLogSignIn -Filter $filter -ErrorAction Stop
                $_ | Add-Member -MemberType NoteProperty -Name "LastNonInteractiveSignInEvent" -value $lastNonInteractiveSignInEvent
            }
            catch {
                # Sign-in activities are stored for 30 days.
                # You can not check events older than 30 days.
                $ex = $_.Exception # Nothing to do...
            }
        }    
    }
}
catch {
    if ("Authentication_RequestFromNonPremiumTenantOrB2CTenant" -eq $_.Exception.Code) {
        Write-Host -ForegroundColor Yellow "This tenant doesn't have Azure AD Premimu License. Skipping load signin activities."
        throw;
    }
    else {
        throw
    }
}

Write-Host "Output data to CSV..."  -BackgroundColor "Black" -ForegroundColor "Green" 

# Show DateTime in UTC
$users | Select-Object Id, UserPrincipalName, @{label = "LastSignInDateUTC"; expression = { $_.SignInActivity.lastSignInDateTime } }, @{label = "AppDisplayName"; expression = { $_.LastSignInEvent.AppDisplayName } },@{label = "LastNonInteractiveSignInDateUTC"; expression = { $_.SignInActivity.lastNonInteractiveSignInDateTime} },@{label = "NonInteractiveAppDisplayName"; expression = { $_.LastNonInteractiveSignInEvent.AppDisplayName } }`
| ConvertTo-Csv -NoTypeInformation `
| Out-File -Encoding utf8 -FilePath $Outfile

# Show DateTime in JST
# $users | Select-Object Id, UserPrincipalName, @{label = "LastSignInDateJST"; expression = { $_.SignInActivity.lastSignInDateTime.AddHours(9) } }, @{label = "AppDisplayName"; expression = { $_.LastSignInEvent.AppDisplayName } },@{label = "LastNonInteractiveSignInDateJST"; expression = { $_.SignInActivity.lastNonInteractiveSignInDateTime.AddHours(9)} },@{label = "NonInteractiveAppDisplayName"; expression = { $_.LastNonInteractiveSignInEvent.AppDisplayName } }`
# | ConvertTo-Csv -NoTypeInformation `
# | Out-File -Encoding utf8 -FilePath $Outfile

Write-Host "Finish!"  -BackgroundColor "Black" -ForegroundColor "Green" 
Disconnect-Graph