<#
  .SYNOPSIS
  Get last sign-in activity of all users in the tenant.
  
  .DESCRIPTION
  This script gets last sign-in activity of all users in the tenant.
  This script requires Azure AD Premium license and Microsoft.Graph module version 1.27.0 or later.

  .PARAMETER CertificateThumbprint
  Thumbprint of the certificate used to connect to Graph API.

  .PARAMETER TenantId
  Tenant ID of the tenant to connect to.

  .Parameter ClientId
  Client ID of the application to connect to.

  .PARAMETER Outfile
  Output file path. Default is $env:USERPROFILE\Desktop\lastSignIns.csv

  .PARAMETER EnableLastSignInActivityDetail
  If this parameter is specified, this script gets last sign-in activity detail of all users in the tenant. This operation might take longer.

  .PARAMETER TimeZone
  Time zone to convert last sign-in date. example 'Tokyo Standard Time'. All available TimeZones can be found '[System.TimeZoneInfo]::GetSystemTimeZones() | Select-Object -ExpandProperty Id'

  .EXAMPLE
  .\Get-LastSignIn.ps1 -CertificateThumbprint "1234567890123456789012345678901234567890" -TenantId "12345678-1234-1234-1234-123456789012" -ClientId "12345678-1234-1234-1234-123456789012" -Outfile "C:\temp\lastSignIns.csv" -TimeZone "Tokyo Standard Time"
#>

Param(
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$CertificateThumbprint,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$TenantId,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$ClientId,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$Outfile = "$env:USERPROFILE\Desktop\lastSignIns.csv",
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$EnbaleLastSignInActivityDetail = $true,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$TimeZone
)

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Error "Microsoft.Graph module does not exist. Please run `"Install-Module -Name Microsoft.Graph`" command as local administrator"
    return;
} else {
    $latestModule = Get-Module -ListAvailable -Name Microsoft.Graph | Sort-Object -Property VErsion -Descending | Select-Object -First 1
    if ($latestModule.Version -lt [version]"1.27.0")
    {
        Write-Error "version 1.27.0 or later is required. Please run `"Update-Module -Name Microsoft.Graph`" command as local administrator"
        return;
    }
}

if ($TimeZone)
{
    #check timezone is available
    try {
        $timeZoneInfo = [System.TimeZoneInfo]::FindSystemTimeZoneById($TimeZone)        
    }
    catch {
        throw [System.ArgumentException]::new("Invalid TimeZone. example 'Tokyo Standard Time'. All available TimeZones can be found '[System.TimeZoneInfo]::GetSystemTimeZones() | Select-Object -ExpandProperty Id'")
    }
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

# Select-MgProfile command is no longer supported in Microsoft.Graph PowerShell v2. If you are using the v1 module and usually select the beta profile, please uncomment the line below to explicitly use the v1.0 endpoint. for more detail, see https://github.com/microsoftgraph/msgraph-sdk-powershell/blob/dev/docs/upgrade-to-v2.md
# Select-MgProfile -Name "v1.0"

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

$props = @(
    "Id", 
    "UserPrincipalName", 
    @{label = "LastSignInDateUTC"; expression = { $_.SignInActivity.lastSignInDateTime } },
    @{label = "AppDisplayName"; expression = { $_.LastSignInEvent.AppDisplayName } },
    @{label = "LastNonInteractiveSignInDateUTC"; expression = { $_.SignInActivity.lastNonInteractiveSignInDateTime} },
    @{label = "NonInteractiveAppDisplayName"; expression = { $_.LastNonInteractiveSignInEvent.AppDisplayName } }
)

if($TimeZone)
{
    $props += @{label = "LastSignInDate($TimeZone)"; expression = { [System.TimeZoneInfo]::ConvertTimeFromUtc($_.SignInActivity.lastSignInDateTime, $timeZoneInfo) } }
}

$users | Select-Object -Property $props | ConvertTo-Csv -NoTypeInformation `
| Out-File -Encoding utf8 -FilePath $Outfile

Write-Host "Finish! OutpuFile: $Outfile"  -BackgroundColor "Black" -ForegroundColor "Green" 
Disconnect-Graph
