Param(
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][ValidateSet("Cert", "Key")][String]$authMethod,
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][String]$clientSecretOrThumbprint,
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][String]$tenantId,
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][String]$clientId,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$resource = "https://graph.microsoft.com",
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$outfile = "$env:USERPROFILE\Desktop\lastSignIns.csv"
)

# Authorization & resource Url
$data = @()

$scope = "$resource/.default"
$scopes = New-Object System.Collections.ObjectModel.Collection["string"]
$scopes.Add($scope)

Function Get-AccessToken() {
    if ($null -eq $script:confidentialApp) {
        Add-Type -Path "Tools\Microsoft.IdentityModel.Abstractions\Microsoft.IdentityModel.Abstractions.dll"
        Add-Type -Path "Tools\Microsoft.Identity.Client\Microsoft.Identity.Client.dll"
        switch ($authMethod) {
            "cert" {
                # Get certificate
                $cert = Get-ChildItem -path cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $clientSecretOrThumbprint }
            
                # Create credential Application
                $script:confidentialApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientId).WithCertificate($cert).withTenantId($tenantId).Build()
            }
            "key" {
                $script:confidentialApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientId).WithClientSecret($clientSecretOrThumbprint).withTenantId($tenantId).Build()
            }
        }
    }
    # Acquire the authentication result
    # ConfidentialClientApplication return token from cache if it valid.
    $authResult = $script:confidentialApp.AcquireTokenForClient($scopes).ExecuteAsync().Result
    if ($null -eq $authResult) {
        Write-Host "ERROR: No Access Token"
        exit
    }
    return $authResult
}

Function Get-AuthorizationHeader {
    $authResult = Get-AccessToken
    $accessToken = $authResult.AccessToken    
    return @{'Authorization' = "Bearer $($accessToken)" }    
}


$reqUrl = "$resource/v1.0/users"
do {
    $headerParams = Get-AuthorizationHeader
    $rData = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $reqUrl).Content | ConvertFrom-Json
    $users += $rData.value
    $reqUrl = ($rData.'@odata.nextLink') + ''
}while ($reqUrl.IndexOf('https') -ne -1)

$data += "UserPrincipalName,Last sign-in date in UTC (Last 30 days)"
foreach ($user in $users) {
    $headerParams = Get-AuthorizationHeader
    $reqUrl = "$resource/v1.0/auditLogs/signIns?&`$filter=userId eq '" + $user.id + "'&`$orderby=createdDateTime desc &`$top=1"
    $rData = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $reqUrl).Content | ConvertFrom-Json
    $data += $user.UserPrincipalName + "," + $rData.value[0].createdDateTime
}

$data | Out-File $outfile -encoding "utf8"