Param(
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][ValidateSet("Cert", "Key")][String]$authMethod,
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][String]$clientSecretOrThumbprint
)

# Authorization & resource Url
$tenantId = "contoso.onmicrosoft.com" 
$resource = "https://graph.microsoft.com"
$scope = "$resource/.default" 
$clientID = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXXX"
$outfile = "$env:USERPROFILE\Desktop\lastLogin.csv"
$data = @()

$scopes = New-Object System.Collections.ObjectModel.Collection["string"]
$scopes.Add($scope)

Function Get-AccessToken() {
    switch ($authMethod) {
        "cert" {
            Add-Type -Path "Tools\Microsoft.Identity.Client\Microsoft.Identity.Client.dll"

            # Get certificate
            $cert = Get-ChildItem -path cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $clientSecretOrThumbprint }
        
            # Create credential Application
            $confidentialApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientID).WithCertificate($cert).withTenantId($tenantId).Build()
        }
        "key" {
            $confidentialApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientID).WithClientSecret($clientSecretOrThumbprint).withTenantId($tenantId).Build()
        }
    }
    # Acquire the authentication result
    $authResult = $confidentialApp.AcquireTokenForClient($scopes).ExecuteAsync().Result
    if ($null -eq $authResult) {
        Write-Host "ERROR: No Access Token"
        exit
    }    
    return $authResult
}

$authResult = Get-AccessToken
$accessToken = $authResult.AccessToken
#
# Compose the access token type and access token for authorization header
#
$headerParams = @{'Authorization' = "Bearer $($accessToken)" }        

$reqUrl = "$resource/v1.0/users"
do {
    if ($null -eq $authResult -or ($authResult.ExpiresOn -lt $(Get-Date).AddMinutes(-10))) {
        $authResult = Get-AccessToken
        $accessToken = $authResult.AccessToken
        #
        # Compose the access token type and access token for authorization header
        #
        $headerParams = @{'Authorization' = "Bearer $($accessToken)" }        
    }
    $rData = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $reqUrl).Content | ConvertFrom-Json
    $users += $rData.value
    $reqUrl = ($rData.'@odata.nextLink') + ''
}while ($reqUrl.IndexOf('https') -ne -1)

$data += "UserPrincipalName,Last sign-in date in UTC (Last 30 days)"
foreach ($user in $users) {
    $reqUrl = "$resource/v1.0/auditLogs/signIns?&`$filter=userId eq '" + $user.id + "'&`$orderby=createdDateTime desc &`$top=1"
    $rData = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $reqUrl).Content | ConvertFrom-Json
    $data += $user.UserPrincipalName + "," + $rData.value[0].createdDateTime
}

$data | Out-File $outfile -encoding "utf8"
