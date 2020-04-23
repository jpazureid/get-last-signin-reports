Param(
    [Parameter(ValueFromPipeline = $true,mandatory=$true)][ValidateSet("Cert", "Key")][String]$authMethod,
    [Parameter(ValueFromPipeline = $true,mandatory=$true)][String]$clientSecretOrThumbprint
)

# Authorization & resource Url
$tenantId = "contoso.onmicrosoft.com" 
$resource = "https://graph.microsoft.com" 
$clientID = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXXX"
$outfile = "C:\Users\xxxxx\Desktop\lastLogin.csv"
$authUrl = "https://login.microsoftonline.com/$tenantId/" 
$data = @()

switch ($authMethod)
{
    "cert" {
        Add-Type -Path ".\Tools\Microsoft.IdentityModel.Clients.ActiveDirectory\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"

        # Get certificate
        $cert = Get-ChildItem -path cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq $clientSecretOrThumbprint}
    
        # Create AuthenticationContext for acquiring token  
        $authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext $authUrl, $false
    
        # Create credential for client application 
        $clientCred = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate $clientID, $cert
    
        # Acquire the authentication result
        $authResult = $authContext.AcquireTokenAsync($resource, $clientCred).Result
        $accessTokenType = $authResult.AccessTokenType
        $accessToken = $authResult.AccessToken
    }
    "key" {
        $postParams = @{
            client_id     = $clientID; 
            client_secret = $clientSecretOrThumbprint;
            grant_type    = 'client_credentials';
            resource      = $resource
        }
        $authResult = (Invoke-WebRequest -Uri ($authUrl + "oauth2/token") -Method POST -Body $postParams) | ConvertFrom-Json
        $accessTokenType = $authResult.token_type
        $accessToken = $authResult.access_token
    }
}

if ($null -ne $accessToken) {
    #
    # Compose the access token type and access token for authorization header
    #
    $headerParams = @{'Authorization' = "$($accessTokenType) $($accessToken)"}
    
    $reqUrl = "$resource/v1.0/users"
    do {
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
}
else {
    Write-Host "ERROR: No Access Token"
}
