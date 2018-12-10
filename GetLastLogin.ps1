Add-Type -Path ".\Tools\Microsoft.IdentityModel.Clients.ActiveDirectory\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"

#
# Authorization & resource Url
#
$tenantId = "contoso.onmicrosoft.com" 
$resource = "https://graph.microsoft.com" 
$clientID = "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXXX"
$client_secret = "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
$outfile = "C:\Users\xxxxx\Desktop\lastLogin.csv"

$data = @()
$users= @()


#
# Authorization & resource Url
#
$authUrl = "https://login.microsoftonline.com/$tenantId/" 

#
# Get certificate
#
$cert = Get-ChildItem -path cert:\CurrentUser\My | Where-Object {$_.Thumbprint -eq $client_secret}

#
# Create AuthenticationContext for acquiring token 
# 
$authContext = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext $authUrl, $false

#
# Create credential for client application 
#
$clientCred = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate $clientID, $cert

#
# Acquire the authentication result
#
$authResult = $authContext.AcquireTokenAsync($resource, $clientCred).Result 


if ($null -ne $authResult.AccessToken) {
    #
    # Compose the access token type and access token for authorization header
    #
    $headerParams = @{'Authorization' = "$($authResult.AccessTokenType) $($authResult.AccessToken)"}
    
    $reqUrl = "$resource/beta/users"
    do {
        $rData = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $reqUrl).Content | ConvertFrom-Json
        $users += $rData.value
        $reqUrl = ($rData.'@odata.nextLink') + ''
    }while ($reqUrl.IndexOf('https') -ne -1)
    

    $data += "UserPrincipalName,Last sign-in date in UTC (Last 30 days)"

    foreach ($user in $users) {
        $reqUrl = "$resource/beta/auditLogs/signIns?&`$filter=userPrincipalName eq '" + [System.Web.HttpUtility]::UrlEncode($user.UserPrincipalName) + "'&`$orderby=createdDateTime desc &`$top=1"
        $rData = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $reqUrl).Content | ConvertFrom-Json
        $data += $user.UserPrincipalName + "," + $rData.value[0].createdDateTime
    }
}
else {
    Write-Host "ERROR: No Access Token"
}

$data | Out-File $outfile -encoding "utf8"