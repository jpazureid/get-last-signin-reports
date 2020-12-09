Param(
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][ValidateSet("Cert", "Key")][String]$authMethod,
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][String]$clientSecretOrThumbprint,
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][String]$tenantId,
    [Parameter(ValueFromPipeline = $true, mandatory = $true)][String]$clientId,
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$resource = "https://graph.microsoft.com",
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$outfile = "$env:USERPROFILE\Desktop\lastSignIns.csv",
    [Parameter(ValueFromPipeline = $true, mandatory = $false)][String]$isPremiumTenant = $true
)

# Authorization & resource Url
$data = @()

$scope = "$resource/.default"
$scopes = New-Object System.Collections.ObjectModel.Collection["string"]
$scopes.Add($scope)

Function Get-AccessToken() {
    if ($null -eq $local:confidentialApp) {
        Add-Type -Path "Tools\Microsoft.Identity.Client\Microsoft.Identity.Client.dll"
        switch ($authMethod) {
            "cert" {
                # Get certificate
                $cert = Get-ChildItem -path cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $clientSecretOrThumbprint }
            
                # Create credential Application
                $local:confidentialApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientId).WithCertificate($cert).withTenantId($tenantId).Build()
            }
            "key" {
                $local:confidentialApp = [Microsoft.Identity.Client.ConfidentialClientApplicationBuilder]::Create($clientId).WithClientSecret($clientSecretOrThumbprint).withTenantId($tenantId).Build()
            }
        }
    }
    # Acquire the authentication result
    # ConfidentialClientApplication return token from cache if it valid.
    $authResult = $local:confidentialApp.AcquireTokenForClient($scopes).ExecuteAsync().Result
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


$reqUrl = "$resource/beta/users/?`$select=userPrincipalName,signInActivity"

do {
    #
    # Get data of all user's last sign-in activity events
    #
    $headerParams = Get-AuthorizationHeader
    $signInActivityJson = (Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $reqUrl).Content | ConvertFrom-Json
    $signInActivityJsonValue = $signInActivityJson.Value
    $numEvents = $signInActivityJsonValue.length

    #
    # Create title for the out put file
    #
    $data += "UserPrincipalName,Last sign-in event date in UTC,Cloud Application"
    #
    # Process data of each user's last sign-in activity event
    #
    for ($j = 0; $j -lt $numEvents; $j++) {

        #
        # Get user and event information
        #
        $userUPN = $signInActivityJsonValue.userPrincipalName[$j]
        $allSignInAct = $signInActivityJsonValue.signInActivity[$j]
        $lastRequestID = $allSignInAct.lastSignInRequestId
        $lastSignin = $allSignInAct.lastSignInDateTime
        #Write-Output "User number $j" "We have UPN is $userUPN"
        #Write-Output "We have Last is $lastSignin"
        #Write-Output "We have LastID is $lastRequestID"

        #
        # Check if event's request id. if id is null then it means the user never had a sctivity event.
        #
        if ($null -ne $lastRequestID -and $isPremiumTenant) {
            $eventReqUrl = "$resource/v1.0/auditLogs/signIns/$lastRequestID"
            #Write-Output "we have eventReqUrl is $eventReqUrl"

            try {
                $signInEventJsonRaw = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $eventReqUrl
                $signInEventJson = $signInEventJsonRaw.Content
                $signInEvent = ($signInEventJson | ConvertFrom-Json)
                $appDisplayName = $signInEvent.appDisplayName
            }
            catch [System.Net.WebException] {
                $statusCode = $_.Exception.Response.StatusCode.value__
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader $stream
                $reader.BaseStream.Position = 0
                $reader.DiscardBufferedData()
                $errorMessage = $reader.ReadToEnd()
                try {
                    $errorObj = ConvertFrom-Json $errorMessage
                }
                catch {
                    Write-Error "Unexpect error: $errorMessage. Continue..."
                }
                if ($statusCode -eq 403 -and ($errorObj.error -and $errorObj.error.code -eq "Authentication_RequestFromNonPremiumTenantOrB2CTenant")) {
                    Write-Error "This tenant doesn't have AAD Premium License or B2C tenant. To show App DisplayName, AAD Premium License is required. isPremiumTenant option is set to false."
                    $isPremiumTenant = $false
                } elseif ($statusCode -eq 404) {
                    #Write-Output "This user does not have sign in activity event in last 30 days."
                    $appDisplayName = "This user does not have sign in activity event in last 30 days."
                } else {
                    Write-Error "UnEpected Error: $errorMessage"
                }
            }

        }
        else {
            #Write-Output "This user never had sign in activity event."
            $lastSignin = "This user never had sign in activity event."
            $appDisplayName = $null
        }
        $data += $userUPN + "," + $lastSignin + "," + $appDisplayName
    }
    $reqUrl = ($signInActivityJson.'@odata.nextLink') + ''

} while ($reqUrl.IndexOf('https') -ne -1)

$data | Out-File $outfile -encoding "utf8"
