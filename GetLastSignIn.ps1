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
        if ($null -ne $lastRequestID) {
            $eventReqUrl = "$resource/v1.0/auditLogs/signIns/$lastRequestID"
            #Write-Output "we have eventReqUrl is $eventReqUrl"

            try {
                $signInEventJasonRaw = Invoke-WebRequest -UseBasicParsing -Headers $headerParams -Uri $eventReqUrl
            }
            catch [System.Net.WebException] {
                $requstStatusCode = $_.Exception.Response.StatusCode.value__
            }

            if ($requstStatusCode -eq 404) {
                #Write-Output "This user does not have sign in activity event in last 30 days."
                $appDisplayName = "This user does not have sign in activity event in last 30 days."
                $data += $userUPN + "," + $lastSignin + "," + $appDisplayName
            }
            else {
                $signInEventJason = $signInEventJasonRaw.Content
                #Write-Output "we have signInEventJason is $signInEventJason"

                $signInEvent = ($signInEventJason | ConvertFrom-Json)
                #Write-Output "we have signInEvent is $signInEvent"
                
                $appDisplayName = $signInEvent.appDisplayName
                #Write-Output "we have appDisplayName is $appDisplayName"

                $data += $userUPN + "," + $lastSignin + "," + $appDisplayName
            }
            $requstStatusCode = $null
        }
        else {
            #Write-Output "This user never had sign in activity event."
            $lastSignin = "This user never had sign in activity event."
            $appDisplayName = $null

            $data += $userUPN + "," + $lastSignin + "," + $appDisplayName
        }
    }
   
    $reqUrl = ($signInActivityJson.'@odata.nextLink') + ''

} while ($reqUrl.IndexOf('https') -ne -1)

$data | Out-File $outfile -encoding "utf8"
