#####################################################
# HelloID-Conn-Target-Office365-License-Revoke
# Version: 1.0.0
#####################################################
$VerbosePreference = "continue"

# Initialize default value's
$success = $false
$config = $configuration | ConvertFrom-Json
$personObj = $person | ConvertFrom-Json
$pRef = $permissionReference | ConvertFrom-Json
$aRef = $accountReference | ConvertFrom-Json

if (-Not($dryRun -eq $true)) {
    try {
        $tokenUri = "https://login.microsoftonline.com/$($config.AADTenantID)/oauth2/token"

        $tokenHeaders = @{
            "content-type" = "application/x-www-form-urlencoded"
        }

        $body = @{
            grant_type    = "client_credentials"
            client_id     = $($config.AADAppId)
            client_secret = $($config.AADAppSecret)
            resource      = "https://graph.microsoft.com"
        }
        $accessToken = Invoke-RestMethod -Uri $tokenUri -Method Post -Headers $tokenHeaders -Body $body

        $upn = $aRef.UserPrincipalName
        $requestUri = "https://graph.microsoft.com/v1.0/users/$upn/assignLicense"
        $requestHeaders = @{
            Authorization  = "Bearer $($accessToken.access_token)";
            'Content-Type' = "application/json";
            Accept         = "application/json";
        }

        $body = @{
            addLicenses = @()
            removeLicenses= @($pRef.skuId)
        } | ConvertTo-Json -Depth 3

        write-verbose $body
        $request = Invoke-RestMethod -Uri $requestUri -Method Post -Headers $requestHeaders -Body $body
        if($request){
            $auditMessage = "Permission '$($pRef.Id)' added to account '$($aRef)'"
            $success = $true
        }
    }
    catch {
        if ( $($ex.Exception.GetType().FullName -eq 'Microsoft.PowerShell.Commands.HttpResponseException') -or $($ex.Exception.GetType().FullName -eq 'System.Net.WebException')){
            $stream = $ErrorObject.Exception.Response.GetResponseStream()
            $stream.Position = 0
            $streamReader = New-Object System.IO.StreamReader $Stream
            $errorResponse = $StreamReader.ReadToEnd()
            $errorMessage = ($errorResponse | ConvertFrom-Json).error.message
            $auditMessage = "Permission for '$($personObj.DisplayName)' not added. Error: $errorMessage"
        }
        else {
            $auditMessage = "Permission for '$($personObj.DisplayName)' not added. Error: $($ex.Exception.Message)"
        }
    }
}

# Send results
$result = [PSCustomObject]@{
    Success = $success
    AuditLogs = $auditLogs
}

Write-Output $result | ConvertTo-Json -Depth 10
