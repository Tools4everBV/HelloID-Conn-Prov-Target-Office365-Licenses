$config = ConvertFrom-Json $configuration
$p = $person | ConvertFrom-Json;
$AADTenantDomain = $config.AADTenantDomain
$AADTenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

$success = $false;
$auditMessage = "Azure identity for person " + $p.DisplayName + "(" +  $ADuserPrincipalName + ")" + " not found";
$accountReference = @{}

[string] $CorrelationuserPrincipalName = $p.Accounts.MicrosoftAzureAD.UserPrincipalName

if ([string]::IsNullOrEmpty($CorrelationuserPrincipalName))
{
    $AADuserPrincipalName = "$($p.name.GivenName).$($p.name.FamilyName)@$AADTenantDomain"

    #$AADuserPrincipalName = $CorrelationuserPrincipalName.Substring(0,$CorrelationuserPrincipalName.indexof('@')+1) + $AADtenantDomain
    $account = @{
        'UserPrincipalName' = $AADuserPrincipalName
    }
   }
else {
    $account = @{
        'UserPrincipalName' = $null
    }
    $auditMessage = "Unable to lookup Azure identity for person " + $p.DisplayName + "(" +  $CorrelationuserPrincipalName + ")" + " because no userPrincipalName was provided";

}

if ([Net.ServicePointManager]::SecurityProtocol -notmatch "Tls12") {
    [Net.ServicePointManager]::SecurityProtocol += [Net.SecurityProtocolType]::Tls12
}


#region functions
function Get-MicrosoftGraphToken(){
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true)]
        [string]
        $AADTenantID,

        [string]
        $AADAppId,

        [string]
        $AADAppSecret
    )
    try {
        $baseUri = "https://login.microsoftonline.com/"
        $authUri = $baseUri + "$AADTenantID/oauth2/token"

        $headers = @{
            "content-type" = "application/x-www-form-urlencoded"
        }

        $body = @{
            grant_type      = "client_credentials"
            client_id       = "$AADAppId"
            client_secret   = "$AADAppSecret"
            resource        = "https://graph.microsoft.com"
        }

        $splatRestMethodParameters = @{
            Uri = $authUri
            Method = 'POST'
            Headers = $headers
            Body = $body
        }
        $Response = Invoke-RestMethod @splatRestMethodParameters
        $accessToken = $Response.access_token;
    }
catch {
        if ($_.ErrorDetails) {
            $errorReponse = $_.ErrorDetails
        }
        if ($_.Exception.Response) {
            $result = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($result)
            $responseReader = $reader.ReadToEnd()
            $errorReponse = $responseReader | ConvertFrom-Json
            $reader.Dispose()
        }
        throw "Could not get MicrosoftGraph token for tenant: $AADTenantID  appID: $AADAppId , message: $($_.exception.message), $($errorReponse.error)"
    }

    return  $accessToken
}
function Invoke-MicrosoftGraphGetCommand{
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true)]
        [string]
        $accessToken,

        [string]
        $command
    )
    $baseUri = "https://graph.microsoft.com/v1.0/"
    $commandUri = $baseUri +  $Command

    #Add the authorization header to the request
    $headers = @{
        Authorization = "Bearer $accessToken";
        'Content-Type' = "application/json";
        Accept = "application/json";
    }
    $splatRestMethodParameters = @{
        Uri = $commandUri
        Method = 'GET'
        Headers = $headers
    }
    try{
    $azureADResponse = Invoke-RestMethod @splatRestMethodParameters
    }
    catch{
        if ($_.ErrorDetails) {
            $errorReponse = $_.ErrorDetails
        }
        if ($_.Exception.Response) {
            $result = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($result)
            $responseReader = $reader.ReadToEnd()
            $errorReponse = $responseReader | ConvertFrom-Json
            $reader.Dispose()
        }
        throw "Could not execute MicrosoftGraph command $commandUri  message: $($_.exception.message), $($errorReponse.error)"
    }
    return $azureADResponse
}
#endregion functions


try{
     $accessToken = Get-MicrosoftGraphToken -AADtenantID $AADTenantID -AADAppId $AADAppId -AADAppSecret $AADAppSecret

     if ($null -ne $account.UserPrincipalName)
     {
        $userPrincipalNameEncoded = [System.Web.HttpUtility]::UrlEncode($account.UserPrincipalName)
        $properties = @("id","displayName","userPrincipalName")
        $command = "users/$userPrincipalNameEncoded" + '?$select=' + ($properties -join ",")
        $azureADuserResponse = Invoke-MicrosoftGraphGetCommand -accessToken $accessToken -command $command

        $success = $true
        $auditMessage ="Azure identity for person " + $p.DisplayName + "(" +  $CorrelationuserPrincipalName + ")" + " successfully correlated";
     }

}
catch {
        $success = $false;
        $auditMessage = "Exception while looking up Azure identity for person " + $p.DisplayName + "(" +  $CorrelationuserPrincipalName + ") : ";
        $auditMessage += $($_.Exception.Message)
}

$result = [PSCustomObject]@{
    Success          = $success;
    AccountReference = @{
            Id = $azureADuserResponse.Id
            UserPrincipalName = $azureADuserResponse.UserPrincipalName
            DisplayName = $azureADuserResponse.displayName
    }
    AuditDetails     = $auditMessage;
    Account    = $account
};

#send result back
Write-Output $result | ConvertTo-Json -Depth 10
