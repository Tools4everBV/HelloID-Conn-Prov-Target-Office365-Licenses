$config = ConvertFrom-Json $configuration

$AADTenantDomain = $config.AADTenantDomain
$AADTenantID = $config.AADtenantID
$AADAppId = $config.AADAppId
$AADAppSecret = $config.AADAppSecret

$success = $false;
$auditMessage = "Azure licenses not collected successfully";

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

     $azureADlicenseResponse = Invoke-MicrosoftGraphGetCommand -accessToken $accessToken -command 'subscribedSkus'

     $Licenses = [System.Collections.Generic.List[psobject]]::new()

     if ($azureADlicenseResponse.value.length -gt 0)
     {
        $Licenses.addRange( [psobject[]] $azureADlicenseResponse.value)
     }

}
catch {

    throw "Could not get Office365 licenses (subscribedSkus), message: $($_.Exception.Message)"
}

$permissions = [System.Collections.Generic.List[psobject]]::new()
foreach ($Lic in $Licenses) {

    $permission = @{
        DisplayName    = $Lic.skuPartNumber
        Identification = @{
            Id = $Lic.Id
            SkuId = $Lic.SkuId
            ServicePlansToDisable = $null
        }
    }
    $permissions.add( $permission )

    # add custom permissions with selected disabled services

    switch  ($Lic.skuPartNumber) {

        "DEVELOPERPACK_E5"   {

            $disabledService1 = "EXCEL_PREMIUM"
            $disabledService2 = "RMS_S_PREMIUM2"

            $disabledServicePlans = [System.Collections.Generic.List[psobject]]::new()
            foreach ($ServicePlan in $Lic.ServicePlans){
                if (($ServicePlan.ServicePlanName -eq ($disabledService1)) -or ($ServicePlan.ServicePlanName -eq ($disabledService2)) ){
                    $disabledServicePlans.Add($ServicePlan)
                }
            }

            $permission = @{
                DisplayName    = $Lic.skuPartNumber + " Without $disabledService1 and $disabledService2"
                Identification = @{
                    Id = $Lic.Id
                    SkuId = $Lic.SkuId
                    ServicePlansToDisable = $disabledServicePlans
                }
            }
            $permissions.add( $permission )
        }
    }
}
Write-Output $permissions | ConvertTo-Json -Depth 10;


#  Definition of returned data https://docs.microsoft.com/en-us/graph/api/subscribedsku-list?view=graph-rest-1.0&tabs=http
# {
#     "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#subscribedSkus",
#     "value": [
#         {
#             "capabilityStatus": "Enabled",
#             "consumedUnits": 14,
#             "id": "48a80680-7326-48cd-9935-b556b81d3a4e_c7df2760-2c81-4ef7-b578-5b5392b571df",
#             "prepaidUnits": {
#                 "enabled": 25,
#                 "suspended": 0,
#                 "warning": 0
#             },
#             "servicePlans": [
#                 {
#                     "servicePlanId": "8c098270-9dd4-4350-9b30-ba4703f3b36b",
#                     "servicePlanName": "ADALLOM_S_O365",
#                     "provisioningStatus": "Success",
#                     "appliesTo": "User"
#                 }
#             ],
#             "skuId": "c7df2760-2c81-4ef7-b578-5b5392b571df",
#             "skuPartNumber": "ENTERPRISEPREMIUM",
#             "appliesTo": "User"
#         },
#         {
#             "capabilityStatus": "Suspended",
#             "consumedUnits": 14,
#             "id": "48a80680-7326-48cd-9935-b556b81d3a4e_d17b27af-3f49-4822-99f9-56a661538792",
#             "prepaidUnits": {
#                 "enabled": 0,
#                 "suspended": 25,
#                 "warning": 0
#             },
#             "servicePlans": [
#                 {
#                     "servicePlanId": "f9646fb2-e3b2-4309-95de-dc4833737456",
#                     "servicePlanName": "CRMSTANDARD",
#                     "provisioningStatus": "Disabled",
#                     "appliesTo": "User"
#                 }
#             ],
#             "skuId": "d17b27af-3f49-4822-99f9-56a661538792",
#             "skuPartNumber": "CRMSTANDARD",
#             "appliesTo": "User"
#         }
#     ]
# }
