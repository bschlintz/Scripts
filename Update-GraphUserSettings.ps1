<# 
 Microsoft provides programming examples for illustration only, without warranty either expressed or
 implied, including, but not limited to, the implied warranties of merchantability and/or fitness 
 for a particular purpose. We grant You a nonexclusive, royalty-free right to use and modify the 
 Sample Code and to reproduce and distribute the object code form of the Sample Code, provided that
 You agree: (i) to not use Our name, logo, or trademarks to market Your software product in which the
 Sample Code is embedded; (ii) to include a valid copyright notice on Your software product in which
 the Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us and Our suppliers
 from and against any claims or lawsuits, including attorneys' fees, that arise or result from the
 use or distribution of the Sample Code.
 
 This sample assumes that you are familiar with the programming language being demonstrated and the 
 tools used to create and debug procedures. Microsoft support professionals can help explain the 
 functionality of a particular procedure, but they will not modify these examples to provide added 
 functionality or construct procedures to meet your specific needs. if you have limited programming 
 experience, you may want to contact a Microsoft Certified Partner or the Microsoft fee-based consulting 
 line at (800) 936-5200. 
#>

<#
  .SYNOPSIS
  Script to report and update on the Delve ContributionToContentDiscoveryDisabled setting for all or a subset of users in your organization.

 .DESCRIPTION
  Script to report and update on the Delve ContributionToContentDiscoveryDisabled setting for all or a subset of users in your organization.

 .PARAMETER ReportOnly
  Generate a CSV report of the users that would be updated with the specified ContributionToContentDiscoveryDisabled parameter setting. No updates will be made when using this switch.

 .PARAMETER ContributionToContentDiscoveryDisabled
  When set to $true, the delegate access to the user's trending API is disabled. 
  When set to $true, documents in the user's Office Delve are disabled. 
  When set to $true, the relevancy of the content displayed in Office 365, for example in Suggested sites in SharePoint Home and the Discover view in OneDrive for Business is affected. 
  Users can control this setting in Office Delve.

 .PARAMETER AllUsers
  Target all users in the organization excluding guests.

 .PARAMETER TargetUsersCsvPath
  Provide a CSV file path containing the target users. The must have at least one column called userPrincipalName for the script to use it.

 .PARAMETER TargetUsersAADGroupName
  Provide an Azure AD group display name containing the target users. The group may be a security group, distribution group of Office 365 group. 
  The app registration must have Group.Read.All permission to use this parameter.

 .EXAMPLE
  .\Update-GraphUserSettings.ps1 -ContributionToContentDiscoveryDisabled $true -ReportOnly -AllUsers
 
  Generate a CSV report of all users which currently have the ContributionToContentDiscoveryDisabled set to FALSE.

 .EXAMPLE
  .\Update-GraphUserSettings.ps1 -ContributionToContentDiscoveryDisabled $true -AllUsers

  Update all users to set ContributionToContentDiscoveryDisabled to TRUE.

 .EXAMPLE
  .\Update-GraphUserSettings.ps1 -ContributionToContentDiscoveryDisabled $true -TargetUsersCsvPath .\Update-GraphUserSettings-TargetUsers.csv

  Update the users listed in the CSV file to set ContributionToContentDiscoveryDisabled to TRUE.
 
 .EXAMPLE
  .\Update-GraphUserSettings.ps1 -ContributionToContentDiscoveryDisabled $true -TargetUsersAADGroupName "Germany Employees"

  Update the users in the 'Germany Employees' distribution group to set ContributionToContentDiscoveryDisabled to TRUE.
#>

# Script Parameters
param 
(
    [Parameter(Mandatory=$false)][switch]$ReportOnly,
    [Parameter(Mandatory=$true)][bool]$ContributionToContentDiscoveryDisabled,
    [Parameter(Mandatory=$false,ParameterSetName="AllUsers")][switch]$AllUsers,
    [Parameter(Mandatory=$false,ParameterSetName="UsersFromCsv")][string]$TargetUsersCsvPath,
    [Parameter(Mandatory=$false,ParameterSetName="UsersFromGroup")][string]$TargetUsersAADGroupName
)

function Get-Token 
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $true)][string]$Tenant,
        [Parameter(Mandatory = $true)][System.Guid]$ClientID,
        [Parameter(Mandatory = $true)][string]$ClientSecret,
        [Parameter(Mandatory = $true)][string]$Resource
    )
    
    begin 
    {
    }  
    process {
        $Body = @{
            grant_type    = "client_credentials"
            scope         = "$Resource/.default"
            client_id     = $ClientID
            client_secret = $ClientSecret
        } 
      
        $RequestToken = @{
            ContentType = 'application/x-www-form-urlencoded'
            Method      = 'POST'
            Body        = $Body
            Uri         = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0/token"
        }
  
        try {
            $response = Invoke-RestMethod @RequestToken
        }
        catch {
            Write-Error $_.Exception
        }
    }
    end 
    {
        $response
    }
}

function Get-AuthenticationHeaders
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken
    )

    begin
    {
    }
    process
    {
        @{
            'Content-Type'  = 'application/json'
            'Authorization' = "Bearer $($AccessToken)"
        }    
    }
    end
    {
    }
}

function Split-ArrayIntoChunks
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][object[]]$Array,
        [Parameter(Mandatory=$true)][int]$ChunkSize
    )

    begin
    {
        $counter = [pscustomobject] @{ Value = 0 }
    }
    process
    {
        $Array | Group-Object -Property { [math]::Floor($counter.Value++ / $ChunkSize) }
    }
    end
    {
    }
}

function Get-AllUsers
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken
    )

    begin
    {
        $users = @()
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=userPrincipalName&`$filter=userType eq 'Member'" #Exclude Guest Accounts
    }
    process
    {
        # get all the users

            do
            {
                $json = Invoke-RestMethod -Uri $uri -Headers $headers -Method GET
                $users += $json.value
                $uri = $json.'@odata.nextLink'
            }
            while( $uri )
    }
    end
    {
       $users
    }
}

function Get-AllUsersInGroup
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$GroupDisplayName
    )

    begin
    {
        $users = @()
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $getGroupIdUri = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$GroupDisplayName'&`$select=id"
        $getGroupMembersUpnUri = "https://graph.microsoft.com/v1.0/groups/{0}/members?`$select=userPrincipalName,userType"
    }
    process
    {
        # get group id

            $groupJson = Invoke-RestMethod -Uri $getGroupIdUri -Headers $headers -Method GET

        # validate we have one matching group

            $groupId = $null
            if ($groupJson -and $groupJson.value -and $groupJson.value.Length -gt 0)
            {
                if ($groupJson.value.length -eq 1) {
                    $groupId = $groupJson.value[0].id
                }
                else {
                    Write-Error "Found more than one group with display name '$GroupDisplayName'"
                }
            }
            else {
                Write-Error "Insufficient permissions or no group exists with display name '$GroupDisplayName'"
            }
            
        
        # get all users in the group

            if ($null -ne $groupId)
            {
                do
                {
                    $getGroupMembersUpnUri = $getGroupMembersUpnUri -f $groupId
                    $groupUsersJson = Invoke-RestMethod -Uri $getGroupMembersUpnUri -Headers $headers -Method GET
                    $users += $groupUsersJson.value
                    $getGroupMembersUpnUri = $groupUsersJson.'@odata.nextLink'
                }
                while( $uri )
            }

        # remove users that are guests

            $users = $users | Where-Object {$_.userType -eq 'Member'}
    }
    end
    {
       $users
    }
}

function CreateBatch-UpdateUserSettings 
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string[]]$UserPrincipalNames,
        [Parameter(Mandatory=$true)][bool]$ContributionToContentDiscoveryDisabled
    )

    begin
    {
    }
    process
    {
        @{
            "requests" = @(foreach($userPrincipalName in $UserPrincipalNames) { 
                @{
                    "id" = $userPrincipalName
                    "url" = "/users/$userPrincipalName/settings"
                    "method" = "PATCH"
                    "headers" = @{
                        "Content-Type" = "application/json"
                    }
                    "body" = @{
                        "contributionToContentDiscoveryDisabled" = $ContributionToContentDiscoveryDisabled
                    }
                }
            })
        } | ConvertTo-Json -Depth 3 -Compress
    }
    end
    {
    }
}

function CreateBatch-GetUserSettings 
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string[]]$UserPrincipalNames
    )

    begin
    {
    }
    process
    {
        @{
            "requests" = @(foreach($userPrincipalName in $UserPrincipalNames) { 
                @{
                    "id" = $userPrincipalName
                    "url" = "/users/$userPrincipalName/settings"
                    "method" = "GET"
                    "headers" = @{
                        "Content-Type" = "application/json"
                    }
                }
            })
        } | ConvertTo-Json -Depth 3 -Compress
    }
    end
    {
    }
}

function Process-GraphBatch
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$Payload,
        [Parameter(Mandatory=$true)][string]$AccessToken
    )

    begin
    {
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/`$batch"
    }
    process
    {
        try
        {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Body $Payload -Method Post

            foreach ($batchResponse in $response.responses) 
            {
                if ($batchResponse.status -ge 400)
                {
                    Write-Error "Error in batch request: $($batchResponse.id). Error: $(ConvertTo-Json $batchResponse.body.error -Depth 5)"
                }
            }             
        }
        catch
        {
            Write-Error "Error sending batch request. Error: $($_.Exception)"
        }
    }
    end
    {
        $response
    }
}


# script requires User.ReadWrite.All 
# script requires Group.Read.All if using the -TargetUsersAADGroupName parameter (optional)

$tenant       = "contoso.onmicrosoft.com"
$clientId     = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxx"                # aka "Application ID" in Azure Portal > Azure Active Directory > App Registrations
$clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"     # aka "Keys" in Azure Portal > Azure Active Directory > App Registrations

# Graph $batch limit is 20
# https://developer.microsoft.com/en-us/graph/docs/concepts/known_issues#json-batching
$batchSize = 20

$token = $null
$token = Get-Token -Tenant $tenant -ClientID $clientId -ClientSecret $clientSecret -Resource "https://graph.microsoft.com"

if( -not $token -or -not $token.access_token ) 
{
    break
}

$targetUsers = @()
switch ($PSCmdlet.ParameterSetName)
{
    "AllUsers" {
        Write-Host "Fetching All Users from Microsoft Graph..."
        $targetUsers = Get-AllUsers -AccessToken $token.access_token | Select-Object -ExpandProperty userPrincipalName
        break
    }

    "UsersFromCsv" {
        Write-Host "Fetching Target Users from Csv..."
        if (Test-Path $TargetUsersCsvPath) {
            $userCsvRows = ConvertFrom-Csv (Get-Content $TargetUsersCsvPath)
            foreach ($user in $userCsvRows) {
                $targetUsers += $user.userPrincipalName
            }
        }
        break
    }

    "UsersFromGroup" {
        Write-Host "Fetching Target Users in AAD Group from Microsoft Graph..."
        $targetUsers = Get-AllUsersInGroup -GroupDisplayName $TargetUsersAADGroupName -AccessToken $token.access_token | Select-Object -ExpandProperty userPrincipalName
        break
    }
}

if ( $targetUsers.Length -eq 0 )
{
    break
}

Write-Host "Splitting $($targetUsers.Count) users into $([Math]::Ceiling($targetUsers.Count / $batchSize)) batches..."
$batches = Split-ArrayIntoChunks -Array $targetUsers -ChunkSize $batchSize

if ($ReportOnly)
{
    # Remove existing CSV if there is one on -ReportOnly mode
    Remove-Item -Path "DelveUserSettingsReport.csv" -Force -ErrorAction SilentlyContinue
}

# Iterate each batch of users to create Graph $batch JSON payload
for ($idx = 0; $idx -lt $batches.Length; $idx++) 
{
    $batch = $batches[$idx].Group

    if ($ReportOnly)
    {
        Write-Host "Processing Batch: $($idx + 1) of $($batches.Length) [Batch Size: $($batch.Count)]`t-ReportOnly"

        $batchPayload = CreateBatch-GetUserSettings -UserPrincipalNames $batch

        $result = Process-GraphBatch -Payload $batchPayload -AccessToken $token.access_token

        foreach ($batchResponse in $result.responses) 
        {
            $userPrincipalName = $batchResponse.id
            $currentSetting = $batchResponse.body.contributionToContentDiscoveryDisabled

            if ($null -ne $currentSetting -and $currentSetting -ne $ContributionToContentDiscoveryDisabled)
            {
                [PSCustomObject] @{
                    UserPrincipalName = $userPrincipalName
                    ContributionToContentDiscoveryDisabled = $currentSetting
                } | Export-Csv -Path "DelveUserSettingsReport.csv" -NoTypeInformation -Append
            }
        }
    }
    else 
    {
        Write-Host "Processing Batch: $($idx + 1) of $($batches.Length) [Batch Size: $($batch.Count)]"

        $batchPayload = CreateBatch-UpdateUserSettings -UserPrincipalNames $batch -ContributionToContentDiscoveryDisabled $ContributionToContentDiscoveryDisabled

        Process-GraphBatch -Payload $batchPayload -AccessToken $token.access_token | Out-Null
    }
}
