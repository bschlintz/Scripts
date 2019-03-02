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
  Script to report and update on the Delve ContributionToContentDiscoveryDisabled setting for all tenant users (excluding guest accounts).

 .DESCRIPTION
  Script to report and update on the Delve ContributionToContentDiscoveryDisabled setting.

 .PARAMETER ReportOnly
  Generate a CSV report of all users that would be updated with the specified ContributionToContentDiscoveryDisabled parameter setting. No updates will be made when using this switch.

 .PARAMETER ContributionToContentDiscoveryDisabled
  When set to $true, the delegate access to the user's trending API is disabled. 
  When set to $true, documents in the user's Office Delve are disabled. 
  When set to $true, the relevancy of the content displayed in Office 365, for example in Suggested sites in SharePoint Home and the Discover view in OneDrive for Business is affected. 
  Users can control this setting in Office Delve.

 .EXAMPLE
  .\Update-GraphUserSettings.ps1 -ContributionToContentDiscoveryDisabled $true -ReportOnly
 
  Generate a CSV report of all users which currently have the ContributionToContentDiscoveryDisabled set to FALSE.

 .EXAMPLE
  .\Update-GraphUserSettings.ps1 -ContributionToContentDiscoveryDisabled $true

  Update all users to set ContributionToContentDiscoveryDisabled to TRUE.
#>

# Script Parameters
param 
(
    [Parameter(Mandatory=$false)][switch]$ReportOnly,
    [Parameter(Mandatory=$true)][bool]$ContributionToContentDiscoveryDisabled
)

if(!(get-package Microsoft.IdentityModel.Clients.ActiveDirectory -RequiredVersion 3.19.8))
{
Install-Package Microsoft.IdentityModel.Clients.ActiveDirectory -RequiredVersion 3.19.8 -Source 'https://www.nuget.org/api/v2' -Scope CurrentUser
}

# https://www.nuget.org/p ackages/Microsoft.IdentityModel.Clients.ActiveDirectory/ 
$package = Get-Package Microsoft.IdentityModel.Clients.ActiveDirectory -RequiredVersion 3.19.8
$packagePath = Split-Path $package.Source -Parent
$dllPath = Join-Path -Path $packagePath -ChildPath "lib/net45/Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
Add-Type -Path $dllPath -ErrorAction Stop

function Get-Token
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$Tenant,
        [Parameter(Mandatory=$true)][System.Guid]$ClientID,
        [Parameter(Mandatory=$true)][string]$ClientSecret,
        [Parameter(Mandatory=$true)][string]$Resource
    )

    begin
    {
        $redirectUri = New-Object System.Uri("urn:ietf:wg:oauth:2.0:oob")
        $authority   = "https://login.microsoftonline.com/$Tenant"
    }
    process
    {
        try
        {
            $clientCredential = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.ClientCredential( $ClientID, $ClientSecret )
            $authContext      = New-Object Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext( $authority )
            $authContext.AcquireTokenAsync( $Resource, $clientcredential ).Result
        }
        catch
        {
            Write-Error $_.Exception
        }
    }
    end
    {
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


# script requries User.ReadWrite.All

$tenant       = "contoso.onmicrosoft.com"
$clientId     = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxx"                # aka "Application ID" in Azure Portal > Azure Active Directory > App Registrations
$clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"     # aka "Keys" in Azure Portal > Azure Active Directory > App Registrations

# Graph $batch limit is 20
# https://developer.microsoft.com/en-us/graph/docs/concepts/known_issues#json-batching
$batchSize = 20

if( -not $token -or $token.ExpiresOn.DateTime -lt (Get-Date) )
{
    $token = Get-Token -Tenant $tenant -ClientID $clientId -ClientSecret $clientSecret -Resource "https://graph.microsoft.com"
}

if( $token.AccessToken )
{
    Write-Host "Fetching Users..."
    $allUsers = Get-AllUsers -AccessToken $token.AccessToken | Select-Object -ExpandProperty userPrincipalName

    Write-Host "Splitting $($allUsers.Count) users into $([Math]::Ceiling($allUsers.Count / $batchSize)) batches..."
    $batches = Split-ArrayIntoChunks -Array $allUsers -ChunkSize $batchSize

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

            $result = Process-GraphBatch -Payload $batchPayload -AccessToken $token.AccessToken

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
    
            Process-GraphBatch -Payload $batchPayload -AccessToken $token.AccessToken | Out-Null
        }
    }
}
