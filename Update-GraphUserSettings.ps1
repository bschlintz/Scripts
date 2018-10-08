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

# https://www.nuget.org/packages/Microsoft.IdentityModel.Clients.ActiveDirectory/ 
Add-Type -Path "C:\Packages\microsoft.identitymodel.clients.activedirectory.3.19.8\lib\net45\Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
Add-Type -Path "C:\Packages\microsoft.identitymodel.clients.activedirectory.3.19.8\lib\net45\Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

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
        $uri = "https://graph.microsoft.com/v1.0/users?`$select=userPrincipalName"
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

function Update-UserSettings
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$UserPrincipalName,
        [Parameter(Mandatory=$true)][bool]$ContributionToContentDiscoveryDisabled,
        [Parameter(Mandatory=$true)][string]$AccessToken
    )

    begin
    {
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/users/$UserPrincipalName/settings"
        $body = @{
            "contributionToContentDiscoveryDisabled" = $ContributionToContentDiscoveryDisabled
        }
    }
    process
    {
        try
        {
            $response = Invoke-RestMethod -Uri $uri -Headers $headers -Body (ConvertTo-Json $body) -Method Patch
        }
        catch
        {
            if(  $_.Exception.Response.StatusCode.value__ -eq 404 )
            {
                Write-Error "A user with upn $UserPrincipalName was not found."
            }
            else
            {
                Write-Error "Error updating user settings: $UserPrincipalName. Error: $($_.Exception)"
            }
        }
    }
    end
    {
    }
}

function Create-UserSettingsBatchPayload 
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
            "requests" = @(1..$batch.Count | ForEach-Object { 
                @{
                    "id" = $_
                    "method" = "PATCH"
                    "url" = "/users/$($UserPrincipalNames[$_-1])/settings"
                    "body" = @{
                        "contributionToContentDiscoveryDisabled" = $ContributionToContentDiscoveryDisabled
                    }
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
                    Write-Error "Error in Batch $($batchResponse.id). Error: $(ConvertTo-Json $batchResponse.body.error -Depth 5)"
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
    }
}


# script requries User.ReadWrite.All

$tenant       = "contoso.onmicrosoft.com"
$clientId     = "xxxxxxx-xxxx-xxxx-xxxx-xxxxxxx"                # aka "Application ID" in Azure Portal > Azure Active Directory > App Registrations
$clientSecret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"     # aka "Keys" in Azure Portal > Azure Active Directory > App Registrations

$contributionToContentDiscoveryDisabled = $false

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

    # Iterate each batch of users  to create Graph $batch JSON payload
    for ($idx = 0; $idx -lt $batches.Length; $idx++) 
    {
        $batch = $batches[$idx].Group
        Write-Host "Processing Batch: $($idx + 1) of $($batches.Length) [Batch Size: $($batch.Count)]"

        $batchPayload = Create-UserSettingsBatchPayload -UserPrincipalNames $batch -ContributionToContentDiscoveryDisabled $ContributionToContentDiscoveryDisabled

        Process-GraphBatch -Payload $batchPayload -AccessToken $token.AccessToken
    }

    # foreach ($user in $allUsers)
    # {
    #     Write-Host "Updating user:" $user.userPrincipalName
	#     Update-UserSettings -UserPrincipalName $user.userPrincipalName -ContributionToContentDiscoveryDisabled $contributionToContentDiscoveryDisabled -AccessToken $token.AccessToken
    # }    
}