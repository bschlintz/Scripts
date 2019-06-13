# https://www.nuget.org/packages/Microsoft.IdentityModel.Clients.ActiveDirectory/
if(!(Get-Package Microsoft.IdentityModel.Clients.ActiveDirectory -RequiredVersion 3.19.8 -ErrorAction SilentlyContinue))
{
    Write-Information "Installing Dependency: Microsoft.IdentityModel.Clients.ActiveDirectory"
    Install-Package Microsoft.IdentityModel.Clients.ActiveDirectory -RequiredVersion 3.19.8 -Source 'https://www.nuget.org/api/v2' -Scope CurrentUser -Confirm:$false -Force -
}

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

function Get-UnifiedGroupInfo
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken
    )

    begin
    {
        $groups = @()
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/groups?`$filter=groupTypes/any(c:c+eq+'Unified')"
    }
    process
    {
        # get all the unified groups

            do
            {
                $json = Invoke-RestMethod -Uri $uri –Headers $headers –Method GET
                $groups += $json.value
                $uri = $json.'@odata.nextLink'
            }
            while( $uri )
    

        # get the owners for each group

            foreach( $group in $groups )
            {
                $uri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/owners"

                do
                {
                    $json = Invoke-RestMethod -Uri $uri –Headers $headers –Method GET
                    $group | Add-Member -MemberType NoteProperty -Name Owners -Value $($json.value | SELECT id, @{ name="ownerName"; expression={$_.displayName}}, mail, userprincipalName)
                    $uri = $json.'@odata.nextLink'
                }
                while( $uri )
            }

        # get the document library for each O365 group

            foreach( $group in $groups )
            {
                $json = Invoke-RestMethod -Uri "https://graph.microsoft.com/beta/groups/$($group.Id)/drive" –Headers $headers –Method GET
                $group | Add-Member -MemberType NoteProperty -Name WebUrl -Value $([System.Web.HttpUtility]::UrlDecode($json.webUrl).Replace("/Shared Documents", ""))
                $group | Add-Member -MemberType NoteProperty -Name SizeGB -Value $([Math]::Round( ($json.Quota.Used / 1GB), 2))
            }

    }
    end
    {
       $groups
    }
}

function Remove-UnifiedGroupViaGraph
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$Id,
        [Parameter(Mandatory=$true)][string]$AccessToken
    )

    begin
    {
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/groups/$Id"
    }
    process
    {
        try
        {
            $response = Invoke-RestMethod -Uri $uri –Headers $headers –Method Delete
        }
        catch
        {
            if(  $_.Exception.Response.StatusCode.value__ -eq 404 )
            {
                 Write-Error "A group with id $Id was not found."
            }
            else
            {
                Write-Error "Erorr deleting group: $Id. Error: $($_.Exception)"
            }
        }
    }
    end
    {
    }
}

# script requries Group.Read.All, User.Read.All.  Requries Group.ReadWrite.All use the delete function
$tenant       = " "
$clientId     = " "  # aka "Application ID" in Azure Portal > Azure Active Directory > App Registrations
$clientSecret = " "  # aka "Keys" in Azure Portal > Azure Active Directory > App Registrations
$csvPath      = "c:\dev\tmp\Remove-UnifiedGroupViaGraph-Sites.csv" # csv must contain at least one column named "webUrl"


if( -not $token -or $token.ExpiresOn.DateTime -gt (Get-Date) )
{
    $token = Get-Token -Tenant $tenant -ClientID $clientId -ClientSecret $clientSecret -Resource "https://graph.microsoft.com"
}


if( $token.AccessToken )
{
    $unifiedGroupInfo = Get-UnifiedGroupInfo -AccessToken $token.AccessToken | SELECT Id, DisplayName, WebUrl, Mail, CreatedDateTime, SizeGB, Visibility, @{ n="Owners"; e={$_.Owners.userPrincipalName -Join ", "}}
    $unifiedGroupInfo | Export-Csv -Path "C:\dev\tmp\UnifiedGroupInfo_$(Get-Date -Format 'yyyyMMdd').csv" -NoTypeInformation
}


<#

    # Requries Group.ReadWrite.All in the Graph API to perform a group delete

    $csvRows = ConvertFrom-Csv (Get-Content $csvPath)

    if ($null -eq $csvRows) {
        Write-Error "Unable to find Sites CSV at path $csvPath"
        break
    }

    $webUrlsToRemove = $csvRows | SELECT -ExpandProperty webUrl
    $groupsToRemove = $unifiedGroupInfo | ? { $webUrlsToRemove -contains $_.WebUrl }

    foreach( $grp in $groupsToRemove ) 
    {
        Remove-UnifiedGroupViaGraph -Id $grp.Id -AccessToken $token.AccessToken
    }
#>

