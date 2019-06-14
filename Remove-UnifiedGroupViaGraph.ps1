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
#>

<#
  .SYNOPSIS
  Multi-use script to report on all Office 365 Groups in your tenant. Once you've looked over the output CSV report, if there are groups you wish to delete, you can feed the
  CSV report back into this script in Remove mode to delete the Office 365 groups by their id. Be sure the CSV that is fed back in contains ONLY the groups you wish to DELETE.

 .PARAMETER Mode
  Options: Report or Remove.   
  Report mode: A CSV report of every Office 365 group in your tenant will be generated and saved to the specified path via -CsvPath param.
  Remove mode: Specify a CSV containing the groups you wish to remove via -CsvPath param. The CSV must contain an 'id' column containing the Office 365 group ids to remove. 

 .PARAMETER CsvPath
  Path to a CSV file that will either be written to when in Report mode or read from when in Remove mode.

 .PARAMETER IncludeOwnersAndStorage
  When in Report mode, optionally include the group owners, SharePoint site URL and SharePoint storage used in the report. 
  Note this may significantly increase the time to generate the report if you have a lot of groups in your tenant.

 .EXAMPLE
  .\Remove-UnifiedGroupViaGraph.ps1 -Mode Report -CsvPath "C:\MyDocuments\Office365Groups.csv" -IncludeOwnersAndStorage

  Report all Office 365 groups in your tenant and include the group owners, site url and storage used. File will be saved to "C:\MyDocuments\Office365Groups.csv"

 .EXAMPLE
  .\Remove-UnifiedGroupViaGraph.ps1 -Mode Report -CsvPath "C:\MyDocuments\Office365Groups.csv"

  Report all Office 365 groups in your tenant and save to a file located at "C:\MyDocuments\Office365Groups.csv"

 .EXAMPLE
  .\Remove-UnifiedGroupViaGraph.ps1 -Mode Remove -CsvPath "C:\MyDocuments\Office365GroupsToDelete.csv"

  Delete all Office 365 groups by their id specified in the file located at "C:\MyDocuments\Office365GroupsToDelete.csv"
#>

param(
  [Parameter(Mandatory = $true)][ValidateSet("Report", "Remove")]$Mode,
  [Parameter(Mandatory = $true)]$CsvPath,
  [Switch]$IncludeOwnersAndStorage
)

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
    

    if ($IncludeOwnersAndStorage) 
    {
      
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


if( -not $token -or $token.ExpiresOn.DateTime -gt (Get-Date) ) 
{
  $token = Get-Token -Tenant $tenant -ClientID $clientId -ClientSecret $clientSecret -Resource "https://graph.microsoft.com"
}


if( $token.AccessToken ) 
{
  
  <# 
    MODE: REPORT 
  #>
  if ($Mode -eq "Report")
  {
    $continue = $false
    if ((Test-Path $CsvPath) -and ($overwriteReply = Read-Host -Prompt "A file already exists at $CsvPath. Do you wish to overwrite it?[y/n]")) 
    {       
      if ($overwriteReply -match "[yY]") 
      { 
        Remove-Item $CsvPath -Force
        $continue = $true
      }
    } 
    else 
    { 
      $continue = $true
    }

    if ($continue) {
      Write-Host "Fetching Office 365 Groups..."
      $unifiedGroupInfo = Get-UnifiedGroupInfo -AccessToken $token.AccessToken

      if ($IncludeOwnersAndStorage) 
      {
        $unifiedGroupInfo = $unifiedGroupInfo | SELECT Id, DisplayName, WebUrl, Mail, CreatedDateTime, SizeGB, Visibility, @{ n = "Owners"; e = { $_.Owners.userPrincipalName -Join ", " } }
      }
      else 
      {
        $unifiedGroupInfo = $unifiedGroupInfo | SELECT Id, DisplayName, Mail, CreatedDateTime, Visibility
      }

      $unifiedGroupInfo | Export-Csv -Path $CsvPath -NoTypeInformation
    }
  }
  
  <# 
    MODE: REMOVE 
  #>
  elseif ($Mode -eq "Remove")
  {
    if (-not (Test-Path $CsvPath) -or $CsvPath -notlike "*.csv") 
    {
      Write-Error "You must specify a path to a CSV file via -CsvPath param when using Remove mode."
    }
    else 
    {
      $csvRows = ConvertFrom-Csv (Get-Content $CsvPath)
      
      $groupIdsToRemove = $csvRows | SELECT -ExpandProperty id

      if ($groupIdsToRemove.Count -gt 0) 
      {        
        if ($removeReply = Read-Host -Prompt "Found $($groupIdsToRemove.Count) groups in the CSV file at $CsvPath. Are you sure you want to remove these?[y/n]") 
        {
          if ($removeReply -match "[yY]") 
          { 
            Write-Host "Removing $($groupIdsToRemove.Count) Office 365 Groups..."
            foreach ( $groupId in $groupIdsToRemove ) 
            {
              Remove-UnifiedGroupViaGraph -Id $groupId -AccessToken $token.AccessToken
            }
          }
        }      
      }
    }
  }
}