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
  Scan all site collections within the tenant to identify sites which have some level of external sharing enabled.
  Generates a CSV report of the sites which have external sharing enabled along with the total external users added to the site.

 .DESCRIPTION
  Scan all site collections within the tenant to identify sites which have some level of external sharing enabled.
  Generates a CSV report of the sites which have external sharing enabled along with the total external users added to the site.
  
  NOTE: This script requires the PowerShell module 'SharePointPnPPowerShellOnline' to be installed. If it is missing, the script will attempt to install it.

  RECOMMENDATION: Add administrator username and password for your tenant to Windows Credential Manager before running script. 
  https://github.com/SharePoint/PnP-PowerShell/wiki/How-to-use-the-Windows-Credential-Manager-to-ease-authentication-with-PnP-PowerShell

  .PARAMETER TenantRootSiteUrl
  SharePoint Tenant Root Site URL

 .EXAMPLE
  .\Report-SharePointExternalSites.ps1 -TenantRootSiteUrl "https://tenant.sharepoint.com"

  Scan all site collections within the tenant and generate a CSV report of sites with external sharing enabled.
#>

param(
    [parameter(Mandatory = $true)]$TenantRootSiteUrl
)

#############################################

$module = Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable
if ($null -eq $module) {
    Write-Output "Installing PowerShell Module: SharePointPnPPowerShellOnline"
    Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber -Confirm:$false
}

Import-Module -Name SharePointPnPPowerShellOnline -WarningAction Ignore
$module = Get-Module -Name SharePointPnPPowerShellOnline

#############################################

Connect-PnPOnline -Url $TenantRootSiteUrl

$context = Get-PnPContext
if ($null -eq $context) {
    Write-Error "Unable to successfully connect to SPO tenant with URL: $($TenantRootSiteUrl)"
    break
}

$timestamp = (Get-Date).ToString("yyyyMMdd.HHmm")
$csvName = "ExternalSites-$timestamp.csv"

Write-Host "LOADING SITES..."
$allSites = Get-PnPTenantSite

$sitesWithExternalSharing = $allSites | Where-Object { $_.SharingCapability -ne 'Disabled' }
$tenant = [Microsoft.Online.SharePoint.TenantManagement.Office365Tenant]::new($context)

foreach ($externalSite in $sitesWithExternalSharing) {
    Write-Host "SITE: $($externalSite.Url)"

    $externalUsers = $tenant.GetExternalUsersForSite($externalSite.Url, 0, 1, "", [Microsoft.Online.SharePoint.TenantManagement.SortOrder]::Descending)

    $context.Load($externalUsers)
    $context.ExecuteQuery()
		
    [PSCustomObject](@{
            SiteUrl                                  = $externalSite.Url
            SharingCapability                        = $externalSite.SharingCapability.ToString()
            ShowPeoplePickerSuggestionsForGuestUsers = $externalSite.ShowPeoplePickerSuggestionsForGuestUsers
            TotalExternalUserCount                   = $externalUsers.TotalUserCount
        }) | Select-Object -Property SiteUrl, SharingCapability, ShowPeoplePickerSuggestionsForGuestUsers, TotalExternalUserCount `
    | Export-Csv -Path $csvName -NoTypeInformation -Append
}

Write-Host "REPORT: $((Get-ChildItem $csvName).FullName)"