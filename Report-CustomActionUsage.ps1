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
  Scan all site collections within the tenant to generate a report on every site-scoped custom action.
  Generates a CSV report containing details about every custom action that was found. 

 .DESCRIPTION
  Scan all site collections within the tenant to generate a report on every site-scoped custom action.
  Generates a CSV report containing details about every custom action that was found. 

  NOTE: This script requires the PowerShell module 'SharePointPnPPowerShellOnline' to be installed. If it is missing, the script will attempt to install it.

  RECOMMENDATION: Add administrator username and password for your tenant to Windows Credential Manager before running script. 
  https://github.com/SharePoint/PnP-PowerShell/wiki/How-to-use-the-Windows-Credential-Manager-to-ease-authentication-with-PnP-PowerShell

  .PARAMETER TenantRootSiteUrl
  SharePoint Tenant Root Site URL

 .EXAMPLE
  .\Report-CustomActionUsage.ps1 -TenantRootSiteUrl "https://tenant.sharepoint.com"

  Scan all site collections within the tenant and generate a CSV report containing every site-scoped custom action.
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

#############################################

Write-Host "Connecting to Tenant Root Site URL: $TenantRootSiteUrl"
Connect-PnPOnline -Url $TenantRootSiteUrl 

$excludedTemplates = @("SPSMSITEHOST#0")        #MySite Host
$excludedTemplates += "GROUP#0"                 #Modern Group Site
$excludedTemplates += "SITEPAGEPUBLISHING#0"    #Modern Communication Site
$excludedTemplates += "STS#3"                   #Modern Team Site (not group connected)

Write-Host "Fetching list of all Site Collections"
$sites = Get-PnPTenantSite | Where-Object { $excludedTemplates -notcontains $_.Template }

Write-Host "Found $($sites.Count) Site Collections"
$timestamp = (Get-Date).ToString("yyyyMMdd.HHmm")
$csvName = "CustomActionUsage-$timestamp.csv"

foreach ($site in $sites) {
  Write-Host " SITE: $($site.Url)"
  Connect-PnPOnline -Url $site.Url
    
  $customActions = @(Get-PnPCustomAction -Scope Site)

  if ($customActions.Count -gt 0) {
    Write-Host "   Found $($customActions.Count) custom actions"
  
    foreach ($customAction in $customActions) {
      $reportRow = @{
        SiteUrl       = $site.Url
        CALocation    = $customAction.Location
        CAId          = $customAction.Id
        CAName        = $customAction.Name
        CADescription = $customAction.Description
        CAScriptBlock = $customAction.ScriptBlock
        CAScriptSrc   = $customAction.ScriptSrc
        CASequence    = $customAction.Sequence
        CAGroup       = $customAction.Group
      }
    
      ([PSCustomObject]$reportRow) | Export-Csv -Path $csvName -NoTypeInformation -Append
    }
  }
}