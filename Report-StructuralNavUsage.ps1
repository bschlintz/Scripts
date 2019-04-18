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
  Scan all site collections and subsites within the tenant to identify usage of structural navigation.
  Generates a CSV report of the sites using Structural Navigation along with some additional information.

 .DESCRIPTION
  Scan all site collections and subsites within the tenant to identify usage of structural navigation.
  Generates a CSV report of the sites using Structural Navigation along with some additional information.
  
  NOTE: This script requires the PowerShell module 'SharePointPnPPowerShellOnline' to be installed. If it is missing, the script will attempt to install it.

  RECOMMENDATION: Add administrator username and password for your tenant to Windows Credential Manager before running script. 
  https://github.com/SharePoint/PnP-PowerShell/wiki/How-to-use-the-Windows-Credential-Manager-to-ease-authentication-with-PnP-PowerShell

  READ MORE: Navigation Options in SharePoint Online
  https://docs.microsoft.com/en-us/office365/enterprise/navigation-options-for-sharepoint-online

  .PARAMETER TenantRootSiteUrl
  SharePoint Tenant Root Site URL

  .PARAMETER ReportSitesWithStructuralNavOnly
  Only include sites or subsites which use structural navigation 

  .PARAMETER ClassicSitesOnly
  Only scan classic site templates (omit modern site templates such as Group Site, Communication Site and Modern non-group connected Team site)

 .EXAMPLE
  .\Report-StructuralNavUsage.ps1 -TenantRootSiteUrl "https://tenant.sharepoint.com"

  Scan all site collections within the tenant and generate a CSV report with every scanned site.

 .EXAMPLE
  .\Report-StructuralNavUsage.ps1 -TenantRootSiteUrl "https://tenant.sharepoint.com" -ReportSitesWithStructuralNavOnly

  Scan all site collections within the tenant and generate a CSV report with only sites that use structural navigation.

 .EXAMPLE
  .\Report-StructuralNavUsage.ps1 -TenantRootSiteUrl "https://tenant.sharepoint.com" -ClassicSitesOnly

  Scan only classic site collections within the tenant and generate a CSV report with every scanned site.

 .EXAMPLE
  .\Report-StructuralNavUsage.ps1 -TenantRootSiteUrl "https://tenant.sharepoint.com" -ReportSitesWithStructuralNavOnly -ClassicSitesOnly

  Scan only classic site collections within the tenant and generate a CSV report with only sites that use structural navigation.

#>

param(
    [parameter(Mandatory=$true)]$TenantRootSiteUrl,
    [switch]$ReportSitesWithStructuralNavOnly,
    [switch]$ClassicSitesOnly
)

#############################################

$module = Get-Module -Name SharePointPnPPowerShellOnline -ListAvailable
if ($null -eq $module) {
    Write-Output "Installing PowerShell Module: SharePointPnPPowerShellOnline"
    Install-Module -Name SharePointPnPPowerShellOnline -Force -AllowClobber -Confirm:$false
}

Import-Module -Name SharePointPnPPowerShellOnline -WarningAction Ignore
$module = Get-Module -Name SharePointPnPPowerShellOnline

#Add SharePoint Publishing CSOM Assembly for NavigationSettings object
Add-Type -Path "$(Split-Path $module.Path)\Microsoft.SharePoint.Client.Publishing.dll"

#############################################
Function New-ReportRow
{
    param($SiteUrl = "", $WebUrl = "", $SitePublishingInfrastructureActivated = $false)
    @{
        SiteUrl = $SiteUrl
        WebUrl = $WebUrl
        SitePublishingInfrastructureActivated = $SitePublishingInfrastructureActivated
        CurrentNavSource = "N/A"
        CurrentNavIsStructural = "FALSE"
        GlobalNavSource = "N/A"
        GlobalNavIsStructural = "FALSE"
        IsSubsite = $false
    }
}
Function Get-WebNavigationSettings
{
    param([parameter(Mandatory=$true)]$Web)

    begin { 
        $result = New-ReportRow -WebUrl $Web.Url
    }    
    process
    {
        try
        {           
            #Get ClientContext from Web object
            $context = $Web.Context
            $context.Load($Web)
            $context.ExecuteQuery()

            #Instantiate Navigation Settings for Web
            $navigationSettings = New-Object Microsoft.SharePoint.Client.Publishing.Navigation.WebNavigationSettings($context, $Web)

            #Load CurrentNavigation and GlobalNavigation
            $context.Load($navigationSettings.CurrentNavigation)
            $context.Load($navigationSettings.GlobalNavigation)

            #Execute Request
            $context.ExecuteQuery()

            #Add NavigationSettings to result
            $result.CurrentNavSource = $navigationSettings.CurrentNavigation.Source.ToString()
            if ($result.CurrentNavSource -eq "PortalProvider") { $result.CurrentNavIsStructural = "TRUE" }
            elseif ($result.CurrentNavSource -eq "Unknown") { $result.CurrentNavIsStructural = "UNKNOWN" }
            else { $result.CurrentNavIsStructural = "FALSE" }

            $result.GlobalNavSource = $navigationSettings.GlobalNavigation.Source.ToString()
            if ($result.GlobalNavSource -eq "PortalProvider") { $result.GlobalNavIsStructural = "TRUE" }
            elseif ($result.GlobalNavSource -eq "Unknown") { $result.GlobalNavIsStructural = "UNKNOWN" }
            else { $result.GlobalNavIsStructural = "FALSE" }
        }
        catch
        {
            Write-Error "Error getting navigation settings for web $($Web.Url). Error: $($_.Exception)"
        }
    }
    end 
    {
        $result
    }
}

#############################################

Write-Host "Connecting to Tenant Root Site URL: $TenantRootSiteUrl"
Connect-PnPOnline -Url $TenantRootSiteUrl 

$excludedTemplates = @("SPSMSITEHOST#0")            #MySite Host

if ($ClassicSitesOnly) {
    $excludedTemplates += "GROUP#0"                 #Modern Group Site
    $excludedTemplates += "SITEPAGEPUBLISHING#0"    #Modern Communication Site
    $excludedTemplates += "STS#3"                   #Modern Team Site (not group connected)
}

Write-Host "Fetching list of all Site Collections"
$sites = Get-PnPTenantSite | Where-Object {$excludedTemplates -notcontains $_.Template}

Write-Host "Found $($sites.Count) Site Collections"
$timestamp = (Get-Date).ToString("yyyyMMdd.HHmm")
$csvName = "StructuralNavUsage-$timestamp.csv"

foreach ($site in $sites) {
    Write-Host " Current: $($site.Url)"
    Connect-PnPOnline -Url $site.Url
    $siteResults = @()

    #check if publishing infrastructure site collection feature is enabled
    $pubInfraFtr = Get-PnPFeature -Scope Site -Identity "f6924d36-2fa8-4f0b-b16d-06b7250180fa" #PublishingSite
    $pubInfraActivated = ($null -ne $pubInfraFtr) -and (([Array]$pubInfraFtr).Count -gt 0)

    #if publishing infrastrucutre not activated, log and skip nav checks
    if (!$pubInfraActivated) {
        if (!$ReportSitesWithStructuralNavOnly) {
            $result = New-ReportRow -SiteUrl $site.Url -WebUrl $site.Url -SitePublishingInfrastructureActivated $false
            ([PSCustomObject]$result) | Export-Csv -Path $csvName -NoTypeInformation -Append
        }
    } 
    else {
        #check root web
        $rootWeb = Get-PnPWeb 
        $siteResults += Get-WebNavigationSettings -Web $rootWeb
    
        #check subsites
        $subsites = Get-PnPSubWebs -Recurse 
        foreach ($subsite in $subsites) {
            $subsiteResult = Get-WebNavigationSettings -Web $subsite
            $subsiteResult.IsSubsite = $true
            $siteResults += $subsiteResult
        }    
    
        #build report
        foreach ($webResult in $siteResults) {
            if (!$ReportSitesWithStructuralNavOnly -or $webResult.CurrentNavIsStructural -ne "FALSE" -or $webResult.GlobalNavIsStructural -ne "FALSE") {
                $webResult.SiteUrl = $site.Url
                $webResult.SitePublishingInfrastructureActivated = $true
                ([PSCustomObject]$webResult) | Export-Csv -Path $csvName -NoTypeInformation -Append
            }
        }
    }
}