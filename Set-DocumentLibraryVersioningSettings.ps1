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
  Apply versioning settings to all document libraries in a target site collection, optionally including subsites.

 .DESCRIPTION
  Apply versioning settings to all document libraries in a target site collection, optionally including subsites.

  Excluded Libraries: Form Templates, Site Assets, Style Library
  
  NOTE: This script requires the PowerShell module 'SharePointPnPPowerShellOnline' to be installed. If it is missing, the script will attempt to install it.

  .PARAMETER SiteUrl
  SharePoint Site Collection URL

  .PARAMETER EnableMajorVersions
  Enable major versioning on the document library

  .PARAMETER EnableMinorVersions
  Enable minor versioning on the document library

  .PARAMETER MajorVersionsLimit
  Major versions limit

  .PARAMETER MinorVersionsLimit
  Minor versions limit

  .PARAMETER IncludeSubsites
  Apply settings to document libraries on subsites if any exist
#>

param(
    [parameter(Mandatory=$true)]$SiteUrl,
    [bool]$EnableMajorVersions = $true,
    [bool]$EnableMinorVersions = $true,
    [int]$MajorVersionsLimit = 30,
    [int]$MinorVersionsLimit = 10,
    [switch]$IncludeSubsites
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

$ExcludedLibraryTitles = @("Form Templates", "Site Assets", "Style Library")

#############################################
Function Process-Web
{
    param([parameter(Mandatory=$true)]$Web)

    begin 
    {
        #Library Settings
        $LibrarySettings = @{
            EnableVersioning = $EnableMajorVersions
            EnableMinorVersions = $EnableMinorVersions
            MajorVersions = $MajorVersionsLimit
            MinorVersions = $MinorVersionsLimit
        }
    }    
    process
    {
        Write-Host "[Site] $($Web.Url)" -ForegroundColor Cyan

        #Document Libraries that are not Excluded
        $ListIncludes = @("Title", "EnableVersioning", "EnableMinorVersions", "MajorVersionLimit", "MajorWithMinorVersionsLimit")
        $Libraries = Get-PnPList -Web $Web -Includes $ListIncludes | Where-Object { $_.BaseTemplate -eq 101 -and $ExcludedLibraryTitles -notcontains $_.Title } 

        #Update Each Library
        foreach ($Library in $Libraries) 
        {
            if ($Library.EnableVersioning -ne $EnableMajorVersions -or $Library.EnableMinorVersions -ne $EnableMinorVersions `
            -or $Library.MajorVersionLimit -ne $MajorVersionsLimit -or $Library.MajorWithMinorVersionsLimit -ne $MinorVersionsLimit) 
            {
                Write-Host "  [Updating Library] $($Library.Title)" -ForegroundColor Green
                Set-PnPList -Identity $Library -Web $Web @LibrarySettings
            }
            else 
            {
                Write-Host "  [Skipped Library - No Changes] $($Library.Title)"
            }
        }
    }
    end 
    {
    }
}

#############################################

Write-Host "Connecting to SharePoint Site Collection URL: $SiteUrl"
Connect-PnPOnline -Url $SiteUrl 

$RootWeb = Get-PnPWeb

Process-Web -Web $RootWeb

#Recursively Process Subsites if -IncludesSubsites was specified
if ($IncludeSubsites) 
{
    foreach ($SubWeb in (Get-PnPSubWebs -Web $RootWeb -Recurse)) 
    {
        Process-Web -Web $SubWeb
    }
}