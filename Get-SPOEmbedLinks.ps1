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
  Evaluate the embed URL for a CSV of file URLs stored within SPO

  .DESCRIPTION
  Evaluate the embed URL for a CSV of file URLs stored within SPO
  Uses the SharePoint Search API to discover the embed URL for documents
  Ensure the calling user has read access to all documents contained in the input CSV
  Produces a new CSV in the script directory called SPOEmbedLinks-<date>.csv containing the original document URL and document embed URL

  NOTE: This script requires the PowerShell module 'SharePointPnPPowerShellOnline' to be installed. If it is missing, the script will attempt to install it.

  RECOMMENDATION: Add administrator username and password for your tenant to Windows Credential Manager before running script. 
  https://github.com/SharePoint/PnP-PowerShell/wiki/How-to-use-the-Windows-Credential-Manager-to-ease-authentication-with-PnP-PowerShell

  .PARAMETER SiteUrl
  SharePoint Site URL (can be any that the calling user has access to)

  .PARAMETER CSVPath
  Specify the path to a CSV file containing a DocumentUrl field

  .PARAMETER UseWebLogin
  Login via web browser

  .EXAMPLE
  .\Get-SPOEmbedLinks.ps1 -SiteUrl "https://tenant.sharepoint.com"

  Generated the embed URL for every document URL specified in a CSV file in the script directory called Get-SPOEmbedLinks-Input.csv

  .EXAMPLE
  .\Get-SPOEmbedLinks.ps1 -SiteUrl "https://tenant.sharepoint.com" -CSVPath C:\temp\documentLinks.csv

  Generated the embed URL for every document URL specified in a CSV file located at C:\temp\documentLinks.csv

  .EXAMPLE
  .\Get-SPOEmbedLinks.ps1 -SiteUrl "https://tenant.sharepoint.com" -CSVPath C:\temp\documentLinks.csv -UseWebLogin

  Generated the embed URL for every document URL specified in a CSV file located at C:\temp\documentLinks.csv, using the browser for login
#>

param(
  [parameter(Mandatory = $true)]$SiteUrl,
  $CSVPath = "$(Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)\Get-SPOEmbedLinks-Sample.csv",
  [switch]$UseWebLogin
)

#############################################

$module = Get-Module SharePointPnPPowerShellOnline -ListAvailable
if ($null -eq $module) {
    Write-Host "Installing PowerShell Module: SharePointPnPPowerShellOnline"
    Install-Module SharePointPnPPowerShellOnline -Force -AllowClobber -Confirm:$false
    Write-Host
}

#############################################

function Split-ArrayIntoChunks {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][object[]]$Array,
    [Parameter(Mandatory = $true)][int]$ChunkSize
  )

  begin {
    $counter = [pscustomobject] @{ Value = 0 }
  }
  process {
    $Array | Group-Object -Property { [math]::Floor($counter.Value++ / $ChunkSize) }
  }
  end {
  }
}

function Get-SPODocEmbedUrls {
  [CmdletBinding()]
  param
  (
    [Parameter(Mandatory = $true)][object[]]$DocLinks
  )
  
  begin { }
  process {
    # Build Search Query to find all Doc Paths
    $searchQuery = "path:$(@($DocLinks | ForEach-Object { $_.CleanUrl }) -join " OR path:")"

    # Send Batch to Search
    $searchResults = Invoke-PnPSearchQuery -Query $searchQuery -MaxResults $DocLinks.Count -SelectProperties "ServerRedirectedEmbedURL", "Path" -ClientType "ContentSearchRegular" -EnableQueryRules:$false
  }
  end { 
    # Turn Search results into an Array of Objects with a Path and Embed property
    @($searchResults.ResultRows | ForEach-Object { 
      return @{ 
        Path = $_["Path"]
        EmbedUrl = $_["ServerRedirectedEmbedURL"] 
      } 
    })
  }
}

#############################################

# Runtime Variables
$TIMESTAMP = (Get-Date).ToString("yyyyMMdd.HHmm")
$EXPORT_CSV_PATH = "$($(Split-Path -Parent -Path $MyInvocation.MyCommand.Definition))\SPOEmbedLinks-$TIMESTAMP.csv"
$BATCH_SIZE = 5
$docLinks = @()

#############################################

# Connect to SPO Site
Write-Host "Connecting to SPO: $SiteUrl"
if ($UseWebLogin) {
  Connect-PnPOnline -Url $SiteUrl -UseWebLogin
}
else {
  Connect-PnPOnline -Url $SiteUrl
}

# Read Input CSV of Document Urls
$csvRows = ConvertFrom-Csv (Get-Content $CSVPath)

if ($null -eq $csvRows) {
  Write-Error "Unable to find CSV at path $CSVPath"
  break
}

# Parse Every Row in Input CSV
foreach ($row in $csvRows)
{
  $docUrl = [Uri]::new($row.DocumentUrl)
  
  # Clean Path: Remove "/:b:/r/", ?csf=1&web=1&e=VetUKT, etc.
  $cleanPath = $docUrl.AbsolutePath -replace "(\/\:\w\:\/\w)", ""
  
  # Build New Absolute URL
  $docUrlClean = "$($docUrl.Scheme)://$($docUrl.Host)$($cleanPath)"

  $docLinks += @{
    OriginalUrl = $docUrl.OriginalString
    CleanUrl = $docUrlClean
    EmbedUrl = ""
  }
}

# Split Document URLs into Batches of $BATCH_SIZE
Write-Host "Splitting $($docLinks.Count) doc links into $([Math]::Ceiling($docLinks.Count / $BATCH_SIZE)) batches..."
$batches = Split-ArrayIntoChunks -Array $docLinks -ChunkSize $BATCH_SIZE

# Iterate Every Batch of Document URLS, send to Search Service to get Embed URLs
for ($idx = 0; $idx -lt $batches.Length; $idx++) {
  $batch = $batches[$idx].Group

  Write-Host "Processing Batch: $($idx + 1) of $($batches.Length) [Batch Size: $($batch.Count)]"

  # Send Batch of Document URLs to Search 
  $docEmbedResults = Get-SPODocEmbedUrls -DocLinks $batch

  # Re-Join the Embed URL with the Originating Document URLs
  foreach ($embedResult in $docEmbedResults) {
    $docLinkMatches = @($docLinks | Where-Object { $_.CleanUrl.ToLower() -eq [System.Web.HTTPUtility]::UrlPathEncode($embedResult.Path).ToLower() })

    foreach ($docLinkMatch in $docLinkMatches) {
      $docLinkMatch.EmbedUrl = $embedResult.EmbedUrl
    }
  }
}

# Export Original Document Url and Embed URL to new CSV
$docLinks | ForEach-Object {
  ([PSCustomObject]@{
    DocumentUrl = $_.OriginalUrl
    EmbedUrl    = $_.EmbedUrl
  }) | Export-Csv -Path $EXPORT_CSV_PATH -NoTypeInformation -Append
}