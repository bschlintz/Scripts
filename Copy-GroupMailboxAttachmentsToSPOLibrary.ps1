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
  Script will copy all attachments from all mail messages in an Office 365 Group's mailbox to a SharePoint document library. 
  Supports one group by specifying individual parameters or multiple groups by CSV.

 .DESCRIPTION
  Script will copy all attachments from all mail messages in an Office 365 Group's mailbox to a SharePoint document library.
  Supports one group by specifying individual parameters or multiple groups by CSV.  
  
  NOTE: This script requires the PowerShell module 'SharePointPnPPowerShellOnline' to be installed. http://aka.ms/pnppowershell

  .PARAMETER GroupId
  Office 365 Group ID. Required for processing a single Group. 

  .PARAMETER SiteUrl
  Optional. SharePoint Site URL where the mail attachments will be uploaded to. Defaults to Office 365 group's SharePoint site.

  .PARAMETER LibraryName
  Optional URL of the SharePoint Document Library where the attachments will be uploaded to. Defaults to 'Shared Documents'

  .PARAMETER FolderName
  Optional folder name where the attachments should be uploaded to within target document library

  .PARAMETER CSVPath
  CSV file to process multiple Office 365 Groups. 
  CSV should have columns 'id', 'siteUrl', 'libraryName' and 'folderName'. Only the 'id' column is required.
  
  If siteUrl is not specified, default will be the Office 365 Group's SharePoint site.
  If libraryName is not specified, default will be 'Shared Documents' (present on every Office 365 Group SharePoint site by default).
  If folderName is not specified, attachments will be uploaded to folders at the root of the library by email subject.

 .EXAMPLE
  .\Copy-GroupMailboxAttachmentsToSPOLibrary.ps1 -GroupId "fc8f7128-587a-4b22-bf3a-7d142dd4fd32" -SiteUrl "https://contoso.sharepoint.com/sites/abc" -LibraryName "Attachments"
 
  Uploads all file attachments from all mail in the specified Office 365 Group's mailbox. 
  Attachments uploaded to a SharePoint site named 'abc' to a library named 'Attachments'.
  A sub-folder with the subject name and timestamp will be created where the attachments will be uploaded into.

 .EXAMPLE
  .\Copy-GroupMailboxAttachmentsToSPOLibrary.ps1 -GroupId "fc8f7128-587a-4b22-bf3a-7d142dd4fd32" -SiteUrl "https://contoso.sharepoint.com/sites/abc" -LibraryName "Attachments" -FolderName "GroupXYZ"
 
  Uploads all file attachments from all mail in the specified Office 365 Group's mailbox. 
  Attachments uploaded to a SharePoint site named 'abc' to a library named 'Attachments' within a sub-folder called 'GroupXYZ'
  Another sub-folder with the subject name and timestamp will be created where the attachments will be uploaded into.

  .EXAMPLE
   .\Copy-GroupMailboxAttachmentsToSPOLibrary.ps1 -CSVPath .\Copy-GroupMailboxAttachmentsToSPOLibrary-TargetGroups.csv

   Processes all rows in the CSV file describing an Office 365 Group's ID and additional options for the location of the processed attachments.
   Uploads all file attachments from all mail in the specified Office 365 Group's mailbox.    
#>

# Script Parameters
param 
(
    [Parameter(Mandatory=$true,ParameterSetName="OneGroup")][string]$GroupId,
    [Parameter(Mandatory=$false,ParameterSetName="OneGroup")][string]$SiteUrl,
    [Parameter(Mandatory=$false,ParameterSetName="OneGroup")][string]$LibraryName = "Shared Documents",
    [Parameter(Mandatory=$false,ParameterSetName="OneGroup")][string]$FolderName = "",
    [Parameter(Mandatory=$true,ParameterSetName="MultipleGroups")][string]$CSVPath
)


# Adapted From: https://ridicurious.com/2019/02/01/retry-command-in-powershell/
function Invoke-RestMethodWithRetry 
{
    [CmdletBinding()]
    param (
        [parameter(Mandatory, ValueFromPipeline)] 
        [ValidateNotNullOrEmpty()]
        [hashtable] $Params,
        [int] $RetryCount = 5,
        [int] $TimeoutInSecs = 10
    )
        
    process {
        $Attempt = 1
        $Flag = $true
        
        do {
            try {
                $PreviousPreference = $ErrorActionPreference
                $ErrorActionPreference = 'Stop'
                Invoke-RestMethod @Params -OutVariable Result
                $ErrorActionPreference = $PreviousPreference

                # flow control will execute the next line only if the command in the scriptblock executed without any errors
                # if an error is thrown, flow control will go to the 'catch' block
                $Flag = $false
            }
            catch {
                if ($Attempt -gt $RetryCount) {
                    Write-Host "Failed to execute Invoke-RestMethod. Total retry attempts: $RetryCount"
                    Write-Host "  [Error] $($_.exception.message) `n"
                    $Flag = $false
                }
                else {
                    Write-Host "[$Attempt/$RetryCount] Failed to execute Invoke-RestMethod. Retrying in $TimeoutInSecs seconds..."
                    Write-Host "  [Error] $($_.exception.message) `n"
                    Start-Sleep -Seconds $TimeoutInSecs
                    $Attempt = $Attempt + 1
                }
            }
        }
        While ($Flag)
        
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

function Get-ConversationsWithAttachments
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$GroupId
    )

    begin
    {
        $conversations = @()
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/conversations?`$filter=hasAttachments eq true" 
    }
    process
    {
        do
        {
            $json = Invoke-RestMethodWithRetry @{ Uri = $uri; Headers = $headers; Method = "GET" }
            $conversations += $json.value
            $uri = $json.'@odata.nextLink'
        }
        while( $uri )
    }
    end
    {
       $conversations
    }
}

function Get-Threads
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$GroupId,
        [Parameter(Mandatory=$true)][string]$ConversationId
    )

    begin
    {
        $threads = @()
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/conversations/$ConversationId/threads" 
    }
    process
    {
        do
        {
            $json = Invoke-RestMethodWithRetry @{ Uri = $uri; Headers = $headers; Method = "GET" }
            $threads += $json.value
            $uri = $json.'@odata.nextLink'
        }
        while( $uri )
    }
    end
    {
       $threads
    }
}

function Get-Posts
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$GroupId,
        [Parameter(Mandatory=$true)][string]$ConversationId,
        [Parameter(Mandatory=$true)][string]$ThreadId
    )

    begin
    {
        $posts = @()
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/conversations/$ConversationId/threads/$ThreadId/posts" 
    }
    process
    {
        do
        {
            $json = Invoke-RestMethodWithRetry @{ Uri = $uri; Headers = $headers; Method = "GET" }
            $posts += $json.value
            $uri = $json.'@odata.nextLink'
        }
        while( $uri )
    }
    end
    {
       $posts
    }
}

function Get-Attachments
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)][string]$AccessToken,
        [Parameter(Mandatory=$true)][string]$GroupId,
        [Parameter(Mandatory=$true)][string]$ConversationId,
        [Parameter(Mandatory=$true)][string]$ThreadId,
        [Parameter(Mandatory=$true)][string]$PostId
    )

    begin
    {
        $attachments = @()
        $headers = Get-AuthenticationHeaders -AccessToken $AccessToken
        $uri = "https://graph.microsoft.com/v1.0/groups/$GroupId/conversations/$ConversationId/threads/$ThreadId/posts/$PostId/attachments" 
    }
    process
    {
        do
        {
            $json = Invoke-RestMethodWithRetry @{ Uri = $uri; Headers = $headers; Method = "GET" }
            $attachments += $json.value
            $uri = $json.'@odata.nextLink'
        }
        while( $uri )
    }
    end
    {
       $attachments
    }
}

# script requires Group.Read.All
# script requires user to have Read/Write access to target Document Library

# Requires User Delegated Authentication to work with Group Conversations (app-only not supported)
# See Known Issues: https://docs.microsoft.com/en-us/graph/known-issues#groups


# Ensure we have a GroupId or we have a valid CSV Path
if ( -not $GroupId -and -not (Test-Path $CSVPath) ) { 
    break 
}

$GroupsToProcess = @()

# Process Single Group
if ( $GroupId ) {
    $GroupsToProcess += @{
        GroupId = $GroupId
        SiteUrl = $SiteUrl
        LibraryName = $LibraryName
        FolderName = $FolderName
    }
}
# Process Multiple Groups from CSV
else {
    $GroupsCSV = ConvertFrom-Csv (Get-Content $CSVPath)

    if ($null -eq $GroupsCSV) {
        break
    }

    foreach ($CSVRow in $GroupsCSV) {
        if ($CSVRow.id) {
            $GroupsToProcess += @{
                GroupId =       $CSVRow.id
                SiteUrl =       if ($CSVRow.siteUrl) { $CSVRow.siteUrl } else { "" }
                LibraryName =   if ($CSVRow.libraryName) { $CSVRow.libraryName } else { "Shared Documents" }
                FolderName =    if ($CSVRow.folderName) { $CSVRow.folderName } else { "" }
            }
        }
    }
}

$token = $null

foreach ($Current in $GroupsToProcess) {
    Write-Host "Current Group: $($Current.GroupId)" -ForegroundColor Green

    $groupSiteUrl = $null
    if ( -not $Current.SiteUrl ) {
        if ( -not $token ) {
            Connect-PnPOnline -Scopes "Group.Read.All"
        }

        $group = Get-PnPUnifiedGroup -Identity $Current.GroupId
        $groupSiteUrl = $group.SiteUrl
    }
    elseif ( -not $token ) {
        Connect-PnPOnline -Scopes "Group.Read.All" -Url $Current.SiteUrl -UseWebLogin
    }

    $token = Get-PnPAccessToken

    if( -not $token -or ($Current.SiteUrl -and -not (Get-PnPWeb)) ) {
        break
    }

    $conversations = Get-ConversationsWithAttachments -AccessToken $token -GroupId $Current.GroupId
    $hasEnsuredSiteAndLibraryExists = $false

    foreach ($conversation in $conversations) {
        $threads = Get-Threads -AccessToken $token -GroupId $Current.GroupId -ConversationId $conversation.id

        Write-Host "  Conversation: $($conversation.topic)"

        $threadsWithAttachments = $threads | ? {$_.hasAttachments}
        foreach ($thread in $threadsWithAttachments) {
            $posts = Get-Posts -AccessToken $token -GroupId $Current.GroupId -ConversationId $conversation.id -ThreadId $thread.id

            $postsWithAttachments = $posts | ? {$_.hasAttachments}
            foreach ($post in $postsWithAttachments) {
                $attachments = Get-Attachments -AccessToken $token -GroupId $Current.GroupId -ConversationId $conversation.id -ThreadId $thread.id -PostId $post.id

                if ($attachments -and @($attachments).Count -gt 0 -and -not $hasEnsuredSiteAndLibraryExists) {

                    try { Get-PnPConnection | Out-Null } catch { Connect-PnPOnline -Url $groupSiteUrl -UseWebLogin }

                    $existingLibrary = Get-PnPList $Current.LibraryName

                    if ($null -eq $existingLibrary) {
                        New-PnPList -Title $Current.LibraryName -Template DocumentLibrary | Out-Null
                    }

                    $hasEnsuredSiteAndLibraryExists = $true
                }

                $hasEnsuredFolderExists = $false
                foreach ($attachment in $attachments) {
                    $bytes = $null
                    $stream = $null
                    try {
                        $bytes = [System.Convert]::FromBase64String($attachment.contentBytes)
                        $stream = [IO.MemoryStream]::new($bytes)
                        $utcDateFormatted = (Get-Date $thread.lastDeliveredDateTime).ToUniversalTime().ToString("yyyyMMddHHmmss")
                        $folderPath = "$($Current.LibraryName)/$($Current.FolderName)/$($thread.topic)_$utcDateFormatted".Replace("//", "/")
        
                        Write-Host "    Uploading Attachment: '$($attachment.name)' to '$((Get-PnPWeb).ServerRelativeUrl)/$folderPath'"
                        if ( -not $hasEnsuredFolderExists ) {
                            Resolve-PnPFolder -SiteRelativePath $folderPath | Out-Null
                            $hasEnsuredFolderExists = $true
                        }
                        Add-PnPFile -FileName $attachment.name -Folder $folderPath -Stream $stream | Out-Null
                    }
                    finally {
                        $bytes = $null
                        if ($stream) { $stream.Close() }
                    }
                }
            }
        }
    }
    Write-Host
}