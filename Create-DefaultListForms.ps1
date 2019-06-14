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
  Script to create default list forms (New, Edit, Display) if they are missing. 
  Will also re-add missing Form Webpart to existing list forms if they are missing. 
  Expects a CSV file called TargetListForms.csv to be in the same directory as the script. 
  CSV should have a siteUrl and listTitle columns.

 .DESCRIPTION
  Script to create default list forms (New, Edit, Display) if they are missing. 
  Will also re-add missing Form Webpart to existing list forms if they are missing. 
  Expects a CSV file called TargetListForms.csv to be in the same directory as the script. 
  CSV should have a siteUrl and listTitle columns.

  Tested with List Templates: 100, 102, 103, 104, 105, 106, 107, 108
  
  NOTE: This script requires the PowerShell module 'SharePointPnPPowerShellOnline' to be installed. If it is missing, the script will attempt to install it.

  RECOMMENDATION: Add administrator username and password for your tenant to Windows Credential Manager before running script. 
  https://github.com/SharePoint/PnP-PowerShell/wiki/How-to-use-the-Windows-Credential-Manager-to-ease-authentication-with-PnP-PowerShell

  .PARAMETER CSVPath
  Specify the path to a CSV file containing siteUrl and listTitle fields.

 .EXAMPLE
  .\Create-DefaultListForms.ps1
 
  Creates list forms for specified siteUrl and listTitle rows in a CSV file in the script directory called Create-DefaultListForms-Sample.csv

 .EXAMPLE
  .\Create-DefaultListForms.ps1 -CSVPath C:\temp\targetlists.csv
 
  Creates list forms for specified siteUrl and listTitle rows in a CSV file located at C:\temp\targetlists.csv
#>

param($CSVPath = "$(Split-Path -Parent -Path $MyInvocation.MyCommand.Definition)\Create-DefaultListForms-Sample.csv")

#############################################

$module = Get-Module SharePointPnPPowerShellOnline -ListAvailable
if ($null -eq $module) {
    Write-Output "Installing PowerShell Module: SharePointPnPPowerShellOnline"
    Install-Module SharePointPnPPowerShellOnline -Force -AllowClobber -Confirm:$false
}

#############################################

$webpartTemplate = @"
<WebPart xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns="http://schemas.microsoft.com/WebPart/v2">
    <Assembly>Microsoft.SharePoint, Version=16.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>  
    <TypeName>Microsoft.SharePoint.WebPartPages.ListFormWebPart</TypeName>
    <ListName xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{{{LIST_ID}}}</ListName>
    <ListId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{{LIST_ID}}</ListId>
    <PageType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{{PAGE_TYPE}}</PageType>
    <FormType xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{{FORM_TYPE}}</FormType>
    <ControlMode xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">{{CONTROL_MODE}}</ControlMode>
    <ViewFlag xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">1048576</ViewFlag>
    <ViewFlags xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">Default</ViewFlags>
    <ListItemId xmlns="http://schemas.microsoft.com/WebPart/v2/ListForm">0</ListItemId>
</WebPart>
"@

Function Create-DefaultListForm
{
    param(
        [parameter(Mandatory=$true)]$List, 
        [parameter(Mandatory=$true)][string]$FormUrl, 
        [parameter(Mandatory=$true)][ValidateSet("Display", "Edit", "New")]$FormType
    )

    begin { }    
    process
    {

        $webpartXml = $webpartTemplate -replace "{{LIST_ID}}", $List.Id.ToString()

        switch ($FormType)
        {
            "Display" { 
                $webpartXml = $webpartXml -replace "{{PAGE_TYPE}}", "PAGE_DISPLAYFORM" 
                $webpartXml = $webpartXml -replace "{{FORM_TYPE}}", "4"  
                $webpartXml = $webpartXml -replace "{{CONTROL_MODE}}", "Display"  
                break;
            }
            "Edit" { 
                $webpartXml = $webpartXml -replace "{{PAGE_TYPE}}", "PAGE_EDITFORM" 
                $webpartXml = $webpartXml -replace "{{FORM_TYPE}}", "6"  
                $webpartXml = $webpartXml -replace "{{CONTROL_MODE}}", "Edit"  
                break;
            }
            "New" { 
                $webpartXml = $webpartXml -replace "{{PAGE_TYPE}}", "PAGE_NEWFORM" 
                $webpartXml = $webpartXml -replace "{{FORM_TYPE}}", "8"  
                $webpartXml = $webpartXml -replace "{{CONTROL_MODE}}", "New"  
                break;
            }
        }

        try
        {           
            #Check if form page already exists
            $listPages = Get-PnPProperty -ClientObject $List.RootFolder -Property Files
            $formPage = $listPages | Where-Object { $_.ServerRelativeUrl.ToLower() -eq $FormUrl.ToLower() }

            if ($null -eq $formPage) {
                Write-Output "  [Creating Form Page] $FormUrl"

                #Create Form
                $formPage = $List.RootFolder.Files.AddTemplateFile($FormUrl, [Microsoft.SharePoint.Client.TemplateFileType]::FormPage)            
            }
            else {
                #Form page exists, check if form is recognized by list (i.e. form page has a form webpart on it)
                $listForms = Get-PnPProperty -ClientObject $List -Property Forms
    
                if ($null -ne $listForms -and $listForms.Count -gt 0) {
                    $existingForm = $list.Forms | Where-Object { $_.ServerRelativeUrl.ToLower() -eq $FormUrl.ToLower() }
                    if ($null -ne $existingForm) {
                        Write-Warning "  [Form Already Exists] $FormUrl"
                        return;
                    }                
                }
            }

            Write-Output "  [Adding Form Webpart] $FormUrl"
            #Get Webpart Manager for Form
            $wpm = $formPage.GetLimitedWebPartManager([Microsoft.SharePoint.Client.WebParts.PersonalizationScope]::Shared)

            #Import Webpart on page
            $wp = $wpm.ImportWebPart($webpartXml)

            #Add webpart to Form
            $wpm.AddWebPart($wp.WebPart, "Main", 1) | Out-Null

            #Execute changes
            $List.Context.ExecuteQuery()                    
        }
        catch
        {
            Write-Error "Error creating form $FormType at $FormUrl. Error: $($_.Exception)"
        }
    }
    end { }
}

#############################################

$csvRows = ConvertFrom-Csv (Get-Content $CSVPath)

if ($null -eq $csvRows) {
    Write-Error "Unable to find CSV at path $CSVPath"
    break
}

foreach ($row in $csvRows) 
{
    Write-Output "[$($row.siteUrl)] [$($row.listTitle)]"

    Connect-PnPOnline -Url $row.siteUrl
    $list = Get-PnPList $row.listTitle
    $listUrl = $list.RootFolder.ServerRelativeUrl
    
    Create-DefaultListForm -List $list -FormUrl "$listUrl/DispForm.aspx" -FormType Display
    Create-DefaultListForm -List $list -FormUrl "$listUrl/EditForm.aspx" -FormType Edit
    Create-DefaultListForm -List $list -FormUrl "$listUrl/NewForm.aspx"  -FormType New
}