<# This script requires SharePoint Server PowerShell Commands #>
param ([Switch]$WhatIf)

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue | Out-Null

function Delete-ListItemInView
{
    [cmdletbinding()]
    param
    (
        [parameter(Mandatory=$true)][Microsoft.SharePoint.SPList]$List,
        [parameter(Mandatory=$true)][Microsoft.SharePoint.SPView]$View
    )

    begin
    {
    }
    process
    {
        $items = $list.GetItems($view)

        $tempItems = @()
        $items | % { $tempItems += $_ }
        
        Write-Output "Deleting $($items.Count) items from $($View.Title)..."
        foreach( $item in $tempItems )
        {
            if( -not $WhatIf.IsPresent )
            {
                $item.Delete()
            }
            else
            {
                Write-Output "Would have deleted item with ID: $($item.Id) and Title: $($item.Name)"
            }
        }
    }
    end
    {
    }
}

#note that the view page or view limit DOES affect the items deleted
$webUrl    = "http://sp16.dev.local/sites/team"

$listTitle = "LargeList"
$viewTitle = "All Items" 

if( $web = Get-SPWeb -Identity $webUrl -ErrorAction SilentlyContinue )
{
    if( $list = $web.Lists.TryGetList( $listTitle ) )
    {
        if( $view = $list.Views | ? Title -eq $viewTitle )
        {
            Delete-ListItemInView -List $list -View $view
        }
        else
        {
            Write-Host "View not found" -ForegroundColor Red
        }
    }
    else
    {
        Write-Host "List not found" -ForegroundColor Red
    }
}
else
{
    Write-Host "Web not found" -ForegroundColor Red
}
