param (
    $SiteUrl = "https://contoso.sharepoint.com/sites/sts0", 
    $ListTitle = "LargeList",
    $ItemStart = 1,
    $ItemEnd = 4999
)

Connect-PnPOnline $SiteUrl

$List = Get-PnPList $ListTitle

$ItemStart..$ItemEnd | ForEach-Object { 
    $lici = [Microsoft.SharePoint.Client.ListItemCreationInformation]::new(); 
    $item = $List.AddItem($lici); 
    $item["Title"] = "Item$_"; 
    $item.Update() 
    if ($_ % 100 -eq 0) {
        Write-Host "Loading Items: $($_-100) to $_"
        $List.Context.ExecuteQuery()
    }
}
Write-Host "Loading Last Batch"
$List.Context.ExecuteQuery()