Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

<#
# !!!!!!!!!!!!!! TEST !!!!!!!!!!!!!!!!!!!!!!!!!!!!!
$tenantUrl = "https://intranet.sharepoint.com"
$siteUrl = "https://intranet.sharepoint.com/sites/sampleListAndForm"
# !!!!!!!!!!!!!! TEST !!!!!!!!!!!!!!!!!!!!!!!!!!!!!
#>

$web = Get-SPWeb -identity $siteUrl
$List = $web.Lists["CustomList"]

$itemsCustomList = $List.items | % { $_.id}
$itemsCustomList | % {$List.items.DeleteItemById($_)}