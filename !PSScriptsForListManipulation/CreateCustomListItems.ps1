Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

$tenantUrl = "https://intranet.sharepoint.com"
$siteUrl = "https://intranet.sharepoint.com/sites/sampleListAndForm"

#Get the Web
$web = Get-SPWeb -identity $siteUrl

#Get the Target List
$List = $web.Lists["State"]

#Create items
$spItem = $List.AddItem()
$spItem["Title"] = "New"
$spItem.Update()
$desc_created++

$spItem = $List.AddItem()
$spItem["Title"] = "For approval"
$spItem.Update()
$desc_created++

$spItem = $List.AddItem()
$spItem["Title"] = "Approved"
$spItem.Update()
$desc_created++

$spItem = $List.AddItem()
$spItem["Title"] = "Canceled"
$spItem.Update()
$desc_created++

write-host 'Created items: ' $desc_created
write-host '---------------------------------------'
write-host 'FINISHED'
write-host '---------------------------------------'