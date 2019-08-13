Install-PnPApp

Get-PnPApp Get-PnPContentType

Get-PnPApp -Scope ToddDocLibApp

Get-PnPList -Includes "Title" | Select-Object title

Get-PnPListItem

$spoAppList = ("Security_Documents", "HR_Policy_and_Procedurs", "Product_Design_docs")

Get-PnPList -Identity ToddDocLibApp

New-PnPList -Title Announcements -Template Announcements

New-PnPList -Title MyContactList -Template Contacts

Add-PnPListItem -List "MyContactList" -Values @{"Title" = "Dixon"; "FirstName"="Bowen"}

$boo = Get-PnPListItem -List MyContactList -Id 2 
Get-PnPField -List MyContactList
