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


$spoContactsLists = ("SupportContacts", "SupplierContacts", "HRContacts")
$spoContactItems = ('"Title" = "Mera"; "FirstName" = "Todd"','"Title" = "Smith"; "FirstName" = "Mark"')
$spoContactList = "SupportContacts"
$spoContactList = Get-Random $spoContactsLists
New-PnPList -Title $spoContactList -Template Contacts

$contactTitle = "Mera"
$contactFirstName = "Todd"
$contactEmail = $contactFirstName + "." + $contactTitle + "@qsft.com"

$spoContactItem = Get-Random $spoContactItems
$spoContactItem = 

Add-PnPListItem -List $spoContactList -Values @{"Title" = $contactTitle; "FirstName" = $contactFirstName; "Email" = $contactEmail}
Add-PnPListItem -List $spoContactList -Values @{"Title" = "Dixon"; "FirstName"="Bowen"}

$here = Get-PnPList -Identity $spoContactList 

if (Get-PnPList -Includes "Title" -Identity $spoContactList ) {
    Write-Host "Exists"
    
}else {
    Write-Host "Nope"
}

