Get-PnPField -List "Demo Docs"

New-PnPList -Title MyDocLib001 -Template DocumentLibrary

$myDocumentPath = "D:\github\docs"

$docs = Get-ChildItem $myDocumentPath

$doc = Get-Random (Get-ChildItem $myDocumentPath)
$doc.FullName
$doc.Extension
$x = $myDocumentPath + "\" + $doc.Name
$x

Add-PnPFile -Path c:\temp\company.master -Folder "_catalogs/masterpage"

Add-PnPFile -Path $doc.FullName -Folder "ToddDocLibApp"

Remove-PnPList -Identity "ToddDocLibApp" -Force

Get-PnPList -Identity "Demo eeeee"

Get-ChildItem "$myDocumentPath" | Where-Object { $_.Extension -match '^*.pdf|^*.docx|^*.xlsx|^*.pptx'}

Get-ChildItem -Filter "*.dll" -Recurse | Where-Object { $_.Name -match '^MyProject.Data.*|^EntityFramework.*' }

Set-PnPListPermission -

Set-PnPListItemPermission -List 'Product Researc and Development' -Identity $ctx.PSCredential.UserName -User 'user@contoso.com' -AddRole 'Contribute' -ClearExisting

Set-PnPListPermission -Identity 'Product Researc and Development' -User $ctx.PSCredential.UserName -AddRole  'Full Control' -ClearExisting

$docLib = Get-Random $docLibraries
$docLib

$ctx = Get-PnPConnection
$ctx.PSCredential.UserName

$mydoclib = Get-PnPList -Identity $docLib
$mydoclib.BreakRoleInheritance($true, $true)
$mydoclib.Update()
$mydoclib.Context.Load($mydoclib)
$mydoclib.Context.ExecuteQuery()

$web = Get-PnPWeb -Identity "/" 
$spoList= Get-PnPList "Testlist" -Web $web 
$spoList.BreakRoleInheritance($true, $true)
$spoList.Update()

$spoList.Context.Load($spoList)
$spoList.Context.ExecuteQuery()

Get-PnPSubWebs -Includes "Title"

Remove-PnPWeb -Url "ProductResearch" -Force
Write-Host "2*z"