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
