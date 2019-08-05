############################################################
# Tenant Information
$tenantName = 'M365x109645'

$tenantPassword = Read-Host "Enter you tenant password"

$AdminSiteName = "https://$tenantName-admin.sharepoint.com/"
$spolSiteBaseURL = 'sharepoint.com/sites/'

$admin ='Admin@' + $tenantName + '.onmicrosoft.com'
$pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
$AzureSPOLCreds = New-Object System.Management.Automation.PSCredential ($admin, $pass)

Connect-SPOService $AdminSiteName -Credential $AzureSPOLCreds

Get-SPOSite
