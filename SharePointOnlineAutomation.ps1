############################################################
# Tenant Information
$tenantName = 'M365x109645'

# $tenantName = "put password here if you like.  You will have to comment out the line below and uncomment this one"
$tenantPassword = Read-Host "Enter you tenant password"

# Number of cycles to pick random admin and perform tasks
$adminCycles = 50

# Min task an admin will run during one cycle
$minAdminTasks = 5
# Max task an admin will run during one cycle
$maxAdminTasks = 10

############################################################

############################################################
# SharePoint Sites, etc.

$AdminSiteURL = "https://$tenantName-admin.sharepoint.com/"
$CompanySiteURL = "https://$tenantName.sharepoint.com/"


function Get-InitialConnectionSPO {
    <#
   .SYNOPSIS
   Get-InitialConnectionSPO - Logs in as Tenant Admin and kicks off the process.

   .DESCRIPTION 
   We must connect as Admin and get a list of Company Administrators. 
   We also check to see if the MSOnline module is installed and if not install.

   .EXAMPLE
   Get-InitialConnectionSPO

   .NOTES
   Written by: Todd Mera

   * Website:	http://Quest.com

   #>

   if (Get-Module -ListAvailable -Name SharePointPnPPowerShellOnline) {
       Write-Host "SharePoint PnP Power Shell Online Module Exists and does not need to be installed"
   } else {
       Write-Host "Share Point PnP Power Shell Online Module Does Not Exist and needs to be installed"
       Install-Module SharePointPnPPowerShellOnline -AllowClobber
   }

   $admin ='Admin@' + $tenantName + '.onmicrosoft.com'
   $pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
   $AzureSPOLCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($admin, $pass)

   Connect-PNPOnline $CompanySiteURL -Credential $AzureSPOLCreds

}

function Connect-RandomSPOUser {
    $pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
    $AzureSPOLCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($newUser, $pass)
 
    Connect-PNPOnline $CompanySiteURL -Credential $AzureSPOLCreds

}

function Get-SPOUsers {
    <#
   .SYNOPSIS
   Get-SPOUsers - Returns a list of SPO users with non-blank email address.

   .DESCRIPTION 
   Get-SPOUsers - Returns a list of users.

   .EXAMPLE
    Get-SPOUsers

   .NOTES
   Written by: Todd Mera

   * Website:	http://Quest.com

   #>
   # Get a list of user from the $AdminRoleName and return list
   $spousers = Get-PnPUser | ? Email -ne ""
}

function Get-RandomSPOUser {
    
    # Get a random user and return email address.
    $getSPOUser = Get-Random $spousers.Email 

    Return $getSPOUser

}


function Start-RandomSPOActivity {


    $newUser = Get-RandomSPOUser
    $newUser
    Connect-RandomSPOUser

}

InitialConnectionSPO
