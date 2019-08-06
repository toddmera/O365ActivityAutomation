############################################################
# Tenant Information
$tenantName = 'M365x109645'

# $tenantName = "put password here if you like.  You will have to comment out the line below and uncomment this one"
$tenantPassword = Read-Host "Enter you tenant password"

# Subwebs to create
$subwebs = ("Product_Research", "Charity", "Carbon_Zero_Project", "Contoso_Softball_Team")

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

function Add-NewSubWeb {

    #Get a random site name from $subwebs
    $randomSubWeb = Get-Random $subwebs

    Write-Host "####################################################"
    Write-Host "#"
    Write-Host "# New random subweb is $randomSubWeb    "
    Write-Host "#"
    Write-Host "####################################################"
    
    if (Get-PnPSubWebs -Identity $randomSubWeb.Title) {
        Write-Host "$randomSubWeb Already exists.  This site will not be created"
    }else{
        Write-Host "$randomSubWeb does NOT exists.  Mr. Robot will attempt to create this communications subsite."
        New-PnPWeb -Title $randomSubWeb -Url $randomSubWeb -Description $randomSubWeb -Locale 1033 -Template "SITEPAGEPUBLISHING#0"
    }
        
}

function Remove-SubWeb {

    #Get a random site name from $subwebs
    $randomSubWeb = Get-Random $subwebs

    Write-Host "##############################################################################"
    Write-Host "#"
    Write-Host "# New random subweb is $randomSubWeb    "
    Write-Host "#"
    Write-Host "##############################################################################"
    
    if (Get-PnPSubWebs -Identity $randomSubWeb) {
        Write-Host "##############################################################################"
        Write-Host "$randomSubWeb exists.  Ms. Pacwoman will remove this site"
        Write-Host "##############################################################################"
        Remove-PnPWeb -Url $randomSubWeb -Force
    }else{
        Write-Host "##############################################################################"
        Write-Host "$randomSubWeb does not exist.  So we will just move on.  Deleting it would be like dividing by Zero."
        Write-Host "##############################################################################"
    }
        
}

############################################################
# Connect as tenant admin to start the whole thing off.
InitialConnectionSPO

# Let's make some random stuff happen
# Start-RandomSPOActivity
