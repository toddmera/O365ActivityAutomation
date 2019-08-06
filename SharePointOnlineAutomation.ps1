############################################################
# Tenant Information
$tenantName = 'M365x109645'

# $tenantName = "put password here if you like.  You will have to comment out the line below and uncomment this one"
# $tenantPassword = Read-Host "Enter you tenant password"
$tenantPassword = "q021Q8ExYU"

# Subwebs to create
$subwebs = ("ProductResearch", "Charity", "CarbonZeroProject", "ContosoSoftballTeam","CorpNews")
$spoSiteDesction = "This site was created by a script."

# Number of cycles to pick random user and perform tasks
$userCycles = 10

# Min task an user will run during one cycle
$minUserTasks = 5
# Max task an user will run during one cycle
$maxUserTasks = 10
############################################################
$randomSubWeb = ""

############################################################
# SharePoint Sites, etc.

# $AdminSiteURL = "https://$tenantName-admin.sharepoint.com/"
$CompanySiteURL = "https://$tenantName.sharepoint.com/"
############################################################


############################################################
# Function list.  We will randomly run these as different users.
# $spoFunctionList = ("Add-NewSubWeb", "Remove-SubWeb")
$spoFunctionList = ("AddRemove-SubWeb")
############################################################


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

function Get-SPOUsers {
    <#
   .SYNOPSIS
   Get-SPOUsers - Returns a list of SPO users with non-blank email address to $spoUsers.

   .DESCRIPTION 
   Get-SPOUsers - Call once and beggining of script to get a list of users.  Returns a list of users.

   .EXAMPLE
    Get-SPOUsers

   .NOTES
   Written by: Todd Mera

   * Website:	http://Quest.com

   #>
   # Get a list of user and return list
   $spoUsers = Get-PnPUser | Where-Object Email -ne ""
   Return $spoUsers
   
}

function Get-RandomSPOUser {
    
    # Get a random user and return email address.
    $getSPOUser = Get-Random $spousers.Email 

    Return $getSPOUser

}

function Connect-RandomSPOUser {([string]$randomUser)
    $pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
    $AzureSPOLCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($newUser, $pass)
 
    Connect-PNPOnline $CompanySiteURL -Credential $AzureSPOLCreds

}

function Add-NewSubWeb {

    Write-Host "############################"
    Write-Host "Adding some random subweb "
    Write-Host "############################"

    #Get a random site name from $subwebs

    Write-Host "####################################################"
    Write-Host "#"
    Write-Host "# New random subweb is $randomSubWeb    "
    Write-Host "#"
    Write-Host "####################################################"
    
    if (Get-PnPWeb -Identity $randomSubWeb) {
        Write-Host "$randomSubWeb Already exists.  This site will not be created"
    }else{
        Write-Host "$randomSubWeb does NOT exists.  $newUser will attempt to create this communications subsite."
        New-PnPWeb -Title $randomSubWeb -Url $randomSubWeb  -Template "SITEPAGEPUBLISHING#0"
        Start-Sleep -Seconds 20
    }
       
        
}

function Remove-SubWeb {

    Write-Host "############################"
    Write-Host "Removing some random subweb "
    Write-Host "############################"

    #Get a random site name from $subwebs
    Write-Host '$randomSubWeb = Get-Random $subwebs (Remove)'

    Write-Host "##############################################################################"
    Write-Host "#"
    Write-Host "# New random subweb is $randomSubWeb    "
    Write-Host "#"
    Write-Host "##############################################################################"
    
    if (Get-PnPWeb -Identity $randomSubWeb) {
        Write-Host "##############################################################################"
        Write-Host "#"
        Write-Host "# $randomSubWeb exists.  $newUser will remove this site"
        Write-Host "#"
        Write-Host "##############################################################################"
        Remove-PnPWeb -Url $randomSubWeb -Force
        Start-Sleep -Seconds 20
    }else{
        Write-Host "##############################################################################"
        Write-Host "#"
        Write-Host "# $randomSubWeb does not exist.  So we will just move on.  Deleting it would be like dividing by Zero."
        Write-Host "#"
        Write-Host "##############################################################################"
    }
        
}


function AddRemove-SubWeb {
    Write-Host "Running AddRemove Function"
    $randomSubWeb = "Charity"
    $site = Get-PnPWeb -Identity $randomSubWeb
    $site
    if ($site){
        Write-Host "True - Delete it"
        Remove-PnPWeb -Url $randomSubWeb -Force

    }else{
        Write-Host "False - Create it"
        New-PnPWeb -Title $randomSubWeb -Url $randomSubWeb  -Template "SITEPAGEPUBLISHING#0"


    }
    
    Start-Sleep -s 10
}

############################################################
# This is where it all happens.
function Start-SPORandomActivity {
    # Start some random activity with a new admin.
    for ($i=0; $i -le $userCycles; $i++){
    # for ($i=0; $i -le (Get-Random -Minimum $minAdminTasks -Maximum $maxAdminTasks); $i++){
        $newUser = Get-RandomSPOUser
        Connect-RandomSPOUser -randomUser $newUser

        # Get a random function from the function list.
        for ($x=0; $x -le (Get-Random -Minimum $minUserTasks -Maximum $maxUserTasks); $x++){
        
            
            $randomSPOFunction = Get-Random -InputObject $spoFunctionList
            Write-Host "***** Running $randomSPOFunction *****"
            Invoke-Expression $randomSPOFunction
        }

        # Disconnect User and start again.
        Disconnect-PnPOnline

    }
}

############################################################
# Connect as tenant admin to start the whole thing off.
Get-InitialConnectionSPO

# Get a list of users to play with.
Get-SPOUsers

# Let's make some random stuff happen
Start-SPORandomActivity

# Testing
# Get-InitialConnectionSPO
# Get-SPOUsers
# $newUser = Get-RandomSPOUser
# Connect-RandomSPOUser -randomUser $newUser
# $randomSPOFunction = Get-Random -InputObject $spoFunctionList
# Invoke-Expression $randomSPOFunction

# if (Get-PnPSubWebs -Identity $randomSubWeb.Title){Write-Host "$randomSubWeb Exists"}else{Write-Host "$randomSubWeb done NOT Exists"}