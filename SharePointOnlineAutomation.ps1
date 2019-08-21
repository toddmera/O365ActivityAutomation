############################################################
# Tenant Information
#  You need to edit a few items.  Look for #--- EDIT THIS ---#

# If you want to enter you tenant each time uncomment the line below and comment out line after.  Example: $tenantName = 'M365x109645'
#--- EDIT THIS ---#
# $tenantName = Read-Host "Enter you tenant password"
$tenantName = 'TenatNameHere'  

# If you want to enter you password each time uncomment the line below and comment out line after.  Example: $tenantPassword = "qofmtur7f7"
#--- EDIT THIS ---#
# $tenantPassword = Read-Host "Enter you tenant password"
$tenantPassword = "PasswordHere"

# Number of cycles to pick random user and perform tasks.  Example, 5 would mean that 5 users will be randomly selected and connected to perform tasks.
# With the below parameters set to $spoUserCycles = 20, $spoMinUserTasks = 1 and $spoMaxUserTasks = 10 the process takes about 30 min.

#--- EDIT THIS ---#
$spoUserCycles = 2

# Min task an user will run during one cycle
#--- EDIT THIS ---#
$spoMinUserTasks = 1
# Max task an user will run during one cycle
#--- EDIT THIS ---#
$spoMaxUserTasks = 10
############################################################

############################################################
# Subwebs to create
$subwebs = ("ProductResearch", "Charity", "CarbonZeroProject", "ContosoSoftballTeam","CorpNews","Patents","SecurityIssues","Birthdays")
$spoSiteDesction = "This site was created by a script."

# List of SharePoint contact lists to create (Apps)
$spoContactsLists = ("SupportContacts", "SupplierContacts", "HRContacts")

# List of contact items to create
$contactTitles = ("Smith", "Johnson", "Williams", "Jones", "Brown", "Davis", "Miller", "Wilson", "Hopsin", "Millen")
$contactFirstNames = ("Lauran", "Flor", "Alexander", "Christine", "Lupita", "Jennine", "Rossie", "Laurel", "Vanda", "Cyril")
$contactEmailSuffix = "@qsft.com"

# This section for document library creation and file uploads.
# File path where you have some docs:
$myDocumentPath = "D:\github\docs"

# List of document libraries to create
$docLibraries = ("Product Research and Development", "ProjectX Design Documents", "Demo Resources and Tools", "Company Picnics")

############################################################


############################################################
# SharePoint Sites, etc.

$AdminSiteURL = "https://$tenantName-admin.sharepoint.com/"
$CompanySiteURL = "https://$tenantName.sharepoint.com/"

$tenant = $tenantName + ".onmicrosoft.com"
############################################################


############################################################
# Function list.  We will randomly run these as different users.
$spoFunctionList = ("CreateRemove-SubWeb", "CreateRemove-ContactList", "CreateRemove-DocumentLibraries")

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

function Connect-RandomSPOUser {
    param (
        [parameter(Mandatory)]
        [string]$randomUser
    )
    $user = $randomUser
    Write-Host "------------------------------------------" 
    Write-Host "--- Connecting new user: $user ---" 
    Write-Host "------------------------------------------" 
    Write-Host

    $pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
    $AzureSPOLCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList ($user, $pass)
    
    Write-Host "#--- Now trying to connnect new user with Connect-PNPOnline ---#"
    Connect-PNPOnline $CompanySiteURL -TenantAdminUrl $AdminSiteURL -Credential $AzureSPOLCreds

}


function CreateRemove-SubWeb {
    Write-Host "-----Getting random subweb-------"
    Write-Host
    $sposubweb = Get-Random $subwebs
    Write-Host "#--- New Subweb is: $sposubweb" ---#
    Write-Host "#--- Lets see if this exists and do the opposite ---#"
    if ((Get-PnPSubWebs -Includes "Title" | Where-Object {$_.Title -eq $sposubweb}).Title) {
        Write-Host "#--- $sposubweb SubWeb DOES exist so we can delete it! ---#" -BackgroundColor Red -ForegroundColor Yellow
        Remove-PnPWeb -Url $sposubweb -Force
    }else{
        Write-Host "#--- It does NOT exist.  Let's create $sposubweb subsite ---#"  -BackgroundColor Red -ForegroundColor Yellow
        New-PnPWeb -Title $sposubweb -Url $sposubweb -Description $spoSiteDesction -Locale 1033 -Template "PROJECTSITE#0"
    }
}

function CreateRemove-ContactList {
    Write-Host "-----Getting random Contacts List -------"
    Write-Host
    $spoContactList = Get-Random $spoContactsLists
    Write-Host "#--- New Contact List is: $spoContactList ---#"
    Write-Host "#--- Lets see if this exists and do the opposite ---#"
    if (Get-PnPList -Includes "Title" -Identity $spoContactList ) {
        Write-Host "#--- This contact list exists.  We will try to delete it. ---#" -BackgroundColor Blue -ForegroundColor Black
        Remove-PnPList -Identity $spoContactList -Force
        
    }else {
        Write-Host "#--- Contact List does not exist.  So let's add it and create some contacts. ---#"  -BackgroundColor Blue -ForegroundColor Black
        New-PnPList -Title $spoContactList -Template Contacts
        Start-Sleep -Seconds (Get-Random -Minimum 5 -Maximum 15)
        Write-Host "#--- Let's add some contacts ---#"  -BackgroundColor Blue -ForegroundColor Black
        $contactTitle = Get-Random $contactTitles
        $contactFirstName = Get-Random $contactFirstNames
        $contactEmail = $contactTitle + "." + $contactFirstName + $contactEmailSuffix
        Add-PnPListItem -List $spoContactList -Values @{"Title" = $contactTitle; "FirstName" = $contactFirstName; "Email" = $contactEmail}

    }

}

function CreateRemove-DocumentLibraries {
    # Let's get a random doc library name
    $docLib = Get-Random $docLibraries
    Write-Host "#--- New Document Library is $docLib ---#"

    if (Get-PnPList -Includes "Title" -Identity $docLib) {
        Write-Host "#--- $docLib DOES exist.  We will delete it. ---#" -BackgroundColor White -ForegroundColor Black
        Remove-PnPList -Identity $docLib -Force
        
    }else {
        Write-Host "#--- $docLib does NOT exist.  We can create it and add some docs. ---#" -BackgroundColor White -ForegroundColor Black
        New-PnPList -Title $docLib -Template DocumentLibrary
        if ((Get-ChildItem -File -Path $myDocumentPath | Measure-Object).Count -eq 0) {
            Write-Host "#--- Did not find any files in $myDocumentPath so we will move on. ---#"
            
        }else {
            $doc = Get-Random (Get-ChildItem -File -Path $myDocumentPath)
            Write-Host "#--- Adding $doc to $docLib ---#"
            Add-PnPFile -Path $doc.FullName -Folder $docLib
            
        }

        Write-Host "#--- We are going to break the Role Inheritance ---#"
        $doclibPermTarget = Get-PnPList -Includes "Title" -Identity $docLib
        $doclibPermTarget.BreakRoleInheritance($true, $true)
        $doclibPermTarget.Update()
        $doclibPermTarget.Context.Load($doclibPermTarget)
        $doclibPermTarget.Context.ExecuteQuery()
        Write-Host "#--- Done.  We broke the inheritance ---#"
        
    }
    
}

############################################################
# This is where it all happens.
function Start-SPORandomActivity {
    # Start some random activity with a new admin.
    for ($i=0; $i -le $spoUserCycles; $i++){
        Write-Host "#--- User cycle $i of $spoUserCycles ---#"
        $newSPOUser = Get-RandomSPOUser
        Write-Host "#--- New user is $newSPOUser ---#"
        Write-Host "#--- Let's try to connect $newSPOUser ---"
        Connect-RandomSPOUser -randomUser $newSPOUser
        Write-Host "#--- Cool, $newSPOUser is now connected" -BackgroundColor Cyan -ForegroundColor Yellow

        $randNumUserTasks = (Get-Random -Minimum $spoMinUserTasks -Maximum $spoMaxUserTasks)
        # Get a random function from the function list and run is some semi-random number of times...
        for ($x=0; $x -le $randNumUserTasks; $x++){
            Write-Host "#--- This user will perform $randNumUserTasks tasks"
            
            $randomSPOFunction = Get-Random -InputObject $spoFunctionList
            Write-Host "***** Running $randomSPOFunction.  This is task $x of $randNumUserTasks *****"
            Invoke-Expression $randomSPOFunction

            $sleepDuration = Get-Random -Minimum 4 -Maximum 10            
            Write-Host "#--- I so tired, I need to sleep for $sleepDuration seconds ---#" -BackgroundColor DarkBlue -ForegroundColor White
            Start-Sleep -Seconds $sleepDuration
        }

        # Disconnect User and start again.
        Write-Host "#--- Disconnecting $newSPOUser so we can start this again. ---#"
        Disconnect-PnPOnline -Connection $AzureSPOLCreds
        Write-Host "#--- $newSPOUser has been disconnected ---#"

    }
}


############################################################
$spoStartTime = Get-date

# Connect as tenant admin to start the whole thing off.    #
Get-InitialConnectionSPO

# Get a list of users to play with.
$spousers = Get-SPOUsers

# Let's make some random stuff happen
Start-SPORandomActivity
Write-Host "============ RUN IS COMPLETE =============="
$spoEndTime = Get-Date

$spoRunDuration = $spoEndTime - $spoStartTime
$spoTotalMinutes =  $spoRunDuration.TotalMinutes
Write-Host "#--- This script took $spoTotalMinutes minutes to run"