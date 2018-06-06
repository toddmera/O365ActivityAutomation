############################################################
# Tenant Information
$tenantName = 'M365x534198'
$adminRoleName = 'Company Administrator'
$forwardingSMTPEmail = 'SomeAddress@Quest.com'

# Min task an admin will run during one session
$minAdminTasks = 5
# Max task an admin will run during one session
$maxAdminTasks = 10

# $tenantName = "put password here if you like.  You will have to comment out the line below and uncomment this one"
$tenantPassword = Read-Host "Enter you tenant password"
############################################################


############################################################
# Initialize variables
$companyAdmins = $null
$unlicenedUsers = $null
$licenedUsers = $null

$functionList = ("Set-ForwardingSMTP", "Remove-ForwardingSMTP", "Set-ForwardingSMTPAlias", "Remove-ForwardingSMTPAlias")

############################################################


function Get-InitialConnection {
     <#
    .SYNOPSIS
    Initial-Connection - Logs in as Tenant Admin and kicks off the process.

    .DESCRIPTION 
    We must connect as Admin and get a list of Company Administrators. 
    We also check to see if the MSOnline module is installed and if not install.

    .EXAMPLE
    Initial-Connection

    .NOTES
    Written by: Todd Mera

    * Website:	http://Quest.com

    #>

    if (Get-Module -ListAvailable -Name MSOnline) {
        Write-Host "MSOnline Module Exists and does not need to be installed"
    } else {
        Write-Host "MSOnline Module Does Not Exist and needs to be installed"
        Import-Module MSOnline
    }

    $admin ='Admin@' + $tenantName + '.onmicrosoft.com'
    $pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $admin, $pass

    Connect-MsolService -Credential $cred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication  Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber

}

function Connect-Admin ([string]$randomAdmin){
     <#
    .SYNOPSIS
    Connect-RandomAdmin - Connects a random admin from the Admin Role specified in $adminRoleName.

    .DESCRIPTION 
    Connects a random admin from the Admin Role specified in $adminRoleName.  

    .EXAMPLE
    Connect-RandomAdmin -randomAdmin <UserPrincipalName>

    .NOTES
    Written by: Todd Mera

    * Website:	http://Quest.com

    #>
    $admin = $randomAdmin
    $pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $admin, $pass

    Connect-MsolService -Credential $cred

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication  Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber

    Write-Host "New Admin is: $admin"

    
}

function Get-NewAdmin {
     <#
    .SYNOPSIS
    Get-NewAdmin - Picks a random admin from the Admin Role specified in $adminRoleName.

    .DESCRIPTION 
    Picks a random admin from the Admin Role specified in $adminRoleName.  

    .EXAMPLE
    Get-NewAdmin

    .NOTES
    Written by: Todd Mera

    * Website:	http://Quest.com

    #>
    
    # Select the principal name for a random user in the $Members object
    $getAdmin = Get-Random $companyAdmins.emailaddress 

    Return $getAdmin
}

function Get-CompanyAdmins {
     <#
    .SYNOPSIS
    Get-CompanyAdmins - Returns a list of users with the Admin Role specified in $adminRoleName.

    .DESCRIPTION 
    Get-CompanyAdmins - Returns a list of users with the Admin Role specified in $adminRoleName.

    .EXAMPLE
    Get-CompanyAdmin

    .NOTES
    Written by: Todd Mera

    * Website:	http://Quest.com

    #>
    # Get a list of user from the $AdminRoleName and return list
    $role = Get-MsolRole -RoleName $adminRoleName
    $roleMembers = Get-MsolRoleMember -RoleObjectId $role.ObjectId -MemberObjectTypes "User" 
    return $roleMembers | Where-Object {$_.emailaddress -notlike 'admin@*'}
    
}


function Set-ForwardingSMTP {
    # Get a list of users that do not have ForwardingSMTPAddress set and set this option.
    $noForwardMailboxes = Get-Mailbox | Where-Object {($_.ForwardingSMTPAddress -eq $null -and $_.RecipientTypeDetails -eq "UserMailbox" -and $_.Name -notlike "admin*")} 

    if ($noForwardMailboxes){
        # Get a random mailbox
        $randomMailbox = Get-Random -InputObject $noForwardMailboxes
        
        # Set forwardingsmtpaddress - This attribute is displayed in the Exchange Admin Portal
        Set-Mailbox -Identity $randomMailbox.Alias -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $forwardingSMTPEmail

        Write-Host "Mail for $randomMailbox has been forwarded to $forwardingSMTPEmail"
    }
   
}

function Remove-ForwardingSMTP {
     # Get list of users that have email forwarding set and turn it off
     $forwardMailboxes = Get-Mailbox | Where-Object {($_.ForwardingSMTPAddress -and $_.RecipientTypeDetails -eq "UserMailbox")} | Sort-Object -Property Name 

     if ($forwardMailboxes){
        # Get a random mailbox
        $randomMailbox = Get-Random -InputObject $forwardMailboxes

        # Remove the forwarding option
        Set-Mailbox -Identity $randomMailbox.Alias -DeliverToMailboxAndForward $false -ForwardingSMTPAddress $null

        Write-Host "Mail for $randomMailbox has been set to Null"

     }


}

function Set-ForwardingSMTPAlias {
     # Get a list of users that do not have ForwardingAddress set and set this option.
     $noForwardMailboxes = Get-Mailbox | Where-Object {($_.ForwardingAddress -eq $null -and $_.RecipientTypeDetails -eq "UserMailbox" -and $_.Name -notlike "admin*")} 

     if ($noForwardMailboxes){
         # Get a random mailbox
         $randomMailbox = Get-Random -InputObject $noForwardMailboxes
         $forwardToAlias = Get-Random -InputObject $noForwardMailboxes

         if ($randomMailbox -ne $forwardToAlias) {
         # Set forwardingaddress - This attribute is NOT displayed in the Exchange Admin Portal.  This is Outlook Rule.
         Set-Mailbox -Identity $randomMailbox.Alias -DeliverToMailboxAndForward $true -ForwardingAddress $forwardToAlias.Alias
 
         Write-Host "Mail for $randomMailbox has been forwarded to $forwardToAlias"
         }
     }


}

function Remove-ForwardingSMTPAlias {
    # Get list of users that have email forwarding set and turn it off
    $forwardMailboxes = Get-Mailbox | Where-Object {($_.ForwardingAddress -and $_.RecipientTypeDetails -eq "UserMailbox")} | Sort-Object -Property Name 

    if ($forwardMailboxes){
       # Get a random mailbox
       $randomMailbox = Get-Random -InputObject $forwardMailboxes

       # Remove the forwarding option
       Set-Mailbox -Identity $randomMailbox.Alias -DeliverToMailboxAndForward $false -ForwardingAddress $null

       Write-Host "Mail for $randomMailbox has been set to Null"
    }

}


function Start-RandomActivity {
    # Start some random activity with a new admin.
    for ($i=0; $i -le (Get-Random -Minimum $minAdminTasks -Maximum $maxAdminTasks); $i++){
        $newAdmin = Get-NewAdmin
        Connect-Admin -randomAdmin $newAdmin

        # Get a random function from the function list.
        for ($x=0; $x -le (Get-Random -Minimum $minAdminTasks -Maximum $maxAdminTasks); $x++){

            $randomFunction = Get-Random -InputObject $functionList
            Invoke-Expression $randomFunction
        }
        

         # Kill the session to prepare for new admin session
         Get-PSSession | Remove-PSSession
    }

    
    # Kill the session to prepare for new admin session
    Get-PSSession | Remove-PSSession
        
}

############################################################
# Connect as tenant admin to start the whole thing off.
Get-InitialConnection
Get-PSSession

$companyAdmins = Get-CompanyAdmins

$companyAdmins | Format-Table

# Start-RandomActivity

############################################################







