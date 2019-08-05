############################################################
# Tenant Information
$tenantName = 'M365x313780'
$adminRoleName = 'Company Administrator'
$forwardingSMTPEmail = 'SomeAddress@Quest.com'
$sharedMailboxes = ("Sales Deptartment", "Marketing Department", "Purchasing Department", "'RandD Department'", "HR Department", "Development Department", "Support Department")

# Number of cycles to pick random admin and perform tasks
$adminCycles = 50

# Min task an admin will run during one cycle
$minAdminTasks = 5
# Max task an admin will run during one cycle
$maxAdminTasks = 10

# $tenantName = "put password here if you like.  You will have to comment out the line below and uncomment this one"
$tenantPassword = Read-Host "Enter you tenant password"
############################################################


############################################################
# Initialize variables
$companyAdmins = $null

$functionList = ("Set-ForwardingSMTP", "Remove-ForwardingSMTP", `
                "Set-ForwardingAlias", "Remove-ForwardingAlias",`
                "Set-RandMailboxPermissions", "Remove-RandMailboxPermissions", `
                "Create-SharedMailbox", "Remove-SharedMailbox")

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
        Install-Module MSOnline
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
    $noForwardMailboxes = Get-Mailbox | Where-Object {($_.ForwardingSMTPAddress -eq $null -and $_.forwardingaddress -eq $null -and $_.RecipientTypeDetails -eq "UserMailbox" -and $_.Name -notlike "admin*")} 

    if ($noForwardMailboxes){
        # Get a random mailbox
        $randomMailbox = Get-Random -InputObject $noForwardMailboxes
        
        # Set forwardingsmtpaddress - This attribute does NOT get displayed in the Exchange Admin Portal
        Set-Mailbox -Identity $randomMailbox.Alias -DeliverToMailboxAndForward $true -ForwardingSMTPAddress $forwardingSMTPEmail

        Write-Host "Mail for $randomMailbox has been forwarded to $forwardingSMTPEmail"
    } else {
        Write-Host "Nothing to process"
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

        Write-Host "ForwardingSMTPAddress and DeliverToMailboxAndForward for $randomMailbox has been set to Null"

     } else {
        Write-Host "Nothing to process"
     }


}

function Set-ForwardingAlias {
     # Get a list of users that do not have ForwardingAddress set and set this option.
     $noForwardMailboxes = Get-Mailbox | Where-Object {($_.ForwardingAddress -eq $null -and $_.ForwardingSMTPAddress -eq $null -and $_.RecipientTypeDetails -eq "UserMailbox" -and $_.Name -notlike "admin*")} 

     if ($noForwardMailboxes){
         # Get a random mailbox
         $randomMailbox = Get-Random -InputObject $noForwardMailboxes
         $forwardToAlias = Get-Random -InputObject $noForwardMailboxes

         if ($randomMailbox -ne $forwardToAlias) {
         # Set forwardingaddress - This attribute is displayed in the Exchange Admin Portal.
         Set-Mailbox -Identity $randomMailbox.Alias -DeliverToMailboxAndForward $true -ForwardingAddress $forwardToAlias.Alias -Confirm:$false
 
         Write-Host "ForwardingAddress for $randomMailbox has been set to $forwardToAlias"
         }
     } else {
        Write-Host "Nothing to process"
     }


}

function Remove-ForwardingAlias {
    # Get list of users that have email forwarding set and turn it off
    $forwardMailboxes = Get-Mailbox | Where-Object {($_.ForwardingAddress -and $_.RecipientTypeDetails -eq "UserMailbox")} | Sort-Object -Property Name 

    if ($forwardMailboxes){
       # Get a random mailbox
       $randomMailbox = Get-Random -InputObject $forwardMailboxes

       # Remove the forwarding option
       Set-Mailbox -Identity $randomMailbox.Alias -DeliverToMailboxAndForward $false -ForwardingAddress $null

       Write-Host "ForwardingAddress and DeliverToMailboxAndForward for $randomMailbox has been set to Null"
    } else {
        Write-Host "Nothing to process"
    }

}

function Set-RandMailboxPermissions {
    # Get random mailboxes that have no permissions assigned to other users 
    $mbxs = Get-Mailbox | Where-Object {($_.RecipientTypeDetails -eq "UserMailbox" -and $_.Name -notlike "admin*")}

    # Get 2 random mailboxes - $mbxIdentity will be assigned full control over $mbxUser mailbox
    $mbxIdentity = Get-Random -InputObject $mbxs
    $mbxUser = Get-Random -InputObject $mbxs

    # Check to see if the mailboxes are the same.  If not, set permissions
    if ($mbxIdentity -ne $mbxUser) {
        Add-MailboxPermission -Identity $mbxIdentity.Alias -User $mbxUser.Alias -AccessRights FullAccess -InheritanceType ALL
        Write-Host "$mbxIdentity has Full Control of mailbox $mbxUser"
        
    } else {
        Write-Host "Nothing to process"
    }
    
}

function Remove-RandMailboxPermissions {
    # Get mailboxes that have been assigned permissions to another mailbox
    $mbxWithPerms = Get-Mailbox | Get-MailboxPermission | `
        Where-Object { `
            ($_.user.tostring() -ne "NT AUTHORITY\SELF") -and `
            ($_.user.tostring() -notlike "admin*") -and `
            ($_.user.tostring() -notlike "Discovery*") -and `
            ($_.IsInherited -eq $false)}

    if ($mbxWithPerms) {
        # Get a random mailbox
        $randomMailbox = Get-Random -InputObject $mbxWithPerms
        $mbxIdentity = $randomMailbox.Identity
        $mbxUser = $randomMailbox.User

        Remove-MailboxPermission -Identity $mbxIdentity -User $mbxUser -AccessRights FullAccess -InheritanceType ALL -Confirm:$false
        Write-Host "Mailbox permission for $mbxIdentity have been removed from $mbxUser"
    } else {
        Write-Host "Nothing to process"
    }

}

function Create-SharedMailbox {
    # Get array of share mailbox names from $sharedMailboxName in config.
    $sharedMailboxName = Get-Random $sharedMailboxes

    # Create a new random shared mailbox if it does not already exist.
    if (Get-Mailbox -RecipientTypeDetails SharedMailbox | Where-Object Name -EQ $sharedMailboxName) {
        Write-Host "$sharedMailboxName already exists.  This loop will exit."        
    } else {
        Write-Host "$sharedMailboxName does NOT exists.  We will attenpt to create this now."
        New-Mailbox -Shared -Name $sharedMailboxName -DisplayName $sharedMailboxName  `
                -Alias ($sharedMailboxName.Split(" ")[0] + ((Get-Random -Maximum 1000).ToString()).PadLeft(4,'0'))

    }

}    

function Remove-SharedMailbox {
    # Get a list of existing shared mailboxes.
    $sharedMailboxes = [array](Get-Mailbox -RecipientTypeDetails SharedMailbox)
    $randomSharedMailbox = Get-Random($sharedMailboxes)

    if ($sharedMailboxes) {
        # Get Random Shared Mailbox
        if ($sharedMailboxes.Length -gt 1) {           
            Remove-Mailbox -Identity $randomSharedMailbox.Alias -Confirm:$false
            
        } elseif ($sharedMailboxes.Length -eq 1) {
            Remove-Mailbox -Identity $randomSharedMailbox.Alias -Confirm:$false
     
        }
    }
      
}

function Start-RandomActivity {
    # Start some random activity with a new admin.
    for ($i=0; $i -le $adminCycles; $i++){
    # for ($i=0; $i -le (Get-Random -Minimum $minAdminTasks -Maximum $maxAdminTasks); $i++){
        $newAdmin = Get-NewAdmin
        Connect-Admin -randomAdmin $newAdmin

        # Get a random function from the function list.
        for ($x=0; $x -le (Get-Random -Minimum $minAdminTasks -Maximum $maxAdminTasks); $x++){

            $randomFunction = Get-Random -InputObject $functionList
            Write-Host "***** Running $randomFunction *****"
            Invoke-Expression $randomFunction
        }

         # Kill the session to prepare for new admin session
         Get-PSSession | Remove-PSSession
    }
}

############################################################
# Connect as tenant admin to start the whole thing off.
Get-InitialConnection
# Get-PSSession

$companyAdmins = Get-CompanyAdmins

$companyAdmins | Format-Table

Start-RandomActivity


############################################################







