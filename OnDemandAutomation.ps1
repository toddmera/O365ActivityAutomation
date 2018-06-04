############################################################
# Tenant Information
$tenantName = 'M365x534198'
$tenantPassword = Read-Host "Enter Your Tenant Password" 
$adminRoleName = 'Company Administrator'
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
        Write-Host "Module Exists"
    } else {
        Write-Host "Module Does Not Exist"
        Import-Module MSOnline
    }

    $admin ='Admin@' + $tenantName + '.onmicrosoft.com'
    $Pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $admin, $Pass

    Connect-MsolService -Credential $cred
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication  Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber

}


function Connect-RandomAdmin ([string]$randomAdmin){
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
    $Pass = ConvertTo-SecureString -String $tenantPassword -AsPlainText -Force
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $admin, $Pass

    Connect-MsolService -Credential $cred

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication  Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber

    
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
    $getAdmin = Get-Random $members.emailaddress
    Write-Host "New Administrator is $getAdmin"
    
    Return $getAdmin
}

############################################################
# Connect as tenant admin to start the whole thing off.
Get-InitialConnection

# Get a list of user from the $AdminRoleName 
$role = Get-MsolRole -RoleName $adminRoleName
$members = Get-MsolRoleMember -RoleObjectId $role.ObjectId -MemberObjectTypes "User"
############################################################



##### Quick and Dirty Tests ##########


# $myNewAdmin = Get-NewAdmin
# Connect-RandomAdmin -randomAdmin $myNewAdmin

$members



