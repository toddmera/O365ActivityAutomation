############################################################
# Tenant Information
$tenantName = 'M365x534198'
# $tenantName = Read-Host "Enter you tenant name ie. M365x534198"
$tenantPassword = Read-Host "Enter you tenant password"
$adminRoleName = 'Company Administrator'
############################################################


############################################################
# Initialize variables
$companyAdmins = $null
$unlicenedUsers = $null
$licenedUsers = $null


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


function Start-RandomActivity {

    for ($i=0; $i -le 10; $i++){
        $newAdmin = Get-NewAdmin
        Connect-Admin -randomAdmin $newAdmin

        $licenedUsers = Get-MsolUser -All | Where-Object {$_.IsLicensed -eq $true -and $_.UserPrincipalName -notlike "admin@*"} 
        
        




        Get-PSSession | Remove-PSSession

    }
    
}

############################################################
# Connect as tenant admin to start the whole thing off.
Get-InitialConnection

$companyAdmins = Get-CompanyAdmins

$companyAdmins | Format-Table

############################################################


# For ($i=0; $i -le 4; $i++){

#     $newAdmin = Get-NewAdmin

#     Write-Host $i " New current Admin is: " $newAdmin
#     Connect-Admin -randomAdmin $newAdmin
#     Write-Host "Session connected.  Will disconnect and start again."
#     Get-PSSession | Remove-PSSession

# } 

##### Quick and Dirty Tests ##########


# $myNewAdmin = Get-NewAdmin
# Connect-RandomAdmin -randomAdmin $myNewAdmin





