# Importing Active Directory Module
Import-Module ActiveDirectory
Import-Module MSOnline
Import-Module ExchangeOnlineManagement

# Connect to MSOnline Online
$MsoSession = Connect-MsolService


# Connect to Exchange Online
$ExoSession = Connect-ExchangeOnline -UserPrincipalName raj.patel@motiotech.com

# Specify the domain controllers
#$domainControllers = @("DC1", "DC2", "DC3", "DC4")
$domainControllers = @("DC1")
$targetDC = "DC1"

# Import .csv file
$accounts = Import-Csv -Path "C:\Users\Administrator\Downloads\Accounts\test-deactivate.csv"
$accountArray = @()

$accounts | ForEach-Object {
    $firstname = $_.FirstName.ToLower()
    $lastname = $_.LastName.ToLower()
    $domain = $_.Domain.ToLower()
    $account = $firstname + "." + $lastname + "@" + $domain
    $accountArray += $account
}

# Converting User Mailbox to Shared Mailbox, Removing Licenses and Deactivating user account on AD
foreach ($account in $accountArray) {
    $emailToSearch = $account
    $user = Get-ADUser -LDAPFilter "(mail=$emailToSearch)" -Server $targetDC
    if ($user) {
        Write-Output "User with email $emailToSearch exists in Active Directory."
        # Convert to shared mailbox
        Set-Mailbox -Identity $emailToSearch -Type Shared -Session $ExoSession
        
        # Remove license
        $licensedUser = Get-MsolUser -UserPrincipalName $emailToSearch -Session $MsoSession
        if ($licensedUser.Licenses) {
            foreach ($license in $licensedUser.Licenses) {
                Write-Output "Removing license: $($license.AccountSkuId)"
                Set-MsolUserLicense -UserPrincipalName $emailToSearch -RemoveLicenses $license.AccountSkuId -Session $MsoSession
            }
        }
        Disable-ADAccount -Identity $user.SamAccountName -Server $targetDC
        Move-ADObject -Identity $user.ObjectGUID -TargetPath "OU=Disabled-Users,DC=motio,DC=local"
    } else {
        Write-Output "User with email $emailToSearch does not exist in Active Directory."
    }
}

# Manually sync all domain controllers after initial changes
foreach ($dc in $domainControllers) {
    repadmin /syncall /AdeP $dc
}

# Disconnect from MSOnline
Disconnect-MsolService

# Disconnect from Exchange Online
Disconnect-ExchangeOnline
