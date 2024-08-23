# Taking user inputs before creating user account.
$FirstName = Read-Host -Prompt "Enter FirstName of the new User "
$LastName = Read-Host -Prompt "Enter LastName of the new User "
$UserType = Read-Host -Prompt "Enter UserType of the new User (E.g: M for Mail-Only or D for Default) "
$Remote = Read-Host -Prompt "Does this user works remotely Please Enter (Y or N) "
$Username = ($FirstName.Substring(0, 1) + $LastName).ToLower()
$FullName = "$LastName, $FirstName"
$DomainSuffix = "@clariontechnologies.com"
$UsernameCounter = 0
$hostname = hostname

# Construct the base username (first initial + last name)
$BaseUsername = ($FirstName.Substring(0, 1) + $LastName).ToLower()

# Function to check if a user with a given username exists
function Test-UserExists {
    param (
        [string]$Username
    )
    try {
        Get-ADUser -Filter {SamAccountName -eq $Username} -ErrorAction Stop
    }
    catch {
        # Handle or log the error as needed
        return $false
    }
}
# Generating random password
#$Password = ConvertTo-SecureString "P@ssw0rd123" -AsPlainText -Force
function Get-RandomPassword {
    $file = Import-Csv '.\Downloads\words.csv'
    $words = $file.words
    $colors = $file.colors
    $specialCharacters = @("!", "@", "#", "$", "%", "^", "&", "*", "?", "_")
    $numbers = 0..9

    $word = Get-Random -InputObject $words
    $specialCharacter = Get-Random -InputObject $specialCharacters
    $color = Get-Random -InputObject $colors
    $number = Get-Random -InputObject $numbers

    $password = $color + $word + $specialCharacter + $number
    return $password
}
$Random = Get-RandomPassword
$Password = ConvertTo-SecureString "$Random" -AsPlainText -Force

# Check if the base username already exists
while (Test-UserExists -Username $BaseUsername) {
    $UsernameCounter++
    # Append the next character of the first name to the base username
    $BaseUsername = ($FirstName.Substring(0, $UsernameCounter + 1) + $LastName).ToLower()
    $Username = $BaseUsername
}

# Creating user account in specific OU.
if ($UserType -eq "M" -or $UserType -eq "m") {
    $OUPath = "OU=MailOnlyUsers,OU=Users - Clarion,DC=clariontechnologies,DC=com"
    New-ADUser -Name $FullName `
               -SamAccountName $Username `
               -UserPrincipalName "$Username$DomainSuffix" `
               -GivenName $FirstName `
               -Surname $LastName `
               -AccountPassword $Password `
               -Enabled $true `
               -EmailAddress "$Username$DomainSuffix" `
               -Path $OUPath
}
elseif ($UserType -eq "D" -or $UserType -eq "d") {
    $validPlantLocations = @(1, 2, 3, 4)
    while ($true) {
        $plant = Read-Host -Prompt "Enter Number of location for which this User belongs to which of the Plant (E.g: 1 for Garland(GTM) or 2 for Anderson(PPC) or 3 for Greenville(DMC) or 4 for Holland(M1)): "
        if ($plant -in $validPlantLocations) {
            break
        }
        else {
            Write-Host "Invalid plant location. Please try again."
        }
    }

    if ($plant -eq 1) {
        $OUPath = "OU=Garland,OU=Users - Clarion,DC=clariontechnologies,DC=com"
        New-ADUser -Name $FullName `
                   -SamAccountName $Username `
                   -UserPrincipalName "$Username$DomainSuffix" `
                   -GivenName $FirstName `
                   -Surname $LastName `
                   -AccountPassword $Password `
                   -Enabled $true `
                   -EmailAddress "$Username$DomainSuffix" `
                   -Path $OUPath
    }
    elseif ($plant -eq 2) {
        $OUPath = "OU=Anderson,OU=Users - Clarion,DC=clariontechnologies,DC=com"
        New-ADUser -Name $FullName `
                   -SamAccountName $Username `
                   -UserPrincipalName "$Username$DomainSuffix" `
                   -GivenName $FirstName `
                   -Surname $LastName `
                   -AccountPassword $Password `
                   -Enabled $true `
                   -EmailAddress "$Username$DomainSuffix" `
                   -Path $OUPath
    }
    elseif ($plant -eq 3) {
        $OUPath = "OU=Greenville,OU=Users - Clarion,DC=clariontechnologies,DC=com"
        New-ADUser -Name $FullName `
                   -SamAccountName $Username `
                   -UserPrincipalName "$Username$DomainSuffix" `
                   -GivenName $FirstName `
                   -Surname $LastName `
                   -AccountPassword $Password `
                   -Enabled $true `
                   -EmailAddress "$Username$DomainSuffix" `
                   -Path $OUPath
    }
    elseif ($plant -eq 4) {
        $OUPath = "OU=Holland,OU=Users - Clarion,DC=clariontechnologies,DC=com"
        New-ADUser -Name $FullName `
                   -SamAccountName $Username `
                   -UserPrincipalName "$Username$DomainSuffix" `
                   -GivenName $FirstName `
                   -Surname $LastName `
                   -AccountPassword $Password `
                   -Enabled $true `
                   -EmailAddress "$Username$DomainSuffix" `
                   -Path $OUPath
    }
}
else {
    Write-Host "You have not selected where you want to create user account."
}

# Adding user to Group if the user works Remotely.
if ($Remote -eq "Y" -or $Remote -eq "y") {
    Add-ADGroupMember -Identity "FSlogix Office Profile" -Members $Username
    Add-ADGroupMember -Identity "Remote Desktop Users" -Members $Username
}

# Start syncing newly created user to Office 365 Admin Center.
Invoke-Command -ComputerName $hostname -ScriptBlock { Start-ADSyncSyncCycle -PolicyType Delta }

# Exporting user information to CSV file
$Time = Get-Date
$Password_PlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))
$UserInfo = [PSCustomObject]@{
    Username = $BaseUsername
    Email = "$BaseUsername$DomainSuffix"
    #Password = "P@ssw0rd123"
    Password = $Password_PlainText
    DateTime = $Time
}

$CsvFilePath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\New-User-Info.csv"
try {
    $UserInfo | Export-Csv -Path $CsvFilePath -Append -NoTypeInformation -ErrorAction Stop
}
catch {
    Write-Warning "An error occurred while creating the CSV file: $_"
}