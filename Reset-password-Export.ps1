function Get-RandomPassword {
    $file = Import-Csv 'C:\Users\Raj\Downloads\words.csv'
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

$OUpath = 'ou=MailOnlyUsers,ou=Users - Clarion,dc=clariontechnologies,dc=com'
$ExportPath = 'C:\Users\Administrator\Downloads\password.csv'

# Initialize an array to store user information
$UserDetails = @()

# Get all users from the specified OU
$Users = Get-ADUser -Filter * -SearchBase $OUpath

# Loop through each user and generate a random password
foreach ($User in $Users) {
    $Pass = Get-RandomPassword

    # Create an object containing username and password
    $UserDetail = [PSCustomObject]@{
        UserName = $User.SamAccountName
        Password = $Pass
    }

    # Add the user details object to the array
    $UserDetails += $UserDetail

    # Reset the user's password
    Set-ADAccountPassword -Identity $User -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Pass -Force)
}

# Export the user details to a CSV file
$UserDetails | Export-Csv -Path $ExportPath -NoTypeInformation
