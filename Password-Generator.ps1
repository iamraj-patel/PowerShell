function Get-RandomPassword {
    $file = Import-Csv 'C:\Users\Raj\Downloads\words.csv'
    $words = $file.words
    $colors = $file.colors
    $specialCharacters = @("!", "@", "#", "$", "%", "^", "&", "*", "?", "_")
    $numbers = 0..9

    # Shuffle the arrays
    $words = $words | Get-Random -Count $words.Count
    $colors = $colors | Get-Random -Count $colors.Count
    $specialCharacters = $specialCharacters | Get-Random -Count $specialCharacters.Count
    $numbers = $numbers | Get-Random -Count $numbers.Count

    # Randomize the order of components
    $componentOrder = @(Get-Random -InputObject @(1, 2, 3, 4) -Count 4)

    # Construct the password based on the randomized order
    $password = ""
    foreach ($component in $componentOrder) {
        switch ($component) {
            1 { $password += $colors[0] }
            2 { $password += $words[0] }
            3 { $password += $specialCharacters[0] }
            4 { $password += $numbers[0] }
        }
    }

    return $password
}

$pass = Get-RandomPassword
Write-Output $pass
