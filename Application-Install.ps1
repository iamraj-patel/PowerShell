$c = Get-ADComputer -Filter 'Name -like "idx-col-5"'

$computers = $c.Name

$sourcefile = "\\win-2019-1\File Share\CollectorConnectWiseControl.msi"

foreach($computer in $computers)
{
    $destination = "\\$computer\C$\Users\Public\Downloads"

    if(!(Test-Path -Path $destination))
    {
        New-Item $destination -Type Directory
    }
    Copy-Item -Path $sourcefile -Destination $destination
    
    Invoke-Command -ComputerName $computer -ScriptBlock{ Start-Process  msiexec.exe  -Wait -ArgumentList '/i "C:\Users\Public\Downloads\CollectorConnectWiseControl.msi" /qr' }
}

