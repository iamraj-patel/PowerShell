# Specify the path to the shared folder
#$sharedFolderPath = Read-Host -Prompt "Enter the path to the shared folder"
$sharedFolderPath = "C:\UPI\Users"

# Function to get the last access and modification times, and the user who accessed/modified the item
function Get-FileInfo($path) {
    $item = Get-Item -LiteralPath $path
    $lastAccessTime = $item.LastAccessTime
    $lastWriteTime = $item.LastWriteTime
    $lastAccessUser = (Get-ACL -LiteralPath $path).Access | Sort-Object -Property AccessControlType -Unique | Where-Object {$_.AccessControlType -eq 'Allow' -and $_.IdentityReference -ne 'NT AUTHORITY\SYSTEM'} | Select-Object -Last 1 -ExpandProperty IdentityReference

    [PSCustomObject]@{
        Path            = $item.FullName
        LastAccessTime  = $lastAccessTime
        LastAccessUser  = $lastAccessUser
        LastWriteTime   = $lastWriteTime
        LastWriteUser   = $lastAccessUser
        ItemType        = $item.Attributes
    }
}

# Recursively get information for files and folders within the shared folder
$results = Get-ChildItem -LiteralPath $sharedFolderPath -Recurse | ForEach-Object { Get-FileInfo -Path $_.FullName }

# Export results to a CSV file
$csvFilePath = Join-Path -Path $env:USERPROFILE -ChildPath "Desktop\SharedFolderAccessInfo.csv"
$results | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "File access and modification information has been exported to $csvFilePath"