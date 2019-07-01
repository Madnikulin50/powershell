param (
    [string]$folder = "d:\test\Files_DB",
    [string]$outfilename = "explore-folder",
    [switch]$force = $false
 )
 



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"

Write-Host "base: " $base
Write-Host "server: " $server
Write-Host "outfile: " $outfile

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}

#powershell.exe -ExecutionPolicy Bypass -Command "./explore_share.ps1" -outfilename explore_folder

Get-ChildItem $folder -Recurse | 
Foreach-Object {
    $cur = $_ | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode"
    Write-Host $_.FullName
    $acl = Get-Acl $_.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
    $hash = Get-FileHash $_.FullName | Select-Object -Property "Algorithm", "Hash"
    $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
    $cur | Add-Member -MemberType NoteProperty -Name Hash -Value $hash -Force
    $cur
} | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append