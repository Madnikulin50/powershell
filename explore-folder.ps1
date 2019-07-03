param (
    [string]$folder = "d:\test\Files_DB",
    [string]$outfilename = "explore-folder",
    [switch]$force = $false
 )
 



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"

Write-Host "base: " $folder
Write-Host "outfile: " $outfile

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}

Function Get-FileHashTSO([String] $FileName,$HashName = "SHA1") 
{
  $FileStream = New-Object System.IO.FileStream($FileName,"Open", "Read") 
  $StringBuilder = New-Object System.Text.StringBuilder 
  [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($FileStream)|%{[Void]$StringBuilder.Append($_.ToString("x2"))} 
  $FileStream.Close() 
  $StringBuilder.ToString() 
}

#powershell.exe -ExecutionPolicy Bypass -Command "./explore_share.ps1" -outfilename explore_folder

Get-ChildItem $folder -Recurse | 
Foreach-Object {
    $cur = $_ | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode"
    Write-Host $_.FullName
    $acl = Get-Acl $_.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
    $path = $_.FullName
    
    
    Try
    {
        $hash = Get-FileHashTSO $path
    }
    Catch {
        Write-Host $PSItem.Exception.Message
        Try
        {
            $hash = Get-FileHash $path | Select-Object -Property "Hash"
        }
        Catch {
            Write-Host $PSItem.Exception.Message
        }
    }
    
    $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
    $cur | Add-Member -MemberType NoteProperty -Name Hash -Value $hash -Force
    $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append    
} 
