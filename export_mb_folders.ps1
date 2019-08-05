param (
    [string]$outfilename = "C:\temp\Maks\explore-mailboxes"    
 )
 



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"
Write-Host "outfile: " $outfile

if (Test-Path $outfile) 
{
    Remove-Item $outfile
}

Get-Mailbox -resultSize Unlimited | Select-Object Alias, Name, DisplayName, ServerName |
foreach-object {
    try
    {
        $cur = $_
        $stat = Get-MailboxFolderStatistics -Identity $cur.alias
        
        $p = Get-MailboxFolderPermission -Identity ($cur.alias + ":\") -ErrorAction SilentlyContinue | Select Identity, FolderName, User, AccessRights
        if ($p -ne $null) {

          [string]$t = "["
          $p| foreach-object {
            if ($t.length -ne 1) {
                $t += ","
            }
            $t += "{Identity: '$($_.Identity)',"
            $t += "FolderName: '$($_.FolderName)',"
            $t += "User: '$($_.User)',"
            $t += "AccessRights: '$($_.AccessRights)'}"
          }
          $t += ","
          Write-Host "permision: " $t
          $cur | Add-Member -MemberType NoteProperty -Name Permissions -Value $t -Force
        }
    
        $cur | Add-Member -MemberType NoteProperty -Name Folders -Value $stat -Force
        $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
   }
   Catch {
            $msg = "Error + $PSItem.Exception.InnerExceptionMessage"
            Write-Host $msg -ForegroundColor Red
  }
}
