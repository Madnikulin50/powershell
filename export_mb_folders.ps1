
param (
    [string]$outfile = 'mb_folders.csv'
 )


Get-Mailbox -resultSize unlimited | Select-Object alias | foreach-object {Get-MailboxFolderStatistics -Identity $_.alias } | Export-csv -Path $outfile -NoTypeInformation -Encoding UTF8