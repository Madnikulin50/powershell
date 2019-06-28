param (
    [string]$base = 'DC=testdomain,DC=local',
    [string]$server = 'kappa.testdomain.local',
    [string]$outfilename = 'export_computers',
    [switch]$force = $false
 )

Write-Host "base: " $base
Write-Host "server: " $server
Write-Host "outfile: " $outfile
#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm


Write-Host "outfilename: " $outfilename


Import-Module ActiveDirectory
# For Windows 7
# https://blogs.msdn.microsoft.com/adpowershell/2009/03/24/active-directory-powershell-installation-using-rsat-on-windows-7/
# Windows Management Framework 5.1

#Sets the OU to do the base search for all user accounts, change as required.
#Simon discovered that some users were missing
#I decided to run the report from the root of the domain

$SearchBase = $base 

#Get Admin accountb credential

$GetAdminact = Get-Credential

#Define variable for a server with AD web services installed

$ADServer = $server

function Export-Shares($machine) {
    Write-Host "export shares from machine: " $machine
    $csvfiles = "$($outfilename)_shares.json"
    net view $machine | Select-Object -Skip  7 | Select-Object -SkipLast 2|
    ForEach-Object -Process {[regex]::replace($_.trim(),'\s+',' ')} |
    ConvertFrom-Csv -delimiter ' ' -Header 'sharename', 'type', 'usedas', 'comment' |
    foreach-object {
        Try {
            Get-Acl "\\$($machine)\$($_.sharename)"
        } Catch {
            # Do nothing
        }
    } |      
    Select-Object -Property "Path", "Owner", "Group", "Access" |
    ConvertTo-Json | Out-File -FilePath $csvfiles -Encoding UTF8 -Append
}



$computers = Get-ADComputer -Filter * -server $ADServer `
-Credential $GetAdminact -searchbase $SearchBase ` |
Select-Object "Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
"sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
"PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "lastlogondate"

$computers | ConvertTo-Json | Out-File -FilePath "$($outfilename)_computers.json" -Encoding UTF8 -Append
$computers | ForEach {Export-Shares $_.Name}
