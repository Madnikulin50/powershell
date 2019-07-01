param (
    [string]$base = 'DC=testdomain,DC=local',
    [string]$server = 'kappa.testdomain.local',
    [string]$outfilename = 'export_ad',
    [switch]$force = $false
 )

Write-Host "base: " $base
Write-Host "server: " $server
Write-Host "outfile: " $outfile
#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm

Import-Module ActiveDirectory

$SearchBase = $base 
$outfile = "$($outfilename).json"

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}


$GetAdminact = Get-Credential


$AllADGroups = Get-ADGroup -server $server `
-Credential $GetAdminact -searchbase $SearchBase `
-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'}

$groups = $AllADGroups | Select-Object "Name", "GivenName", "Surname", "dn", "sn", "cn", "distinguishedName",
"whenCreated", "whenChanged", "memberOf", "objectSid", "DisplayName", 
"sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
"Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
"Mail", "userAccountControl", "PasswordNeverExpires", "PasswordExpired", "DoesNotRequirePreAuth",
"CannotChangePassword", "PasswordNotRequired", "TrustedForDelegation", "TrustedToAuthForDelegation",
"Manager", "Enabled", "ObjectClass"

$groups | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append

Write-Host "groups export finished to: " $outfile

$AllADUsers = Get-ADUser -server $server `
-Credential $GetAdminact -searchbase $SearchBase `
-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'}

$users = $AllADUsers | Select-Object "Name", "GivenName", "Surname", "dn", "sn", "cn", "distinguishedName",
"whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
"sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
"Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
"Mail", "userAccountControl", "PasswordNeverExpires", "PasswordExpired", "DoesNotRequirePreAuth",
"CannotChangePassword", "PasswordNotRequired", "TrustedForDelegation", "TrustedToAuthForDelegation",
"Manager", "Enabled", "lastlogondate", "ObjectClass"

$users | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append

Write-Host "users export finished to: " $outfile

$computers = Get-ADComputer -Filter * -server $server `
-Credential $GetAdminact -searchbase $SearchBase ` |
Select-Object "Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
"sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
"PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "lastlogondate", "ObjectClass"

$computers | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append

Write-Host "computers export finished to: " $outfile