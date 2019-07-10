param (
    [string]$base = 'DC=acme,DC=local',
    [string]$server = 'acme.local',
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
$outfile = "$($outfilename)_$LogDate.json"

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}


$GetAdminact = Get-Credential

$domain = Get-ADDomain -server $server -Credential $GetAdminact

Write-Host "domain: " $domain.NetBIOSName

$domain | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append

Get-ADGroup -server $server `
-Credential $GetAdminact -searchbase $SearchBase `
-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} | Select-Object "Name", "GivenName", "Surname", "sn", "cn", "distinguishedName",
"whenCreated", "whenChanged", "memberOf", "objectSid", "DisplayName", 
"sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
"Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
"Mail", "userAccountControl", "Manager", "ObjectClass" | Foreach-Object {
  $cur = $_ 
  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force
  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "groups export finished to: " $outfile

Get-ADUser -server $server `
-Credential $GetAdminact -searchbase $SearchBase `
-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} | Select-Object "Name", "GivenName", "Surname", "sn", "cn", "distinguishedName",
"whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
"sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
"Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
"Mail", "userAccountControl", "PasswordNeverExpires", "PasswordExpired", "DoesNotRequirePreAuth",
"CannotChangePassword", "PasswordNotRequired", "TrustedForDelegation", "TrustedToAuthForDelegation",
"Manager", "Enabled", "lastlogondate", "ObjectClass" | Foreach-Object {
  $cur = $_  
  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"

  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force

  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "users export finished to: " $outfile


# | Select-Object "Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
#"sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
#"PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "lastlogondate", "ObjectClass"

Get-ADComputer -Filter * -Properties * -server $server  `
-Credential $GetAdminact -searchbase $SearchBase | Foreach-Object {
  $cur = $_
  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force  
  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "computers export finished to: " $outfile
