param (
    [string]$base = 'DC=acme,DC=local',
    [string]$server = 'acme.local',
    [string]$outfilename = 'export_ad',
    [string]$user = "",
    [string]$pwd = "",
    [switch]$force = $false,
    [string] $start = ""
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


if ($user -ne "") {
    $pass = ConvertTo-SecureString -AsPlainText $pwd -Force    
    $GetAdminact = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $user, $pass    
} else {
    $GetAdminact = Get-Credential
}

$domain = Get-ADDomain -server $server -Credential $GetAdminact

Write-Host "domain: " $domain.NetBIOSName

$domain | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append


if ($start -ne "") {
  Write-Host "start: " $start
  $starttime = [datetime]::ParseExact($start,'yyyyMMddHHmmss', $null)
}

function Get-ADPrincipalGroupMembershipRecursive() {
  Param(
      [string] $dsn,
      [array]$groups = @()
  )

  $obj = Get-ADObject -server $server  -Credential $GetAdminact $dsn -Properties memberOf

  foreach( $groupDsn in $obj.memberOf ) {

      $tmpGrp = Get-ADObject -server $server  -Credential $GetAdminact $groupDsn -Properties * | Select-Object "Name", "cn", "distinguishedName", "objectSid", "DisplayName", "memberOf"

      if( ($groups | where { $_.DistinguishedName -eq $groupDsn }).Count -eq 0 ) {
          $add = $tmpGrp 
          $groups +=  $tmpGrp           
          $groups = Get-ADPrincipalGroupMembershipRecursive $groupDsn $groups
      }
  }

  return $groups
}

Get-ADGroup -server $server `
-Credential $GetAdminact -searchbase $SearchBase `
-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} | Select-Object "Name", "GivenName", "Surname", "sn", "cn", "distinguishedName",
"whenCreated", "whenChanged", "memberOf", "objectSid", "DisplayName", 
"sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
"Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
"Mail", "userAccountControl", "Manager", "ObjectClass", "logonCount", "UserPrincipalName"| Foreach-Object {
  $cur = $_ 
  if ($start -ne "") {
    if ($cur.whenChanged -lt $starttime) {
      Write-Host "skip " $cur.Name
      return
    }

  }

  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force
  
  $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
  $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

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
"Manager", "Enabled", "lastlogondate", "ObjectClass", "logonCount", "LogonHours", "UserPrincipalName" | Foreach-Object {
  $cur = $_  
  if ($start -ne "") {
    if (($cur.whenChanged -lt $starttime) -and ($cur.lastlogondate -lt $starttime)){
      Write-Host "skip " $cur.Name
      return
    }
    Write-Host "write " $cur.Name

  }

  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"

  if ($cur.thumbnailPhoto -ne $null) {
    $cur.thumbnailPhoto =[Convert]::ToBase64String($cur.thumbnailPhoto)
  }

  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force

  $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
  $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "users export finished to: " $outfile


# | Select-Object "Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
#"sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
#"PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "lastlogondate", "ObjectClass"

Get-ADComputer -Filter * -Properties * -server $server  -Credential $GetAdminact -searchbase $SearchBase |
 Select-Object "Name", "dn", "sn", "cn", "distinguishedName", "whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
"sAMAccountName", "IPv4Address", "IPv6Address", "OperatingSystem", "OperatingSystemHotfix", "OperatingSystemServicePack", "OperatingSystemVersion",
"PrimaryGroup", "ManagedBy", "userAccountControl", "Enabled", "lastlogondate", "ObjectClass", "DNSHostName", "ObjectCategory", "LastBadPasswordAttempt"| Foreach-Object {
  $cur = $_
  if ($start -ne "") {
    if (($cur.whenChanged -lt $starttime) -and ($cur.lastlogondate -lt $starttime)) {
      Write-Host "skip " $cur.Name
      return
    }
    Write-Host "write " $cur.Name

  }
  $ntname = "$($domain.NetBIOSName)\$($cur.sAMAccountName)"
  $cur | Add-Member -MemberType NoteProperty -Name NTName -Value $ntname -Force
  
  $allGroups = ADPrincipalGroupMembershipRecursive $cur.DistinguishedName 
  $cur | Add-Member -MemberType NoteProperty -Name AllGroups -Value $allGroups -Force

  $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append
}

Write-Host "computers export finished to: " $outfile


