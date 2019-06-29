param (
    [string]$base = 'DC=testdomain,DC=local',
    [string]$server = 'kappa.testdomain.local',
    [string]$outfilename = 'export_users',
    [switch]$force = $false
 )

Write-Host "base: " $base
Write-Host "server: " $server
Write-Host "outfile: " $outfile
#Create a variable for the date stamp in the log file

$LogDate = get-date -f yyyyMMddhhmm

#Define CSV and log file location variables
#they have to be on the same location as the script

$csvfile = $outfile

#import the ActiveDirectory Module

Import-Module ActiveDirectory


#Sets the OU to do the base search for all user accounts, change as required.
#Simon discovered that some users were missing
#I decided to run the report from the root of the domain

$SearchBase = $base 

#Get Admin accountb credential

$GetAdminact = Get-Credential

#Define variable for a server with AD web services installed

$ADServer = $server

#Find users that are not disabled
#To test, I moved the following users to the OU=ADMigration:
#Philip Steventon (kingston.gov.uk/RBK Users/ICT Staff/Philip Steventon) - Disabled account
#Joseph Martins (kingston.gov.uk/RBK Users/ICT Staff/Joseph Martins) - Disabled account
#may have to get accountb status with another AD object

#Define "Account Status" 
#Added the Where-Object clause on 23/07/2014
#Requested by the project team. This 'flag field' needs
#updated in the import script when users fields are updated
#The word 'Migrated' is added in the Notes field, on the Telephone tab.
#The LDAB object name for Notes is 'info'. 

$AllADUsers = Get-ADUser -server $ADServer `
-Credential $GetAdminact -searchbase $SearchBase `
-Filter * -Properties * | Where-Object {$_.info -NE 'Migrated'} #ensures that updated users are never exported.

$users = $AllADUsers | Select-Object "Name", "GivenName", "Surname", "dn", "sn", "cn", "distinguishedName",
"whenCreated", "whenChanged", "memberOf", "badPwdCount", "objectSid", "DisplayName", 
"sAMAccountName", "StreetAddress", "City", "state", "PostalCode", "Country", "Title",
"Company", "Description", "Department", "OfficeName", "telephoneNumber", "thumbnailPhoto",
"Mail", "userAccountControl", "PasswordNeverExpires", "PasswordExpired", "DoesNotRequirePreAuth",
"CannotChangePassword", "PasswordNotRequired", "TrustedForDelegation", "TrustedToAuthForDelegation",
"Manager", "Enabled", "lastlogondate"

$users | ConvertTo-Json | Out-File -FilePath "$($outfilename).json" -Encoding UTF8 -Append

