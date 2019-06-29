﻿param (
    [string]$base = 'DC=testdomain,DC=local',
    [string]$server = 'kappa.testdomain.local',
    [string]$outfile = 'export_users.csv',
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

$AllADUsers |
Select-Object @{Label = "First Name";Expression = {$_.GivenName}},
@{Label = "Last Name";Expression = {$_.Surname}},
@{Label = "DN";Expression = {$_.dn}},
@{Label = "SN";Expression = {$_.sn}},
@{Label = "CN";Expression = {$_.cn}},
@{Label = "Distinguished Name";Expression = {$_.distinguishedName}},
@{Label = "When Created";Expression = {$_.whenCreated}},
@{Label = "When Changed";Expression = {$_.whenChanged}},
@{Label = "Member Of";Expression = {$_.memberOf}},
@{Label = "Bad Pwd Count";Expression = {$_.badPwdCount}},
@{Label = "SID";Expression = {$_.objectSid}},
@{Label = "Display Name";Expression = {$_.DisplayName}},
@{Label = "Logon Name";Expression = {$_.sAMAccountName}},
@{Label = "Full address";Expression = {$_.StreetAddress}},
@{Label = "City";Expression = {$_.City}},
@{Label = "State";Expression = {$_.st}},
@{Label = "Post Code";Expression = {$_.PostalCode}},
@{Label = "Country/Region";Expression = {if (($_.Country -eq 'GB')  ) {'United Kingdom'} Else {''}}},
@{Label = "Job Title";Expression = {$_.Title}},
@{Label = "Company";Expression = {$_.Company}},
@{Label = "Description";Expression = {$_.Description}},
@{Label = "Department";Expression = {$_.Department}},
@{Label = "Office";Expression = {$_.OfficeName}},
@{Label = "Phone";Expression = {$_.telephoneNumber}},
@{Label = "Photo";Expression = {$_.thumbnailPhoto}},
@{Label = "Email";Expression = {$_.Mail}},
@{Label = "UAC";Expression = {$_.userAccountControl}},
@{Label = "Password Never Expires";Expression = {$_.PasswordNeverExpires}},
@{Label = "Password Expired";Expression = {$_.PasswordExpired}},
@{Label = "Does Not Require Pre Auth";Expression = {$_.DoesNotRequirePreAuth}},
@{Label = "Cannot Change Password";Expression = {$_.CannotChangePassword}},
@{Label = "Password Not Required";Expression = {$_.PasswordNotRequired}},
@{Label = "Trusted For Delegation";Expression = {$_.TrustedForDelegation}},
@{Label = "Trusted To Auth For Delegation";Expression = {$_.TrustedToAuthForDelegation}},
@{Label = "Manager";Expression = {%{(Get-AdUser $_.Manager -server $ADServer -Properties DisplayName).DisplayName}}},
@{Label = "Account Status";Expression = {if (($_.Enabled -eq 'TRUE')  ) {'Enabled'} Else {'Disabled'}}}, # the 'if statement# replaces $_.Enabled
@{Label = "Last Logon Date";Expression = {$_.lastlogondate}} | 

#Export CSV report

Export-Csv -Path $csvfile -NoTypeInformation -Encoding UTF8

