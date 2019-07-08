
[CmdLetBinding(DefaultParameterSetName = "NormalRun")]
Param(
    [Parameter(Mandatory = $False, Position = 1, ParameterSetName = "NormalRun")] $computers = ("acme.local"),
    [Parameter(Mandatory = $False, Position = 7, ParameterSetName = "NormalRun")] $outfilename = "events",
    [Parameter(Mandatory = $False, Position = 10, ParameterSetName = "NormalRun")] [ValidateSet("All","Logon","Service","User","Computer", "Clean", "File", "MSSQL", "RAS", "USB")] [array]$target="All"
)

$EventSource="All"
$EventLevel="All"
$NumberOfLastEventsToGet = 300
$EventLogName = ("Security")


$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"

Write-Host "computers: " $computers
Write-Host "outfilename: " $server
Write-Host "target: " $target
Write-Host "EventSource: " $EventSource
Write-Host "EventLevel: " $EventLevel
Write-Host "EventLogName: " $EventLogName

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}


$id_logon = ("4776","4672", "4624", "4634", "4800", "4801")
$id_service = ("7036","7031")
$id_usermagement = ("4720", "4722", "4723", "4724", "4725", "4726", "4738", "4740", "4767", "4780", "4794", "5376", "5377")
$id_computermagement = ("4720", "4722", "4725", "4726", "4738", "4740", "4767")
$id_clean = ("1102")
$id_file = ("4656", "4663", "4660", "4624")
$id_mssql = ("18456")
$id_printer = ("307")
$id_ras = ("20249", "20250", "20253", "20255", "20258", "20266", "20271", "20272")
$id_usb = ("2003")

      




$GetAdminact = Get-Credential

<# ------- SCRIPT_HEADER (Only Get-Help comments and Param() above this point) ------- #>
#Initializing a $Stopwatch variable to use to measure script execution
$stopwatch = [system.diagnostics.stopwatch]::StartNew()
#Using Write-Debug and playing with $DebugPreference -> "Continue" will output whatever you put on Write-Debug "Your text/values"
# and "SilentlyContinue" will output nothing on Write-Debug "Your text/values"
$DebugPreference = "Continue"
# Set Error Action to your needs
$ErrorActionPreference = "SilentlyContinue"
#Script Version
<# ---------------------------- /SCRIPT_HEADER ---------------------------- #>

<# -------------------------- DECLARATIONS -------------------------- #>
$FilterHashProperties = $null
$Answer = ""
$Events4All = @()
<# /DECLARATIONS #>
<# -------------------------- FUNCTIONS -------------------------- #>
function IsEmpty($Param){
    If ($Param -eq "All" -or $Param -eq "" -or $Param -eq $Null -or $Param -eq 0) {
        Return $True
    } Else {
        Return $False
    }
}


<# /FUNCTIONS #>
<# -------------------------- EXECUTIONS -------------------------- #>
Write-Host "Starting script..."
function ExportFor($eid) {
    $FilterHashProperties = @{
        LogName = $EventLogName
    }
    
    If (!(IsEmpty $EventSource)){
        $FilterHashProperties.Add('ProviderName',$EventSource)
    }
    
    If (!(IsEmpty $eid)){
        $FilterHashProperties.Add("ID",$eid)
    }
    
    If (!(IsEmpty $EventLevel)){
        for ($i=0;$i -lt $($EventLevel.count);$i++){
            $EventLevel[$i] = switch ($EventLevel[$i]) {
                "LogAlways" {0}
                "Critical" {1}
                "Error" {2}
                "Warning" {3}
                "Information" {4}
                "Verbose" {5}
            }
        }
        $FilterHashProperties.Add('Level',$EventLevel)
    }
    
    $msg = ("About to collect events on $($computers.count)") + $(If ($($computers.count) -gt 1){" machines"}Else{" machine"})
    Write-host $msg
    
    Foreach ($computer in $computers)
    {
        $msg = "Checking Computer $Computer"
        Write-host $msg -BackgroundColor yellow -ForegroundColor Blue
        Try
        {
            $LastEvent = Get-WinEvent -Credential $GetAdminact -ComputerName $Computer -Logname 'Application' -oldest -MaxEvents 1
            Write-host "Event logs on $computer goes as far as $($LastEvent.TimeCreated)"
            Try
            {
                $Events = Get-WinEvent -Credential $GetAdminact -FilterHashtable $FilterHashProperties -MaxEvents $NumberOfLastEventsToGet -Computer $Computer -ErrorAction stop | select MachineName, LogName, TimeCreated, LevelDisplayName, ProviderName, ID, Message
                Write-host "Found at least $($Events.count) events ! Here are the $NumberOfLastEventsToGet last ones :"
                $Events | Select -first $NumberOfLastEventsToGet
                $Events | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append 
            }
            Catch
            {
                Write-Host "No such events with EventID = $($FilterHashProperties.ID) in the $($FilterHashProperties.LogName) event log on this computer..." -ForegroundColor Green
            }
            Finally
            {
                Write-Host "OK_"
            }
        }
        Catch
        {
            $msg = "Error accessing Event Logs of $computer + $PSItem.Exception.InnerExceptionMessage"
            Write-Host $msg -ForegroundColor Red
        }
    }
    
    $msg = "Found $($Events4all.count) Events in total ..."
    Write-host $msg -BackgroundColor blue -ForegroundColor yellow    
}


Foreach ($i in $target)
{
    if ($i -eq "Logon" -or $i -eq "All") {
        Write-Host "EventID: " $id_logon
        ExportFor($id_logon)                
    }

    if ($i -eq "Service" -or $i -eq "All") {
        Write-Host "EventID: " $id_service
        ExportFor($id_service)
    }

    if ($i -eq "User" -or $i -eq "All") {
        Write-Host "EventID: " $id_usermagement
        ExportFor($id_usermagement)
    }

    if ($i -eq "Computer" -or $i -eq "All") {
        Write-Host "EventID: " $id_computermagement
        ExportFor($id_computermagement)
    }

    if ($i -eq "Clean" -or $i -eq "All") {
        
        Write-Host "EventID: " $id_clean
        ExportFor($id_clean)
    }

    if ($i -eq "File" -or $i -eq "All") {
        Write-Host "EventID: " $id_file
        ExportFor($id_file)
    }

    if ($i -eq "MSSQL" -or $i -eq "All") {
        Write-Host "EventID: " $id_mssql
        ExportFor($id_mssql)
    }

    if ($i -eq "RAS" -or $i -eq "All") {
        Write-Host "EventID: " $id_ras
        ExportFor($id_ras)
    }

    if ($i -eq "USB" -or $i -eq "All") {
        Write-Host "EventID: " $id_usb
        ExportFor($id_usb)
    }    
}



$msg = "I'm done. I exported the results to a file located on the same directory as the Script"
Write-host $msg

<# /EXECUTIONS #>
<# ---------------------------- SCRIPT_FOOTER ---------------------------- #>
#Stopping StopWatch and report total elapsed time (TotalSeconds, TotalMilliseconds, TotalMinutes, etc...)
$stopwatch.Stop()
$msg = "The script took $($StopWatch.Elapsed.TotalSeconds) seconds to execute..."
Write-Host $msg

<# ---------------- /SCRIPT_FOOTER (NOTHING BEYOND THIS POINT) ----------- #>
