param (
    [string]$folder = "z:\test\Files_DB",
    [string]$outfilename = "explore-folder",
    [switch]$force = $false,
    [switch]$extruct = $false
 )
 



$LogDate = get-date -f yyyyMMddhhmm 
$outfile = "$($outfilename)_$LogDate.json"

Write-Host "base: " $folder
Write-Host "outfile: " $outfile

if (Test-Path $outfile) 
{
  Remove-Item $outfile
}

Function Get-MKVS-FileHash([String] $FileName,$HashName = "SHA1") 
{
  $FileStream = New-Object System.IO.FileStream($FileName,"Open", "Read") 
  $StringBuilder = New-Object System.Text.StringBuilder 
  [System.Security.Cryptography.HashAlgorithm]::Create($HashName).ComputeHash($FileStream)|%{[Void]$StringBuilder.Append($_.ToString("x2"))} 
  $FileStream.Close() 
  $StringBuilder.ToString() 
}

function Get-MKVS-DocText([String] $FileName) {
    $Word = New-Object -ComObject Word.Application
    $Word.Visible = $false
    $Word.DisplayAlerts = 0
    $text = ""
    Try
    {
        $Document = $Word.Documents.Open($FileName)
       
        $Document.Paragraphs | ForEach-Object {
            $text += $_.Range.Text
        }
        $Document.Close()
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
        Remove-Variable Word        
        return $text
    }
    Catch {
        Write-Host $PSItem.Exception.Message
        $Document.Close()
        $Word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word)
        Remove-Variable Word
    }
    
    return $text
}

function Get-MKVS-XlsText([String] $FileName) {
    $excel = New-Object -ComObject excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = 0
    $text = ""
    Try    
    {
        $wb = $excel.workbooks.open($FileName)
        foreach ($sh in $wb.Worksheets) {
            #Write-Host "sheet: " $sh.Name            
            $endRow = $sh.UsedRange.SpecialCells(11).Row
            $endCol = $sh.UsedRange.SpecialCells(11).Column
            Write-Host "dim: " $endRow $endCol
            for ($r = 1; $r -le $endRow; $r++) {
                for ($c = 1; $c -le $endCol; $c++) {
                    $t = $sh.Cells($r, $c).Text
                    $text += $t
                    #Write-Host "text cel: " $r $c $t
                }
            }
        }
    }
    Catch {
        Write-Host $PSItem.Exception.Message
    }
    #Write-Host "text: " $text
    $excel.Workbooks.Close()
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
    Remove-Variable excel
    return $text
}

function Get-MKVS-FileText([String] $FileName, [String] $Extension) {
    Write-Host "filename: " $FileName
    Write-Host "ext: " $Extension

    switch ($Extension) {
        ".doc" {
            return Get-MKVS-DocText $FileName
        }
        ".docx" {
            return Get-MKVS-DocText $FileName
        }
        ".xls" {
            return Get-MKVS-XlsText $FileName
        }
        ".xlsx" {
            return Get-MKVS-XlsText $FileName
        }
    }
    return ""    
}

#powershell.exe -ExecutionPolicy Bypass -Command "./explore_share.ps1" -outfilename explore_folder

#[System.Reflection.Assembly]::LoadFrom('d:/t2/research/Interop.Dsofile.dll')
#function Get-Summary([string]$file) {
#    $oled = New-Object -COM DSOFile.OleDocumentProperties
#    $oled.Open($file, $true, [DSOFile.dsoFileOpenOptions]::dsoOptionDefault)
#    $spd =  $oled.SummaryProperties
    #return $spd
#    $oled.close()
#    return $spd # return here instead
#}

Get-ChildItem $folder -Recurse | 
Foreach-Object {
    $cur = $_ | Select-Object -Property "Name", "FullName", "BaseName", "CreationTime", "LastAccessTime", "LastWriteTime", "Attributes", "PSIsContainer", "Extension", "Mode", "Length"
    Write-Host $_.FullName
    $acl = Get-Acl $_.FullName | Select-Object -Property "Owner", "Group", "AccessToString", "Sddl"
    $path = $_.FullName
    $ext = $_.Extension
    
    
    Try
    {
        $hash = Get-MKVS-FileHash $path
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

    if ($extruct -eq $true)
    {
        Try
        {
            $text =  Get-MKVS-FileText $path $ext
            $cur | Add-Member -MemberType NoteProperty -Name Text -Value $text -Force
        }
        Catch {
            Write-Host $PSItem.Exception.Message       
        }    
    }
    
    $cur | Add-Member -MemberType NoteProperty -Name ACL -Value $acl -Force
    $cur | Add-Member -MemberType NoteProperty -Name Hash -Value $hash -Force
    Try
    {
        #$summary = Get-Summary $path
        #$cur | Add-Member -MemberType NoteProperty -Name Summary -Value $summary -Force
    }
    Catch {
        Write-Host $PSItem.Exception.Message       
    }

    $cur | ConvertTo-Json | Out-File -FilePath $outfile -Encoding UTF8 -Append    
} 
