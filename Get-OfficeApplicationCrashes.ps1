
# Quick Scrappy Script to Grep the Windows Event Log for Office Application Crashes

# Based on the crash generating a "1000" ID in the Applicaiton Event Log

$events = Get-WinEvent -FilterHashtable @{logname='application'; id=1000} | Where-Object {$_.Message -match "Faulting application"}

cls

$Officeevents = $events | Where-Object {$_.Message -match "16.0."}
Write-host ""
Write-Host "Office 16.x Applicaton Crashes"

$Officeevents | select-object TimeCreated,Id,Message | Sort-Object Message | Format-Table

$FirstOfficeEvent = $Officeevents | sort-object TimeCreated | Select-Object -First 1
$LastOfficeEvent = $Officeevents | sort-object TimeCreated | Select-Object -Last 1


$days = ($LastOfficeEvent.TimeCreated - $FirstOfficeEvent.TimeCreated).Days

Write-Host ""
Write-Host "Total Number of all Office Crashes $($officeevents.count) over $days days"
Write-Host ""

### Get office version

##################### Office 2016 ################################################

# Office 2016 x86 C2R

$2016CR2x86Test = Test-Path ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\lync.exe")

IF ($2016CR2x86Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\MSOSB.DLL")
$SfBUCCAPI = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\root\Office16\UccApi.dll")
$InstallType = "Office 2016 x86 C2R"
}

# Office 2016 x86 MSI


$2016MSIx86Test = Test-Path ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\lync.exe")

IF ($2016MSIx86Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\MSOSB.DLL")
$SfBUCCAPI = Get-Item ("${Env:ProgramFiles(x86)}" + "\Microsoft Office\Office16\UccApi.dll")
$InstallType = "Office 2016 x86 MSI"
}

# Office 2016 x64 C2R

$2016CR2x64Test = Test-Path ("${Env:ProgramFiles}" + "\Microsoft Office\root\Office16\lync.exe")

IF ($2016CR2x64Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office\root\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office\root\Office16\MSOSB.DLL")
$SfBUCCAPI = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office\root\Office16\UccApi.dll")
$InstallType = "Office 2016 x64 C2R"
}

# Office 2016 x64 MSI


$2016MSIx64Test = Test-Path ("${Env:ProgramFiles}" + "\Microsoft Office\Office16\lync.exe")

IF ($2016MSIx64Test -eq $true)
{
$SFBEXE = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office\Office16\lync.exe")
$SfBMSO = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office\Office16\MSOSB.DLL")
$SfBUCCAPI = Get-Item ("${Env:ProgramFiles}" + "\Microsoft Office\Office16\UccApi.dll")
$InstallType = "Office 2016 x64 MSI"
}



# Output

$sfbexeVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($SFBEXE)


Write-Host " "
Write-Host "Office Install Type:         $InstallType"
Write-Host "Office Current Version:      $($sfbexeVersion.fileversion)"
Write-Host " "

