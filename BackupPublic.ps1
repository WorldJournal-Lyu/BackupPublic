<#

BackupPublic.ps1

    2018-08-23 Initial Creation

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Get-Module -ListAvailable WorldJournal.* | Remove-Module -Force
Get-Module -ListAvailable WorldJournal.* | Import-Module -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = (Get-WJEmail -Name lyu).MailAddress
$mailSbj    = $scriptName
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 50 -Path $log

###################################################################################





$newspage = (Get-WJPath -Name newspage).Path
$mailMsg = $mailMsg + (Write-Log -Verb "NEWSPAGE" -Noun $newspage -Path $log -Type Long -Status Normal -Output String) + "`n"

# Flag hasError 

if( $false ){
    $hasError = $true
}



###################################################################################

# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam