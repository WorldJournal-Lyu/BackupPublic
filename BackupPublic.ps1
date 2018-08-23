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





$public = (Get-WJPath -Name public).Path
$public_adphoto = (Get-WJPath -Name public_adphoto).Path
$public_adtext = (Get-WJPath -Name public_adtext).Path
$camp = (Get-WJPath -Name camp).Path
$campDate = (Get-Date).AddDays(-180)
$workDate = Get-Date
$workPath = ($camp + $workDate.ToString("yyyyMMdd") + "\")

Write-Log -Verb "public" -Noun $public -Path $log -Type Short -Status Normal
Write-Log -Verb "public_adphoto" -Noun $public_adphoto -Path $log -Type Short -Status Normal
Write-Log -Verb "public_adtext" -Noun $public_adtext -Path $log -Type Short -Status Normal
Write-Log -Verb "camp" -Noun $camp -Path $log -Type Short -Status Normal
Write-Log -Verb "campDate" -Noun $campDate -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal
Write-Log -Verb "workPath" -Noun $workPath -Path $log -Type Short -Status Normal

Write-Line -Length 50 -Path $log




# Move ADPHOTO files

Write-Log -Verb "MOVE FILES" -Noun $public_adphoto -Path $log -Type Long -Status Normal

Get-ChildItemPlus $public_adphoto | Where-Object { 

    -not $_.PSIsContainer -and ($_.CreationTime -lt $campDate)

} | Sort-Object -Descending | Move-Files -From $public -To $workPath | ForEach-Object{

    Write-Log -Verb "moveFrom" -Noun $_.MoveFrom -Path $log -Type Short -Status Normal
    Write-Log -Verb "moveTo" -Noun $_.MoveTo -Path $log -Type Short -Status Normal
    Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status

    if($_.Status -eq "Bad"){

        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status -Output String) + "`n"
        $hasError = $true

    }

}

Write-Line -Length 50 -Path $log




# Delete empty folders in ADPHOTO

Write-Log -Verb "REMOVE EMPTY FOLDERS" -Noun $public_adphoto -Path $log -Type Long -Status Normal

Get-ChildItemPlus $public_adphoto | Where-Object { 

     $_.PSIsContainer 

} | Sort-Object -Descending | ForEach-Object{

    if((Get-ChildItem $_.FullName).Count -eq 0){

        try{

            $temp = $_.FullName
            Remove-Item $_ -Recurse -Force -ErrorAction Stop
            Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good

        }catch{

            $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
            $hasError = $true

        }

    }

}

Write-Line -Length 50 -Path $log




# Move ADTEXT files

Write-Log -Verb "MOVE FILES" -Noun $public_adtext -Path $log -Type Long -Status Normal

Get-ChildItemPlus $public_adtext | Where-Object { 

    -not $_.PSIsContainer -and ($_.CreationTime -lt $campDate)

} | Sort-Object -Descending | Move-Files -From $public -To $workPath | ForEach-Object{

    Write-Log -Verb "moveFrom" -Noun $_.MoveFrom -Path $log -Type Short -Status Normal
    Write-Log -Verb "moveTo" -Noun $_.MoveTo -Path $log -Type Short -Status Normal
    Write-Log -Verb $_.Verb -Noun $_.Noun -Path $log -Type Long -Status $_.Status

    if($_.Status -eq "Bad"){

        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status $_.Status -Output String) + "`n"
        $hasError = $true

    }

}

Write-Line -Length 50 -Path $log



# Delete empty folders in ADTEXT

Write-Log -Verb "REMOVE EMPTY FOLDERS" -Noun $public_adtext -Path $log -Type Long -Status Normal

Get-ChildItemPlus $public_adtext | Where-Object { 

     $_.PSIsContainer

} | Sort-Object -Descending | ForEach-Object{

    if((Get-ChildItem $_.FullName).Count -eq 0){

        try{

            $temp = $_.FullName
            Remove-Item $_ -Recurse -Force -ErrorAction Stop
            Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good

        }catch{

            $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
            $hasError = $true

        }

    }

}





###################################################################################

Write-Line -Length 50 -Path $log

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