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





# Define date variables

$campDate   = (Get-Date).AddDays(-180)
$workDate   = Get-Date

# Define server variables

$camp       = (Get-WJPath -Name camp).Path
$public     = (Get-WJPath -Name public).Path
$public_adphoto = (Get-WJPath -Name public_adphoto).Path
$public_adtext  = (Get-WJPath -Name public_adtext).Path
$workPath   = ($camp + $workDate.ToString("yyyyMMdd") + "\")

# Define arrays

$pathList = @( $public_adphoto, $public_adtext )

# Log variables

Write-Log -Verb "campDate" -Noun $campDate -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

Write-Log -Verb "camp" -Noun $camp -Path $log -Type Short -Status Normal
Write-Log -Verb "public" -Noun $public -Path $log -Type Short -Status Normal
Write-Log -Verb "public_adphoto" -Noun $public_adphoto -Path $log -Type Short -Status Normal
Write-Log -Verb "public_adtext" -Noun $public_adtext -Path $log -Type Short -Status Normal
Write-Log -Verb "workPath" -Noun $workPath -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log

$pathList | ForEach-Object{ Write-Log -Verb "pathList" -Noun $_ -Path $log -Type Short -Status Normal }
Write-Line -Length 50 -Path $log





Write-Log -Verb "START" -Noun "pathList" -Path $log -Type Long -Status Normal
Write-Line -Length 50 -Path $log

foreach($path in $pathList){




    
    # 1 Rename files that contains brackets

    Write-Log -Verb "REMOVE BRACKET" -Noun $path -Path $log -Type Long -Status System

    Get-ChildItem $path -Recurse | Sort-Object FullName -Descending | ForEach-Object { 
        if($_.BaseName -match "\["){
            Write-Log -Verb "HAS BRACKET" -Noun $_.FullName -Path $log -Type Long -Status Normal
            $newName = ($_.BaseName.Replace("[","(")).Replace("]",")") + "_r" + $_.Extension
            $newParent  = Split-Path $_.FullName
            $newFullName = Join-Path -Path $newParent -ChildPath $newName
            Write-Log -Verb "newName" -Noun $newName -Path $log -Type Short -Status Normal
            Write-Log -Verb "newFullName" -Noun $newFullName -Path $log -Type Short -Status Normal
            try{
                Rename-Item -LiteralPath $_.FullName -NewName $newFullName
                Write-Log -Verb "RENAME" -Noun $newFullName -Path $log -Type Long -Status Good
            }catch{
                $mailMsg = $mailMsg + (Write-Log -Verb "RENAME" -Noun $newFullName -Path $log -Type Long -Status Bad -Output String) + "`n"
                $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status Bad -Output String) + "`n"
                $hasError = $true
            }
        }
    }

    Write-Line -Length 50 -Path $log
    




    # 2 Move files

    Write-Log -Verb "MOVE FILES" -Noun $path -Path $log -Type Long -Status System

    Get-ChildItem $path -Recurse | Where-Object { 
        -not $_.PSIsContainer -and ($_.CreationTime -lt $campDate)
    } | Sort-Object FullName -Descending | ForEach-Object{
        $newParent = ($_.DirectoryName).Replace($public,$workPath)
        if(!(Test-Path $newParent)){
            New-Item $newParent -ItemType Directory | Out-Null
        }
        $newFullName = Join-Path -Path $newParent -ChildPath $_.Name

        try{
            Move-Item $_.FullName $newFullName -ErrorAction Stop
            Write-Log -Verb "MOVE" -Noun $newFullName -Path $log -Type Long -Status Good
        }catch{
            $mailMsg = $mailMsg + (Write-Log -Verb "MOVE" -Noun $newFullName -Path $log -Type Long -Status Bad -Output String) + "`n"
            $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception -Path $log -Type Short -Status Bad -Output String) + "`n"
            $hasError = $true
        }
    }

    Write-Line -Length 50 -Path $log





    # 3 Delete empty folders

    Write-Log -Verb "REMOVE EMPTY FOLDERS" -Noun $path -Path $log -Type Long -Status System

    Get-ChildItem $path -Recurse | Where-Object { 
         $_.PSIsContainer 
    } | Sort-Object FullName -Descending | ForEach-Object{
        if((Get-ChildItem $_.FullName).Count -eq 0){
            try{
                $temp = $_.FullName
                Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop
                Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
            }catch{
                $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
                $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
                $hasError = $true
            }
        }
    }

    Write-Line -Length 50 -Path $log
    




}

Write-Log -Verb "END" -Noun "pathList" -Path $log -Type Long -Status Normal





##################################################################################

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