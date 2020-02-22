<#
.SYNOPSIS
	Install software remotely in a group of computers and retry the installation in case of error.
.DESCRIPTION
	This script install software remotely in a group of computers and retry the installation in case of error.
    It uses PowerShell to perform the installation. Target computer must allow Windows PowerShell Remoting.
.PARAMETER AppPath
	Path to the application executable, It can be a network or local path because entire folder will be copied to remote computer before installing and deleted after installation. 
    Example: 'C:\Software\TeamViewer\TeamvieverHost.msi' (Folder TeamViewer will be copied to remote computer before run ejecutable)
.PARAMETER AppArgs
    Application arguments to perform silent installation.
    Example: '/S /R settings.reg'
.PARAMETER LocalPath
    Local path of the remote computer where copy application directory.
    Default: 'C:\temp'
.PARAMETER Retries
    Number of times to retry failed installations.
    Default: 5
.PARAMETER TimeBetweenRetries
    Seconds to wait before retrying failed installations.
    Default: 60
.PARAMETER ComputerList
    List of computers in install software. You can only use one source of target computers: ComputerList, OU or CSV.
    Example: Computer001,Computer002,Computer003 (Without quotation marks)
.PARAMETER LogPath
    Path where save log file.
    Default: My Documents
.PARAMETER WMIQuery
    WMI Query to execute in remote computers. Software will be installed if query returns values.
    Example: 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="64"' (64 bit computers)
    Example: 'select * from Win32_Processor where DeviceID="CPU0" and AddressWidth="32"' (32 bit computers)
    Default: None
.EXAMPLE
    .EXAMPLE
    .\InstallSoftwareRemotely.ps1 -AppPath 'C:\Path\to\exe' -AppArgs '/verysilent /norestart'
.EXAMPLE
	.\InstallSoftwareRemotely.ps1 -AppPath "\\UNC\Path\to\exe" -AppArgs "/S" -ComputerList Computer001,Computer002,Computer003
.NOTES 
	Author: Juan Granados 
	Date:   November 2017
#>

Param(
    [Parameter(Mandatory = $true, Position = 0)] 
    [ValidateNotNullOrEmpty()]
    [string]$AppPath,
    [Parameter(Mandatory = $false, Position = 1)] 
    [ValidateNotNullOrEmpty()]
    [string]$AppArgs = "None",
    [Parameter(Mandatory = $false, Position = 2)] 
    [ValidateNotNullOrEmpty()]
    [string]$LocalPath = "C:\temp",
    [Parameter(Mandatory = $false, Position = 3)] 
    [ValidateNotNullOrEmpty()]
    [int]$Retries = 5,
    [Parameter(Mandatory = $false, Position = 4)] 
    [ValidateNotNullOrEmpty()]
    [int]$TimeBetweenRetries = 60,
    [Parameter(Mandatory = $false, Position = 5)] 
    [ValidateNotNullOrEmpty()]
    [string[]]$ComputerList,
    [Parameter(Mandatory = $false, Position = 6)] 
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\temp"
)

#Requires -RunAsAdministrator

#Functions

Add-Type -AssemblyName System.IO.Compression.FileSystem

Function Copy-WithProgress {
    Param([string]$Source, [string]$Destination)

    $Source = $Source.tolower()
    $Filelist = Get-Childitem $Source –Recurse
    $Total = $Filelist.count
    $Position = 0
    If (!(Test-Path $Destination)) {
        New-Item $Destination -Type Directory | Out-Null
    }
    foreach ($File in $Filelist) {
        $Filename = $File.Fullname.tolower().replace($Source, '')
        $DestinationFile = ($Destination + $Filename)
        try {
            Copy-Item $File.FullName -Destination $DestinationFile -Force
        }
        catch { throw $_.Exception }
        $Position++
        Write-Progress -Activity "Copying data from $source to $Destination" -Status "Copying File $Filename" -PercentComplete (($Position / $Total) * 100)
    }
}

Function Set-Message([string]$Text, [string]$ForegroundColor = "White", [int]$Append = $True) {

    if ($Append) {
        $Text | Out-File $LogPath -Append
    }
    else {
        $Text | Out-File $LogPath
    }
    Write-Host $Text -ForegroundColor $ForegroundColor
}

Function InstallRemoteSoftware([string]$Computer) {
    try {
        Return Invoke-Command -computername $Computer -ScriptBlock {
            $Application = $args[0]
            $AppArgs = $args[1]
            $ApplicationName = $Application.Substring($Application.LastIndexOf('\') + 1)
            $ApplicationFolderPath = $Application.Substring(0, $Application.LastIndexOf('\'))
            $ApplicationExt = $Application.Substring($Application.LastIndexOf('.') + 1)
            Write-Host "Installing $($ApplicationName) on $($env:COMPUTERNAME)"
            If ($ApplicationExt -eq "msi") {
                If ($AppArgs -ne "None") {
                    Write-Host "Installing as MSI: msiexec /i $($Application) $($AppArgs)"
                    $p = Start-Process "msiexec" -ArgumentList "/i $($Application) $($AppArgs)" -Wait -Passthru
                }
                else {
                    Write-Host "Installing as MSI: msiexec /i $($Application)"
                    $p = Start-Process "msiexec" -ArgumentList "/i $($Application) /quiet /norestart" -Wait -Passthru
                }
            }
            ElseIf ($AppArgs -ne "None") {
                Write-Host "Executing $Application $AppArgs"
                $p = Start-Process $Application -ArgumentList $AppArgs -Wait -Passthru
            }
            Else {
                Write-Host "Executing $Application"
                $p = Start-Process $Application -Wait -Passthru
            }
            $p.WaitForExit()
            if ($p.ExitCode -ne 0) {
                Write-Host "Failed installing with error code $($p.ExitCode)" -ForegroundColor Red
                $Return = $($env:COMPUTERNAME)
            }
            else {
                $Return = 0
            }
            Write-Host "Deleting $($ApplicationFolderPath)"
            Remove-Item $($ApplicationFolderPath) -Force -Recurse
            Return $Return
        } -ArgumentList "$($LocalPath)\$($ApplicationFolderName)\$($ApplicationName)", $AppArgs
    }
    catch { throw $_.Exception }
}

$ErrorActionPreference = "Stop"

#Initialice log
if (!(test-path $LogPath)) {
    try {
        mkdir -Path $LogPath
    }
    catch {
        $($_.Exception.Message)
        Exit 1;
    }
}

$LogPath += "\InstallSoftwareRemotely_" + $(get-date -Format "yyyy-mm-dd_hh-mm-ss") + ".txt"
Set-Message "Start remote installation on $(get-date -Format "yyyy-mm-dd hh:mm:ss")" -Append $False

#Initial validations.

If (!(Test-Path $AppPath)) {
    Set-Message "Error accessing $($AppPath). The script can not continue"
    Exit 1
}
If (!$ComputerList) {
    Set-Message "You have to set a list of computers, OU or CSV." -ForegroundColor Red
    Exit 1
}

$ApplicationName = $AppPath.Substring($AppPath.LastIndexOf('\') + 1)
$ApplicationFolderPath = $AppPath.Substring(0, $AppPath.LastIndexOf('\'))
$ApplicationFolderName = $ApplicationFolderPath.Substring($ApplicationFolderPath.LastIndexOf('\') + 1)
$ComputerWithError = [System.Collections.ArrayList]@()
$ComputerWithSuccess = [System.Collections.ArrayList]@()
$ComputerSkipped = [System.Collections.ArrayList]@()
$TotalRetries = $Retries
$TotalComputers = $ComputerList.Count

Do {
    Set-Message "-----------------------------------------------------------------"
    Set-Message "Attempt $(($TotalRetries - $Retries) +1) of $($TotalRetries)" -ForegroundColor Cyan
    Set-Message "-----------------------------------------------------------------"
    $Count = 1
    ForEach ($Computer in $ComputerList) {
        Set-Message "COMPUTER $($Computer.ToUpper()) ($($Count) of $($ComputerList.Count))" -ForegroundColor Yellow
        $Count++
        Set-Message "Coping $($ApplicationFolderPath) to \\$($Computer)\$($LocalPath -replace ':','$')"
        try {
            Copy-WithProgress "$ApplicationFolderPath" "\\$($Computer)\$("$($LocalPath)\$($ApplicationFolderName)" -replace ':','$')"
        }catch {
            Set-Message "Error copying folder: $($_.Exception.Message)" -ForegroundColor Red
            $ComputerWithError.Add($Computer) | Out-Null
            Continue;
        }
        try {
            $ExitCode = InstallRemoteSoftware $Computer
            If ($ExitCode) {
                $ComputerWithError.Add($Computer) | Out-Null
                Set-Message "Error installing $($ApplicationName)." -ForegroundColor Red
            }
            else {
                Set-Message "$($ApplicationName) installed successfully." -ForegroundColor Green
                $ComputerWithSuccess.Add($Computer) | Out-Null
            }
        }catch {
            Set-Message "Error on remote execution: $($_.Exception.Message)" -ForegroundColor Red
            $ComputerWithError.Add($Computer) | Out-Null
            try {
                Set-Message "Deleting \\$($Computer)\$($LocalPath -replace ':','$')\$($ApplicationFolderName)"
            }catch {
                Set-Message "Error on remote deletion: $($_.Exception.Message)" -ForegroundColor Red
            }
            Remove-Item "\\$($Computer)\$($LocalPath -replace ':','$')\$($ApplicationFolderName)" -Force -Recurse
        }
    }
    If ($ComputerWithError.Count -eq 0) {
        break
    }
    $Retries--
    If ($Retries -gt 0) {
        $ComputerList = $ComputerWithError
        $ComputerWithError = [System.Collections.ArrayList]@()
        If ($TimeBetweenRetries -gt 0) {
            Set-Message "Waiting $($TimeBetweenRetries) seconds before next retry..."
            Sleep $TimeBetweenRetries
        }
    }
}While ($Retries -gt 0)

If ($ComputerWithError.Count -gt 0) {
    Set-Message "-----------------------------------------------------------------"
    Set-Message "Error installing $($ApplicationName) on $($ComputerWithError.Count) of $($TotalComputers) computers:"
    Set-Message $ComputerWithError
    $csvContents = @()
    ForEach ($Computer in $ComputerWithError) {
        $row = New-Object System.Object
        $row | Add-Member -MemberType NoteProperty -Name "Name" -Value $Computer
        $csvContents += $row
    }
    $CSV = (get-date).ToString('yyyyMMdd-HH_mm_ss') + "ComputerWithError.csv"
    $csvContents | Export-CSV -notype -Path "$([Environment]::GetFolderPath("MyDocuments"))\$($CSV)" -Encoding UTF8
    Set-Message "Computers with error exported to CSV file: $([Environment]::GetFolderPath("MyDocuments"))\$($CSV)" -ForegroundColor DarkYellow
    Set-Message "You can retry failed installation on this computers using parameter -CSV $([Environment]::GetFolderPath("MyDocuments"))\$($CSV)" -ForegroundColor DarkYellow
}
If ($ComputerWithSuccess.Count -gt 0) {
    Set-Message "-----------------------------------------------------------------"
    Set-Message "$([math]::Round((($ComputerWithSuccess.Count * 100) / $TotalComputers), [System.MidpointRounding]::AwayFromZero) )% Success installing $($ApplicationName) on $($ComputerWithSuccess.Count) of $($TotalComputers) computers:"
    Set-Message $ComputerWithSuccess
}
Else {
    Set-Message "-----------------------------------------------------------------"
    Set-Message "Installation of $($ApplicationName) failed on all computers" -ForegroundColor Red
}
If ($ComputerSkipped.Count -gt 0) {
    Set-Message "-----------------------------------------------------------------"
    Set-Message "$($ComputerSkipped.Count) skipped of $($TotalComputers) computers:"
    Set-Message $ComputerSkipped
}