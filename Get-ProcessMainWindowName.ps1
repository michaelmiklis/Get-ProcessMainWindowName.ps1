######################################################################
## (C) 2018 Michael Miklis (michaelmiklis.de)
##
##
## Filename:      Get-ProcessMainWindowName.ps1
##
## Version:       1.0
##
## Release:       Final
##
## Requirements:  -none-
##
## Description:   Monitor a process' main window name
##
## This script is provided 'AS-IS'.  The author does not provide
## any guarantee or warranty, stated or implied.  Use at your own
## risk. You are free to reproduce, copy & modify the code, but
## please give the author credit.
##
####################################################################
Set-PSDebug -Strict
Set-StrictMode -Version latest

function Get-ProcessMainWindowName
{
    <#
        .SYNOPSIS
        Monitor a process' main window name
  
        .DESCRIPTION
        The Get-ProcessMainWindowName CMDlet writes all main window
        titles of the started process to the default output
  
        .PARAMETER ProcessName
        Process commandline
  
        .PARAMETER ProcessArguments
        Commandline arguments for process
 
        .PARAMETER TimeToMonitor
        Seconds how long this script will monitor
  
        .EXAMPLE
        Set-MSOLLicenseToADGroupMembers -GroupName "Office365_E3" -License "contoso:ENTERPRISEPACK"
    #>

    param (
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()][string]$ProcessName,
        [parameter(Mandatory=$false)]
        [string]$ProcessArguments,
        [parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()][int]$TimeToMonitor
    )

    # Start process to monitor
    $Process = [diagnostics.process]::start($ProcessName, $ProcessArguments)

    $Process.WaitForInputIdle();

    $count = 0

    do
    {
        # Get main windows title
        $WndTitle = $Process.MainWindowTitle

        # Output window title
        Write-Host $WndTitle

        # Refresh process data for next loop
        $Process.Refresh()

        # wait 100ms 
        Start-Sleep -Milliseconds 100

        $count++;
    } while ($count -le $($TimeToMonitor*10))
}

Get-ProcessMainWindowName -ProcessName "winword.exe" -TimeToMonitor 3