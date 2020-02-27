function Win32Restart-Computer {
<#
.SYNOPSIS
Restarts one or more computers using the WMI Win32_OperatingSystem method.
.DESCRIPTION
Restarts, shuts down, logs off or powers down one or more computers. This relies on WMI's Win32_OperatingSystem class. Supports common parameters -verbose, -whatif, and -confirm
.PARAMETER ComputerName
One or more computer names to operate against. Accepts pipeline input ByValue and ByPropertyName
.PARAMETER action
Can be Restart, LogOff, Shutdown, or PowerOff
.PARAMETER force
$true or $false to force the action; defaults to $false
#>

    [CmdletBinding(
        SupportsShouldProcess=$true,
        ConfirmImpact="High"
    )]
    param(
        [parameter(
            Mandatory=$true,
            ValueFromPipeline=$true,
            ValueFromPipelineByPropertyName=$true
        )]
        [string[]]$ComputerName,

        [parameter(
            Mandatory=$true
        )]
        [string][validateset("Restart","LogOff","Shutdown","PowerOff")]$action,

        [boolean]$force = $false
    )

    BEGIN {
        #translate action to numeric value required by the method
        switch ($action) {
            "Restart" {
                $_action = 2
                break
            }
            "LogOff" {
                $_action = 0
                break
            }
            "Shutdown" {
                $_action = 1
            }
            "PowerOff" {
                $_action = 8
            }
        }

        # to force, add 4 to the value
        if ($force) {
            $_action += 4
        }

        Write-Verbose "Action set to $action"
    }

    PROCESS {
        Write-Verbose "Attempting to connect to $($ComputerName)"

        # this is how we support -whatif and -confirm
        # which are enabled by the SupportsShouldProcess
        # parameter int the cmdlet binding
        if ($PSCmdlet.ShouldProcess($ComputerName)) {
            Get-WmiObject -Class Win32_OperatingSystem -ComputerName $ComputerName | `
                Invoke-WmiMethod -Name Win32Shutdown -Argumentlist $_action
        }
    }

    END {

    }
}

'localhost' | Win32Restart-Computer -action LogOff -WhatIf