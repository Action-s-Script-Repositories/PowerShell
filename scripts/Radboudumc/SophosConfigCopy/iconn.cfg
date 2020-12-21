param(
    [Parameter(Mandatory=$false)]
    [switch]$Uninstall
)

$RunDirectory = $PSScriptRoot # If file is run as ps1, this will be populated with the script directory
if([string]::IsNullOrEmpty($runDirectory)) { $runDirectory = Get-Location | Select-Object -ExpandProperty Path } # If $runDirectory is empty, code snippets probably run, so variable will be populated throug Get-Location (make sure your terminal is in the directory where the script file is located)

$curDate = Get-Date -format "yyyy-MM-dd HHmmss"

$LogPath = "C:\temp\"
if(!(test-path $LogPath)) { New-Item -Path $LogPath -ItemType Directory -Force -Confirm:$false }
$LogFile = "$($LogPath)\$($curDate)_SophosUpdate.log"

function Write-Log { 
    [CmdletBinding()] 
    Param ( 
        [Parameter(Mandatory=$true, 
                ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 

        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path="$($PSScriptRoot)\PowerShellLog.log", 
        
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
        
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 

    Begin { 
        ## Use -Verbose when calling the function to get verbose messages to show

        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        #$VerbosePreference = 'Continue' 

    }

    Process { 
        if ((Test-Path $Path) -AND $NoClobber) {
            # If the file already exists and NoClobber was specified, do not write to the log.
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
        } elseif (!(Test-Path $Path)) {
            # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path.  
            try {
                Write-Verbose "Creating $Path." 
                $NewLogFile = New-Item $Path -Force -ItemType File
            } catch [Exception] {
                Write-Error "Failed to create $Path due to the following error:"
                Write-Error $_
            }
        } else { 
            # Nothing to see here yet. 
        } 
        
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
        
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) {
            'Error' {
                #Write-Error $Message 
                $LevelText = 'ERROR:' 
            } 
            'Warn' { 
                #Write-Warning $Message 
                $LevelText = 'WARNING:' 
            } 
            'Info' { 
                #Write-Verbose $Message 
                $LevelText = 'INFO:' 
            } 
        } 
        
        Write-Verbose -Message "$FormattedDAte $LevelText $Message"

        if($loggingPreference -eq 'Continue') {
            if ($loggingFilePreference)
            {
                $LogFile=$loggingFilePreference
            }
            else
            {
                $LogFile=$Path
            }
            
            # Write log entry to $Path
            Write-Output "$FormattedDate $LevelText $Message" | Out-File -FilePath $LogFile -Append 
        }
    } 
    
    End {

    } 
}

$DefaultConfigPath = "C:\ProgramData\Sophos\AutoUpdate\DefaultConfig"
$ConfigPath = "C:\ProgramData\Sophos\AutoUpdate\Config"
$ConfigFile = "iconn.cfg"
$BackupFile = "iconnBackUp.cfg"

if($Uninstall.IsPresent) {
    Write-Log -Message "Uninstall detected." -Level "Info"
    try{
        Copy-Item -Path "$($DefaultConfigPath)\$($BackupFile)" -Destination "$($ConfigPath)\$ConfigFile)" -Force -Confirm:$false -ErrorAction stop
        Write-Log -Message "`"$($DefaultConfigPath)\$($BackupFile)`" restored successfully" -Level "Info"
    } catch [EXCEPTION] {
        Write-Log -Message "Failed to restore `"$($DefaultConfigPath)\$($BackupFile)`"" -Level "Error"
        Write-Log -Message "Stacktrace:"
        Write-Log -Message $_ -Level "Error"
        Exit 100
    }
} else {
    if(!(test-path $DefaultConfigPath)) {
        Write-Log -Message "`"$($DefaultConfigPath)`" not found. Creating..." -Level "Warn"
        try {
            New-Item -Path $DefaultConfigPath -ItemType Directory -Force -Confirm:$false -ErrorAction stop
            Write-Log -Message "`"$($DefaultConfigPath)`" created successfully" -Level "Info"
        } catch [EXCEPTION] {
            Write-Log -Message "Failed to create `"$($DefaultConfigPath)`"" -Level "Error"
            Write-Log -Message "Stacktrace:"
            Write-Log -Message $_ -Level "Error"
            Exit 99
        }
    }

    if(test-path "$($ConfigPath)\$($ConfigFile)") {
        Write-Log -Message "`"$($ConfigPath)\$($ConfigFile)`" found. Creating backup..." -Level "Info"
        try {
            Copy-Item -Path "$($ConfigPath)\$($ConfigFile)" -Destination "$($DefaultConfigPath)\$($BackupFile)" -Force -Confirm:$false -ErrorAction stop
            Write-Log -Message "`"$($ConfigPath)\$($ConfigFile)`" copied successfully." -Level "Info"
        } catch [EXCEPTION] {
            Write-Log -Message "Failed to ccopy `"$($ConfigPath)\$($ConfigFile)`"" -Level "Error"
            Write-Log -Message "Stacktrace:"
            Write-Log -Message $_ -Level "Error"
            Exit 98
        }
    }

    if(test-path "$($runDirectory)\$($ConfigFile)") {
        Write-Log -Message "`"$($runDirectory)\$($ConfigFile)`" found. Copying correct config..." -Level "Info"
        try {
            Copy-Item -Path "$($runDirectory)\$($ConfigFile)" -Destination "$($ConfigPath)\$($BackupFile)" -Force -Confirm:$false
            Write-Log -Message "`"$($runDirectory)\$($ConfigFile)`" copied successfully." -Level "Info"
        } catch [EXCEPTION] {
            Write-Log -Message "Failed to ccopy `"$($runDirectory)\$($ConfigFile)`"" -Level "Error"
            Write-Log -Message "Stacktrace:"
            Write-Log -Message $_ -Level "Error"
            Exit 97
        }
    }
}

Exit 0