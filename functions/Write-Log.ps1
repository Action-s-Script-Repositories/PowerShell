<# 
.Synopsis
    Write-Log writes a message to a specified log file with the current time stamp. 
.DESCRIPTION
    The Write-Log function is designed to add logging capability to other scripts. 
    In addition to writing output and/or verbose you can write to a log file for 
    later debugging. 
.NOTES 
    Created by: Henk van Malsen 
    Modified: 04-02-2020 18:42:00   
    
    Changelog: 
        * Code simplification and clarification
        * Added documentation. 
        * Renamed LogPath parameter to Path to keep it standard
        * Revised the Force switch to work as it should
.PARAMETER Message 
    Message is the content that you wish to add to the log file.  
.PARAMETER Path 
    The path to the log file to which you would like to write. By default the function will  
    create the path and file if it does not exist.  
.PARAMETER Level 
    Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational) 
.PARAMETER NoClobber 
    Use NoClobber if you do not wish to overwrite an existing file.
.PARAMETER Vebose
    Use Verbose if you want the log message to be shown in the console as well
.EXAMPLE 
    Write-Log -Message 'Log message'
    Writes the message to $ENV:temp\PowerShellLog.log. 
.EXAMPLE 
    Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log 
    Writes the content to the specified log file and creates the path and file specified.  
.EXAMPLE 
    Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error 
    Writes the message to the specified log file as an error message, and writes the message to the error pipeline.
.EXAMPLE
    Write-Log -Message 'Log message verbose' -Verbose
    Writes the message to $ENV:temp\PowerShellLog.log and pipes output to console
.LINK 

#> 
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