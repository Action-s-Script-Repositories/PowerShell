#requires -module ImportExcel
If (Get-Module ImportExcel){
 Write-Host -ForegroundColor 'Green' "Module is allready installed."
}
Else {
    Try{
    Write-Host -ForegroundColor 'Green' "Start instsalling ImportExcel module." -NoNewline
    Install-Module -Name ImportExcel -Scope AllUsers -Force
     Write-Host -ForegroundColor 'Green' "Module is installed." -NoNewline
     }
     Catch [Exception] {
        Write-Host -ForegroundColor 'Red' "[ERROR] " -NoNewline
        Write-Host "Failed to install the module"
        Write-Host -ForegroundColor 'Red' $_ -ErrorAction Stop    
        return $false
        }
 }

#Function to set the NTFS rights and to remove specific rights
function Set-NTFS {
    param(
        [parameter(Mandatory = $true)][string]$Folder,
        [parameter(Mandatory = $true)][string]$Identity,
        [parameter(Mandatory = $true)][string]$AccessRights
    )

    switch ($AccessRights)
    {
        "L" {$rights = 'ReadAndExecute';break}
        "R" {$rights = 'ReadAndExecute';break}
        "M" {$rights = 'Modify';break}
        "F" {$rights = 'FullControl';break}
    }

    #Remove Inheritance
    $Acl = Get-ACL -path $Folder
    $Acl.SetAccessRuleProtection($true,$true)
    Set-Acl -Path $Folder -AclObject $Acl 

     
 

    #Remove all the Nivel Specific ACL settings from the path
    $Acl = Get-ACL -path $Folder
    $CustomerPermissions = $Acl.Access | Where {$_.IdentityReference -like "$($Domain)\035_NTFS*"} 
    Foreach ($CustomerPermission in $CustomerPermissions){
         Try{
            $Acl.RemoveAccessRuleSpecific($CustomerPermission)
            Set-Acl -Path $Folder -AclObject $Acl 
            Write-Host -ForegroundColor 'Green' "[SUCCESS] " -NoNewline
            Write-Host "NTFS permissions removed"
         }
        Catch{
            Write-Host -ForegroundColor 'Red' "[ERROR] " -NoNewline
            Write-Host "Failed to remove NTFS permissions"
            Write-Host -ForegroundColor 'Red' $_
         }
    }


    $Inheritance = 'None'
    $Propagation = 'None'
    $Type = 'Allow'

    $ACE = New-Object System.Security.AccessControl.FileSystemAccessRule($Identity,$Rights,$Inheritance,$Propagation,$Type)

    $Acl = Get-Acl -Path $Folder
    
    try {
        $Acl.AddAccessRule($ACE)
    } catch [EXCEPTION] {
        Write-Host -ForegroundColor 'Red' "[ERROR] " -NoNewline
        Write-Host "Failed to set NTFS permissions"
        Write-Host -ForegroundColor 'Red' $_
        return $false

    }



    try {
        Set-Acl -Path $Folder -AclObject $Acl -ErrorAction Stop 
        Write-Host -ForegroundColor 'Green' "[SUCCESS] " -NoNewline
        Write-Host "NTFS permissions set"
    } catch [EXCEPTION] {
        Write-Host -ForegroundColor 'Red' "[ERROR] " -NoNewline
        Write-Host "Failed to set NTFS permissions"
        Write-Host -ForegroundColor 'Red' $_
    }
}

#Variable for the NETBIOS domain name
$Domain = "NIVEL"

$RunDirectory = $PSScriptRoot # If file is run as ps1, this will be populated with the script directory
if([string]::IsNullOrEmpty($runDirectory)) { $runDirectory = Get-Location | Select-Object -ExpandProperty Path } # If $runDirectory is empty, code snippets probably run, so variable will be populated throug Get-Location (make sure your terminal is in the directory where the script file is located)

$PermissionsFile = "NTFS_Permissions_TEST.xlsx" # Inputfile variable
Write-host -ForegroundColor 'Green' "$($PermissionsFile) is the file being used"

$DataFromExcel = Import-Excel -Path "$($RunDirectory)\$($PermissionsFile)" # Import data from inputfile
Write-host -ForegroundColor 'Green' "$($RunDirectory)\$($PermissionsFile) is the file being used"

foreach($DataLine in $DataFromExcel) { # Loop through all lines from inputfile
    $DataLine.PSObject.Properties | ForEach-Object { # Loop through all object properties of the current line
        if(!($_.Name -eq "Foldernaam")) { # Skip the Foldernaam property as we do not need this here
            if(!([string]::IsNullOrEmpty($_.Value))) { # If the property value is not null, then proceed
                Set-NTFS -folder $dataLine.Foldernaam -identity "$($DOMAIN)\$($_.Name)" -accessRights $_.Value # Call Set-NTFS function with proper arguments ($_.Name is property name as in AD group, $_.Value is property value as in AccessRights)
            }
        }
    }
}