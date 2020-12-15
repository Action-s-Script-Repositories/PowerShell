#requires -module ImportExcel

function Set-NTFS {
    param(
        [parameter(Mandatory = $true)][string]$folder,
        [parameter(Mandatory = $true)][string]$identity,
        [parameter(Mandatory = $true)][string]$accessRights
    )

    switch ($accessRights)
    {
        "R" {$rights = 'ReadAndExecute';break}
        "M" {$rights = 'Modify';break}
        "F" {$rights = 'FullControl';break}
    }

    $inheritance = 'ContainerInherit, ObjectInherit'
    $propagation = 'None'
    $type = 'Allow'

    $ACE = New-Object System.Security.AccessControl.FileSystemAccessRule($identity,$rights,$inheritance,$propagation, $type)

    $Acl = Get-Acl -Path $folder
    $Acl.AddAccessRule($ACE)

    try {
        Set-Acl -Path $folder -AclObject $Acl -WhatIf
        Write-Host -ForegroundColor 'Green' "[SUCCESS] " -NoNewline
        Write-Host "NTFS permissions set"
    } catch [EXCEPTION] {
        Write-Host -ForegroundColor 'Red' "[ERROR] " -NoNewline
        Write-Host "Failed to set NTFS permissions"
        Write-Host -ForegroundColor 'Red' $_
    }
}

$runDirectory = $PSScriptRoot # If file is run as ps1, this will be populated with the script directory
if([string]::IsNullOrEmpty($runDirectory)) { $runDirectory = Get-Location | Select-Object -ExpandProperty Path } # If $runDirectory is empty, code snippets probably run, so variable will be populated throug Get-Location (make sure your terminal is in the directory where the script file is located)
$permissionsFile = "NTFS_Permissions_TEST.xlsx" # Inputfile variable

$DataFromExcel = Import-Excel -Path "$($runDirectory)\$($permissionsFile)" # Import data from inputfile

foreach($dataLine in $DataFromExcel) { # Loop through all lines from inputfile
    $dataLine.PSObject.Properties | ForEach-Object { # Loop through all object properties of the current line
        if(!($_.Name -eq "Foldernaam")) { # Skip the Foldernaam property as we do not need this here
            if(!([string]::IsNullOrEmpty($_.Value))) { # If the property value is not null, then proceed
                Set-NTFS -folder $dataLine.Foldernaam -identity $_.Name -accessRights $_.Value # Call Set-NTFS function with proper arguments ($_.Name is property name as in AD group, $_.Value is property value as in AccessRights)
            }
        }
    }
}