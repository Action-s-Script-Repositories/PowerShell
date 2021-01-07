

$runDirectory = $PSScriptRoot # If file is run as ps1, this will be populated with the script directory
if([string]::IsNullOrEmpty($runDirectory)) { $runDirectory = Get-Location | Select-Object -ExpandProperty Path } # If $runDirectory is empty, code snippets probably run, so variable will be populated throug Get-Location (make sure your terminal is in the directory where the script file is located)

$UsersWithForwardEnabled = Get-Mailbox -Resultsize Unlimited | Select-Object DisplayName,PrimarySMTPAddress,ForwardingAddress | Where-Object {$_.ForwardingAddress -ne $null} | Sort-Object DisplayName

$processedObjects = @()

$UsersWithForwardEnabled | ForEach-Object {
    $ForwardingAddressInfo = $null
    
    $ForwardingAddressInfo = Get-Contact $_.ForwardingAddress | Select-Object WindowsEmailAddress
    
    if($null -eq $ForwardingAddressInfo) {
        $Name = ($_.ForwardingAddress).Split('/')[3]
        $ForwardingAddressInfo = Get-ADUser -Filter {Name -like $Name} -Property mail | Select-Object mail
        if($null -eq $ForwardingAddressInfo) {
            $EmailAddress = "N/A"
        } else {
            $EmailAddress = $ForwardingAddressInfo.mail
        }
    } else {
        $EmailAddress = $ForwardingAddressInfo.WindowsEmailAddress
    }

    $properties = @{Name = $_.DisplayName
        PrimarySMTPAddress = $_.PrimarySMTPAddress
        ForwardingContact = $_.ForwardingAddress
        ForwardingMailAddress = $EmailAddress
        }
    $processedObjects += New-Object -TypeName PSObject -Property $Properties | Select-Object Name,PrimarySMTPAddress,ForwardingContact,ForwardingMailAddress
}

$processedObjects | Export-Excel -Path "$($runDirectory)\UsersWithForwardEnabled.xlsx" -WorksheetName "UsersWithForwardEnabled" -Title "Users with Forward enabled" -TitleBold -AutoSize -AutoFilter